#!/usr/bin/env python3
"""
import_mbox.py: Import an Outlook archive (converted to mbox/eml) into a
Google Workspace Gmail mailbox, preserving folder structure as labels.

Authentication: a Google Cloud service account with domain-wide delegation,
impersonating the target mailbox. No end-user login required.

See README.md for one-time setup. Quick usage:

    # Preview only (no writes): prints the label tree and per-folder counts
    python import_mbox.py --source ./client-archive/ \
        --user jane@client.com --key service-account.json --dry-run

    # Test batch, then full run
    python import_mbox.py --source ./client-archive/ \
        --user jane@client.com --key service-account.json --max 25
    python import_mbox.py --source ./client-archive/ \
        --user jane@client.com --key service-account.json
"""

import argparse
import base64  # noqa: F401  (kept for reference; media upload used instead)
import hashlib
import io
import json
import mailbox
import os
import random
import re
import sys
import time
from datetime import datetime

# The Google client libraries are only needed for the actual import (building
# the service and inserting messages). Import them lazily so the module still
# loads for --dry-run discovery and for tests on a machine without the deps;
# the friendly "install requirements" error is raised when the service is built.
try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
    from googleapiclient.http import MediaIoBaseUpload
    _GOOGLE_IMPORT_ERROR = None
except ImportError as _import_err:
    service_account = build = MediaIoBaseUpload = None
    HttpError = Exception  # placeholder; real runs exit before this is referenced
    _GOOGLE_IMPORT_ERROR = _import_err

# Scopes must exactly match what is authorized in the Admin console
# (Security > API controls > Domain-wide delegation).
SCOPES = [
    "https://www.googleapis.com/auth/gmail.insert",
    "https://www.googleapis.com/auth/gmail.labels",
]

# Gmail rejects raw messages larger than ~50 MB.
MAX_MESSAGE_BYTES = 50 * 1024 * 1024

# Reserved Gmail system label names that cannot be created as user labels.
RESERVED_LABELS = {
    "INBOX", "SENT", "DRAFT", "DRAFTS", "SPAM", "TRASH", "CHAT",
    "STARRED", "IMPORTANT", "UNREAD",
}

STATE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "state")


# --------------------------------------------------------------------------- #
# Auth
# --------------------------------------------------------------------------- #
def build_gmail_service(key_path, user):
    """Build a Gmail API client impersonating `user` via domain-wide delegation."""
    if _GOOGLE_IMPORT_ERROR is not None:
        sys.exit(
            "Missing dependencies. Activate your venv and run:\n"
            "    pip install -r requirements.txt"
        )
    creds = service_account.Credentials.from_service_account_file(
        key_path, scopes=SCOPES
    )
    delegated = creds.with_subject(user)
    return build("gmail", "v1", credentials=delegated, cache_discovery=False)


# --------------------------------------------------------------------------- #
# Source discovery
# --------------------------------------------------------------------------- #
def sanitize_segment(name):
    """Clean one folder-name segment for use in a Gmail label path."""
    name = name.replace("/", "-").strip()
    name = re.sub(r"\s+", " ", name)
    return name or "Untitled"


def path_to_label(rel_path, drop_ext, prefix):
    """Turn a relative file/dir path into a nested Gmail label name."""
    parts = [p for p in rel_path.split(os.sep) if p not in ("", ".")]
    if drop_ext and parts:
        parts[-1] = os.path.splitext(parts[-1])[0]
    segments = [sanitize_segment(p) for p in parts if p]
    if prefix:
        segments = [sanitize_segment(prefix)] + segments
    return "/".join(segments) if segments else "Imported"


def discover_sources(source, prefix):
    """
    Return a list of (label, kind, path) work units.

    kind == "mbox": path is an mbox file -> iterate messages within it.
    kind == "eml":  path is a single .eml file -> one message.

    Auto-detects: if any .mbox files exist under source we use mbox mode;
    otherwise we treat .eml files as a folder tree.
    """
    source = os.path.abspath(source)

    if os.path.isfile(source):
        if source.lower().endswith(".eml"):
            label = path_to_label(os.path.basename(source), True, prefix)
            return [(label, "eml", source)]
        label = path_to_label(os.path.basename(source), True, prefix)
        return [(label, "mbox", source)]

    mbox_files, eml_files = [], []
    for root, _dirs, files in os.walk(source):
        for f in files:
            full = os.path.join(root, f)
            low = f.lower()
            if low.endswith(".mbox"):
                mbox_files.append(full)
            elif low.endswith(".eml"):
                eml_files.append(full)

    units = []
    if mbox_files:
        for full in sorted(mbox_files):
            rel = os.path.relpath(full, source)
            units.append((path_to_label(rel, True, prefix), "mbox", full))
    elif eml_files:
        for full in sorted(eml_files):
            rel_dir = os.path.relpath(os.path.dirname(full), source)
            label = path_to_label(rel_dir, False, prefix)
            units.append((label, "eml", full))
    return units


# --------------------------------------------------------------------------- #
# Message helpers
# --------------------------------------------------------------------------- #
def message_id_for(raw_bytes, headers_get):
    """Use the Message-ID header, or a stable hash of the bytes as a fallback."""
    mid = headers_get("Message-ID") or headers_get("Message-Id")
    if mid:
        return mid.strip()
    return "sha256:" + hashlib.sha256(raw_bytes).hexdigest()


def is_unread(headers_get):
    """
    Infer read/unread from the mbox Status / X-Status flags.
    'R' in Status means the message was read. Absent info -> treat as read
    (typical for an archive migration; avoids a huge unread count).
    """
    status = (headers_get("Status") or "") + (headers_get("X-Status") or "")
    if "R" in status:
        return False
    if status:  # status present but no R -> unread
        return True
    return False


def iter_messages(kind, path, on_error=None):
    """Yield (raw_bytes, headers_get_callable) for each message in a work unit.

    A single message that fails to parse is skipped (reported via on_error)
    rather than aborting the whole run. In a large archive one corrupt message
    must not strand the remaining thousands.
    """
    if kind == "mbox":
        box = mailbox.mbox(path, factory=None, create=False)
        try:
            for index, key in enumerate(box.keys()):
                try:
                    msg = box[key]          # parsing happens on access
                    raw = msg.as_bytes()
                except Exception as e:      # noqa: BLE001
                    if on_error:
                        on_error(path, index, e)
                    continue
                yield raw, msg.get
        finally:
            box.close()
    else:  # eml
        with open(path, "rb") as fh:
            raw = fh.read()
        import email

        parsed = email.message_from_bytes(raw)
        yield raw, parsed.get


# --------------------------------------------------------------------------- #
# Label management
# --------------------------------------------------------------------------- #
class LabelManager:
    def __init__(self, service, dry_run):
        self.service = service
        self.dry_run = dry_run
        self.cache = {}  # label name -> id
        if not dry_run:
            self._load_existing()

    def _load_existing(self):
        resp = self.service.users().labels().list(userId="me").execute()
        for lbl in resp.get("labels", []):
            self.cache[lbl["name"]] = lbl["id"]

    def ensure(self, label_name):
        """Return the id for label_name, creating it (and ancestors) if needed."""
        if self.dry_run:
            return "DRYRUN"
        if label_name in self.cache:
            return self.cache[label_name]

        # Create each ancestor in the path so nesting is clean.
        parts = label_name.split("/")
        running = []
        last_id = None
        for part in parts:
            running.append(part)
            name = "/".join(running)
            if name in self.cache:
                last_id = self.cache[name]
                continue
            last_id = self._create(name)
            self.cache[name] = last_id
        return last_id

    def _create(self, name):
        create_name = name
        # A leaf segment matching a reserved system name can't be created.
        if name.split("/")[-1].upper() in RESERVED_LABELS:
            create_name = name + " (Imported)"
            if create_name in self.cache:
                return self.cache[create_name]
        body = {
            "name": create_name,
            "labelListVisibility": "labelShow",
            "messageListVisibility": "show",
        }
        try:
            lbl = self.service.users().labels().create(
                userId="me", body=body
            ).execute()
            return lbl["id"]
        except HttpError as e:
            # 409: already exists (race or case difference) -> reload and reuse.
            if e.resp.status == 409:
                self._load_existing()
                if create_name in self.cache:
                    return self.cache[create_name]
            raise


# --------------------------------------------------------------------------- #
# Checkpoint ledger
# --------------------------------------------------------------------------- #
def ledger_paths(user):
    safe = re.sub(r"[^A-Za-z0-9_.@-]", "_", user)
    os.makedirs(STATE_DIR, exist_ok=True)
    return (
        os.path.join(STATE_DIR, f"{safe}.jsonl"),
        os.path.join(STATE_DIR, f"{safe}-errors.log"),
    )


def load_done(ledger_file):
    done = set()
    if os.path.exists(ledger_file):
        with open(ledger_file, "r", encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if not line:
                    continue
                try:
                    done.add(json.loads(line)["message_id"])
                except (ValueError, KeyError):
                    continue
    return done


# --------------------------------------------------------------------------- #
# Insert with backoff
# --------------------------------------------------------------------------- #
class DailyLimitReached(Exception):
    """Gmail signalled the per-user/project daily quota is exhausted for today.

    Unlike per-minute rate limits (which we retry), this won't clear until the
    quota resets at midnight Pacific, so the caller should stop the whole run
    and resume on a later day via the checkpoint.
    """

    def __init__(self, reason):
        self.reason = reason
        super().__init__(f"Gmail daily limit reached ({reason})")


def parse_size(text):
    """Parse a human size like '450MB' or '2GB' into bytes. Plain int = bytes."""
    s = text.strip().upper()
    for suffix, mult in (("GB", 1024**3), ("MB", 1024**2), ("KB", 1024), ("B", 1)):
        if s.endswith(suffix):
            return int(float(s[:-len(suffix)]) * mult)
    return int(s)


def insert_message(service, raw_bytes, label_ids, max_retries=6):
    media = MediaIoBaseUpload(
        io.BytesIO(raw_bytes), mimetype="message/rfc822", resumable=False
    )
    body = {"labelIds": label_ids}
    attempt = 0
    while True:
        try:
            return service.users().messages().insert(
                userId="me",
                internalDateSource="dateHeader",
                body=body,
                media_body=media,
            ).execute()
        except HttpError as e:
            status = e.resp.status
            reason = ""
            try:
                reason = json.loads(e.content.decode("utf-8"))["error"]["errors"][0]["reason"]
            except Exception:
                pass
            # Daily quota / bandwidth exhaustion won't clear until the Pacific
            # midnight reset -- stop the run cleanly instead of hammering it.
            if reason in ("dailyLimitExceeded", "quotaExceeded"):
                raise DailyLimitReached(reason)
            retryable = status in (429, 500, 502, 503) or reason in (
                "rateLimitExceeded", "userRateLimitExceeded", "backendError",
            )
            attempt += 1
            if not retryable or attempt > max_retries:
                raise
            sleep = min(60, (2 ** attempt)) + random.uniform(0, 1)
            time.sleep(sleep)


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser(
        description="Import an Outlook mbox/eml archive into Workspace Gmail, "
        "preserving folders as labels."
    )
    ap.add_argument("--source", required=True,
                    help="Directory tree of .mbox or .eml files (or a single file).")
    ap.add_argument("--user", required=True,
                    help="Mailbox to import into, e.g. jane@client.com.")
    ap.add_argument("--key", required=True,
                    help="Path to the service-account JSON key.")
    ap.add_argument("--label-prefix", default="",
                    help="Optional top-level label to nest everything under "
                         "(e.g. 'Outlook').")
    ap.add_argument("--dry-run", action="store_true",
                    help="Preview the label tree and counts; write nothing.")
    ap.add_argument("--max", type=int, default=0,
                    help="Import at most N messages (0 = no limit). For testing.")
    ap.add_argument("--no-resume", action="store_true",
                    help="Ignore the checkpoint ledger and attempt every message.")
    ap.add_argument("--max-bytes-per-run", default="",
                    help="Cap total inserted bytes per run, e.g. '450MB' or '2GB'. "
                         "Stops cleanly when reached; re-run to continue. Useful for "
                         "staying under Gmail's per-user daily upload limit.")
    args = ap.parse_args()

    max_bytes = parse_size(args.max_bytes_per_run) if args.max_bytes_per_run else 0

    if not os.path.exists(args.source):
        sys.exit(f"Source not found: {args.source}")
    if not args.dry_run and not os.path.exists(args.key):
        sys.exit(f"Service-account key not found: {args.key}")

    units = discover_sources(args.source, args.label_prefix)
    if not units:
        sys.exit("No .mbox or .eml files found under the source path.")

    # Group work units by label for a tidy preview.
    by_label = {}
    for label, kind, path in units:
        by_label.setdefault(label, []).append((kind, path))

    print(f"Source : {os.path.abspath(args.source)}")
    print(f"Mailbox: {args.user}")
    mode = "eml" if units[0][1] == "eml" else "mbox"
    print(f"Format : {mode}")
    print(f"Labels detected ({len(by_label)}):")
    for label in sorted(by_label):
        print(f"  - {label}")
    print()

    ledger_file, error_file = ledger_paths(args.user)
    done = set() if args.no_resume else load_done(ledger_file)
    if done:
        print(f"Resuming: {len(done)} message(s) already imported, will skip.\n")

    service = None if args.dry_run else build_gmail_service(args.key, args.user)
    labels = LabelManager(service, args.dry_run)

    stats = {"inserted": 0, "skipped": 0, "failed": 0, "scanned": 0, "bytes": 0}
    limit_hit = False
    stop_run = False
    stop_reason = ""

    def on_msg_error(path, index, exc):
        """A message that couldn't be parsed: log it and keep going."""
        stats["failed"] += 1
        if args.dry_run:
            print(f"  ! Skipping unparseable message #{index} in {path}: {exc}")
        else:
            with open(error_file, "a", encoding="utf-8") as ef:
                ef.write(f"{_now()}\tPARSE_ERROR\t{path}\t#{index}\t{exc}\n")

    for label in sorted(by_label):
        label_id = labels.ensure(label)
        per_label = 0
        for kind, path in by_label[label]:
            try:
                for raw, hget in iter_messages(kind, path, on_error=on_msg_error):
                    if args.max and stats["inserted"] >= args.max:
                        limit_hit = True
                        break
                    stats["scanned"] += 1
                    mid = message_id_for(raw, hget)

                    if mid in done:
                        stats["skipped"] += 1
                        continue

                    if args.dry_run:
                        per_label += 1
                        continue

                    if len(raw) > MAX_MESSAGE_BYTES:
                        stats["failed"] += 1
                        with open(error_file, "a", encoding="utf-8") as ef:
                            ef.write(f"{_now()}\tTOO_LARGE\t{label}\t{mid}\t"
                                     f"{len(raw)} bytes\n")
                        continue

                    # Stop before exceeding the per-run byte cap (stay under the
                    # daily upload ceiling). The checkpoint lets a re-run continue.
                    if max_bytes and stats["bytes"] + len(raw) > max_bytes:
                        stop_run = True
                        stop_reason = f"--max-bytes-per-run cap ({args.max_bytes_per_run})"
                        break

                    label_ids = [label_id]
                    if is_unread(hget):
                        label_ids.append("UNREAD")

                    try:
                        result = insert_message(service, raw, label_ids)
                        stats["inserted"] += 1
                        stats["bytes"] += len(raw)
                        per_label += 1
                        done.add(mid)
                        with open(ledger_file, "a", encoding="utf-8") as lf:
                            lf.write(json.dumps({
                                "message_id": mid,
                                "label": label,
                                "gmail_id": result.get("id"),
                            }) + "\n")
                        if stats["inserted"] % 100 == 0:
                            print(f"  ... {stats['inserted']} imported so far "
                                  f"({stats['bytes'] // (1024 * 1024)} MB this run)")
                    except DailyLimitReached as e:
                        stop_run = True
                        stop_reason = f"Gmail daily limit reached ({e.reason})"
                        break
                    except Exception as e:  # noqa: BLE001
                        stats["failed"] += 1
                        with open(error_file, "a", encoding="utf-8") as ef:
                            ef.write(f"{_now()}\tFAILED\t{label}\t{mid}\t{e}\n")
            except Exception as e:  # noqa: BLE001
                print(f"  ! Could not read {path}: {e}")
                continue
            if limit_hit or stop_run:
                break
        if args.dry_run:
            print(f"  {label}: {per_label} message(s)")
        if limit_hit:
            print("\nReached --max limit; stopping.")
            break
        if stop_run:
            break

    print("\n" + "=" * 48)
    if args.dry_run:
        print(f"DRY RUN: scanned {stats['scanned']} message(s) across "
              f"{len(by_label)} label(s). Nothing written.")
    else:
        print("Run stopped early." if stop_run else "Import complete.")
        print(f"  Inserted: {stats['inserted']} "
              f"({stats['bytes'] // (1024 * 1024)} MB this run)")
        print(f"  Skipped (already imported): {stats['skipped']}")
        print(f"  Failed:   {stats['failed']}")
        if stats["failed"]:
            print(f"  See errors: {error_file}")
        print(f"  Checkpoint: {ledger_file}")
        if stop_run:
            print(f"\n  Stopped: {stop_reason}")
            print("  This is expected for a large import. Re-run the SAME command to")
            print("  continue -- already-imported messages are skipped automatically.")


def _now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


if __name__ == "__main__":
    main()
