# MBOX → Gmail (Workspace) Importer

Import an Outlook for Mac archive into a Google Workspace mailbox, preserving the
original **folder structure as Gmail labels**, the **original dates**, and
**read/unread** state. Built for a super admin to run against a user's mailbox using a
service account with domain-wide delegation, so no end-user login is needed.

Messages are added with the Gmail API's `messages.insert`: they land in **All Mail**
(not the Inbox), skip the spam filter, and keep their original sent/received date.

---

## 1. Convert the `.olm` first (Olminator)

Use **Olminator** to convert the Outlook `.olm` to a folder of mail. Pick the export
option that **preserves the folder structure**. Either of Olminator's outputs works:

- a tree of `.mbox` files (one per Outlook folder), or
- a tree of `.eml` files (each Outlook folder a subfolder, each message a `.eml`).

The importer auto-detects which one you have. Each file/subfolder path becomes the
label, with sub-folders nested using Gmail's `Parent/Child` convention.

---

## 2. One-time Google setup (super admin)

### a. Google Cloud project + service account
1. Go to <https://console.cloud.google.com>, create or select a project.
2. **APIs & Services → Library →** enable the **Gmail API**.
3. **APIs & Services → Credentials → Create credentials → Service account.** Name it
   (e.g. `mbox-importer`) and create it.
4. Open the service account → **Keys → Add key → Create new key → JSON.** Save the
   downloaded file next to this script (e.g. `service-account.json`). **Keep it secret.**
5. On the service account's **Details** page, note its **Unique ID** (a long number).
   This is the OAuth **Client ID** used for delegation.

### b. Authorize domain-wide delegation (Admin console)
1. Go to <https://admin.google.com> → **Security → Access and data control → API
   controls → Domain-wide delegation → Manage domain-wide delegation.**
2. **Add new.** Client ID = the service account's Unique ID from step a.5.
3. OAuth scopes (comma-separated, exactly these two):
   ```
   https://www.googleapis.com/auth/gmail.insert,https://www.googleapis.com/auth/gmail.labels
   ```
4. **Authorize.** Delegation can take a few minutes to propagate.

### c. Python environment
```bash
git clone https://github.com/mominthefog/outlook-to-gmail.git
cd outlook-to-gmail
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

---

## 3. Run it

Always preview first, then test a small batch, then do the full run.

```bash
# 1) Dry run: prints the detected label tree and per-folder counts. Writes nothing.
python import_mbox.py --source ./client-archive/ \
    --user jane@client.com --key service-account.json --dry-run

# 2) Test batch: import 25 messages and check them in Gmail.
python import_mbox.py --source ./client-archive/ \
    --user jane@client.com --key service-account.json --max 25

# 3) Full run.
python import_mbox.py --source ./client-archive/ \
    --user jane@client.com --key service-account.json
```

### Options
| Flag | Meaning |
|------|---------|
| `--source` | Directory tree of `.mbox`/`.eml` (or a single file). |
| `--user` | Mailbox to import into, e.g. `jane@client.com`. |
| `--key` | Path to the service-account JSON key. |
| `--label-prefix` | Optional top-level label to nest everything under (e.g. `Outlook`). |
| `--dry-run` | Preview only; no writes. |
| `--max N` | Import at most N messages (for testing). |
| `--no-resume` | Ignore the checkpoint and attempt every message again. |

---

## 4. Verify

1. **Dry run:** the printed label tree matches the Outlook folders; counts look right.
2. **Test batch (`--max 25`):** in the client's Gmail, confirm the messages appear under
   the correct label, dates look original, read/unread is right, and nothing hit the
   Inbox or Spam.
3. **Full run:** spot-check a couple of labels' message counts in Gmail against the dry
   run.
4. **Idempotency:** re-run the same command. It should skip everything via the
   checkpoint and insert no duplicates.
5. Review `state/<user>-errors.log` for any failures; just re-run to retry them (already
   imported messages are skipped).

---

## 5. How state works

- `state/<user>.jsonl`: one line per imported message (its `Message-ID`, label, and the
  new Gmail id). This is what makes re-runs safe and resumable. Delete it only if you
  want to allow re-importing the same archive.
- `state/<user>-errors.log`: messages that failed after retries, or were skipped for
  being over Gmail's 50 MB limit.

---

## 6. Revoke access when done

1. Admin console → Domain-wide delegation → remove the importer's entry.
2. Google Cloud → the service account → **Keys** → delete the JSON key (and delete the
   service account if it won't be reused).

---

## Notes & limits
- Messages over Gmail's 50 MB limit are skipped and logged (rare, very large attachments).
- If the same message exists in two Outlook folders, it is imported once under the first
  folder seen (Gmail treats a message as one object; this avoids duplicates).
- `.eml` exports usually carry no read/unread flag, so those import as read. `.mbox`
  exports that include `Status` flags have read/unread preserved.
- Attachments travel inside each raw message, so they import intact.
