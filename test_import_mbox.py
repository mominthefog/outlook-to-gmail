#!/usr/bin/env python3
"""Offline unit tests for import_mbox.py.

These exercise the pure helpers and the mbox reader only -- no Google
credentials, no network. Run with:

    python -m unittest -v
"""

import mailbox
import os
import tempfile
import unittest

import import_mbox as m


class ParseSizeTests(unittest.TestCase):
    def test_units(self):
        self.assertEqual(m.parse_size("450MB"), 450 * 1024**2)
        self.assertEqual(m.parse_size("2GB"), 2 * 1024**3)
        self.assertEqual(m.parse_size("512KB"), 512 * 1024)
        self.assertEqual(m.parse_size("1024B"), 1024)

    def test_plain_int_is_bytes(self):
        self.assertEqual(m.parse_size("1500"), 1500)

    def test_fractional_and_case_insensitive(self):
        self.assertEqual(m.parse_size("1.5gb"), int(1.5 * 1024**3))
        self.assertEqual(m.parse_size(" 450 mb ".replace(" mb ", "MB").strip()),
                         450 * 1024**2)


class SanitizeAndLabelTests(unittest.TestCase):
    def test_sanitize_segment(self):
        self.assertEqual(m.sanitize_segment("a/b"), "a-b")
        self.assertEqual(m.sanitize_segment("  spaced   out "), "spaced out")
        self.assertEqual(m.sanitize_segment("   "), "Untitled")

    def test_path_to_label_nesting(self):
        label = m.path_to_label(os.path.join("Inbox", "Projects.mbox"), True, "")
        self.assertEqual(label, "Inbox/Projects")

    def test_path_to_label_prefix(self):
        label = m.path_to_label(os.path.join("Inbox", "Projects.mbox"), True,
                                "Outlook")
        self.assertEqual(label, "Outlook/Inbox/Projects")

    def test_path_to_label_keeps_ext_when_not_dropping(self):
        label = m.path_to_label(os.path.join("Inbox", "msg"), False, "")
        self.assertEqual(label, "Inbox/msg")

    def test_empty_path_falls_back(self):
        self.assertEqual(m.path_to_label("", True, ""), "Imported")


class ReadUnreadTests(unittest.TestCase):
    def _hget(self, headers):
        return lambda name, default=None: headers.get(name, default)

    def test_read_flag(self):
        self.assertFalse(m.is_unread(self._hget({"Status": "RO"})))

    def test_unread_when_status_present_without_r(self):
        self.assertTrue(m.is_unread(self._hget({"Status": "O"})))

    def test_no_status_treated_as_read(self):
        self.assertFalse(m.is_unread(self._hget({})))


class MessageIdTests(unittest.TestCase):
    def test_uses_message_id_header(self):
        def hget(name, default=None):
            return {"Message-ID": " <abc@x> "}.get(name, default)
        self.assertEqual(m.message_id_for(b"raw", hget), "<abc@x>")

    def test_hash_fallback_is_stable(self):
        def hget(name, default=None):
            return None
        a = m.message_id_for(b"same bytes", hget)
        b = m.message_id_for(b"same bytes", hget)
        self.assertTrue(a.startswith("sha256:"))
        self.assertEqual(a, b)


class DiscoverAndIterTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()

    def _write_mbox(self, rel, messages):
        path = os.path.join(self.tmp, rel)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        box = mailbox.mbox(path)
        for subject, body, status in messages:
            msg = mailbox.mboxMessage()
            msg["Subject"] = subject
            msg["Message-ID"] = f"<{subject}@test>"
            if status:
                msg["Status"] = status
            msg.set_payload(body)
            box.add(msg)
        box.flush()
        box.close()
        return path

    def test_discover_sources_maps_folders_to_labels(self):
        self._write_mbox("Inbox.mbox", [("a", "x", "RO")])
        self._write_mbox(os.path.join("Work", "Projects.mbox"), [("b", "y", "O")])
        units = m.discover_sources(self.tmp, "")
        labels = sorted(u[0] for u in units)
        self.assertEqual(labels, ["Inbox", "Work/Projects"])
        self.assertTrue(all(kind == "mbox" for _, kind, _ in units))

    def test_iter_messages_yields_and_preserves_status(self):
        path = self._write_mbox("Inbox.mbox", [("read", "x", "RO"),
                                               ("unread", "y", "O")])
        seen = list(m.iter_messages("mbox", path))
        self.assertEqual(len(seen), 2)
        reads = {hget("Subject"): m.is_unread(hget) for _, hget in seen}
        self.assertFalse(reads["read"])
        self.assertTrue(reads["unread"])

    def test_iter_messages_reports_errors_without_aborting(self):
        # Simulate a message whose access raises, and confirm iteration
        # continues for the rest (the resilience fix).
        path = self._write_mbox("Inbox.mbox", [("a", "x", "RO"),
                                               ("b", "y", "RO"),
                                               ("c", "z", "RO")])
        box = mailbox.mbox(path)
        keys = list(box.keys())
        box.close()

        real_getitem = mailbox.mbox.__getitem__
        bad_key = keys[1]

        def flaky_getitem(self, key):
            if key == bad_key:
                raise ValueError("corrupt message")
            return real_getitem(self, key)

        errors = []
        mailbox.mbox.__getitem__ = flaky_getitem
        try:
            got = list(m.iter_messages("mbox", path,
                                       on_error=lambda p, i, e: errors.append((i, str(e)))))
        finally:
            mailbox.mbox.__getitem__ = real_getitem

        self.assertEqual(len(got), 2)          # the two good messages still came through
        self.assertEqual(len(errors), 1)       # the bad one was reported, not fatal
        self.assertIn("corrupt message", errors[0][1])


if __name__ == "__main__":
    unittest.main()
