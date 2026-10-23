"""Bounded handling of observed Windows sharing locks during report save."""
import unittest
from unittest.mock import patch

from reconciliation import _replace_file


def sharing_error():
    error = PermissionError("temporarily locked")
    error.winerror = 5
    return error


class AtomicSaveTests(unittest.TestCase):
    def test_temporary_windows_lock_is_retried_before_success(self):
        with patch("reconciliation.os.replace", side_effect=[sharing_error(), None]) as replace:
            with patch("reconciliation.clock_time.sleep") as pause:
                _replace_file("source.xlsx", "target.xlsx")
        self.assertEqual(replace.call_count, 2)
        pause.assert_called_once_with(0.05)

    def test_permanent_lock_is_reported_after_bounded_retries(self):
        with patch("reconciliation.os.replace", side_effect=sharing_error()) as replace:
            with patch("reconciliation.clock_time.sleep") as pause:
                with self.assertRaises(PermissionError):
                    _replace_file("source.xlsx", "target.xlsx")
        self.assertEqual(replace.call_count, 5)
        self.assertEqual(pause.call_count, 4)
