"""Bounded handling of observed Windows sharing locks during report save."""
import unittest
import tempfile
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from reconciliation import _preserve_cell_values, _replace_file


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

    def test_value_preservation_replace_failure_cleans_staged_workbook(self):
        with tempfile.TemporaryDirectory(prefix="gawa-value-save-") as directory:
            folder = Path(directory)
            original, target = folder / "original.xlsx", folder / "target.xlsx"
            for path, value in ((original, "原始数据"), (target, "已有输出")):
                workbook = Workbook()
                workbook.active["A1"] = value
                workbook.save(path)
                workbook.close()
            preserved = {path: path.read_bytes() for path in (original, target)}
            with patch("reconciliation._replace_file", side_effect=PermissionError("permanently locked")) as replace:
                with self.assertRaises(PermissionError):
                    _preserve_cell_values(original, target)
            replace.assert_called_once()
            self.assertEqual(set(folder.iterdir()), {original, target})
            for path, content in preserved.items():
                self.assertEqual(path.read_bytes(), content)
