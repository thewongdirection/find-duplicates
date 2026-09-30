"""Feature parity: run the PowerShell tool and the Python port on the same
folder tree and require identical reports.

Needs PowerShell 7 (``pwsh``) on the PATH; skipped otherwise, unless the
FIND_DUPLICATES_REQUIRE_PARITY environment variable is set (as in CI), in
which case a missing ``pwsh`` is a failure.
"""

from __future__ import annotations

import contextlib
import io
import math
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from datetime import timedelta
from xml.etree import ElementTree

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from find_duplicates import cli  # noqa: E402
from tests.helpers import SAVED, add_file, read_worksheet  # noqa: E402

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
PS_SCRIPT = os.path.join(REPO_ROOT, "Find-Duplicates.ps1")
PWSH = shutil.which("pwsh")
REQUIRED = bool(os.environ.get("FIND_DUPLICATES_REQUIRE_PARITY"))
DATE_COLUMN = 1

# name, content, saved-date offset in seconds: covers every matching rule.
TREE = [
    ("a/report.doc", "same", 0),
    ("b/report.doc", "same", 0),
    ("b/c/d/report.doc", "same", 0),
    ("a/Photo.JPG", "img", 0),              # names differ only by case
    ("b/photo.jpg", "img", 0.4),            # sub-second difference
    ("a/x.txt", "first", 0),                # two sets under one name + date
    ("b/x.txt", "first", 0),
    ("c/x.txt", "other", 0),
    ("d/x.txt", "other", 0),
    ("e/x.txt", "unique", 0),
    ("a/late.txt", "same", 0),              # different saved dates
    ("b/late.txt", "same", 60),
    ("a/size.txt", "short", 0),             # different sizes
    ("b/size.txt", "much longer", 0),
    ("[set]/amp & <lt>.txt", "esc", 0),     # wildcard and XML characters
    ("other/amp & <lt>.txt", "esc", 0),
]


@unittest.skipUnless(PWSH or REQUIRED, "PowerShell 7 (pwsh) is not installed")
class ParityTests(unittest.TestCase):
    def test_python_and_powershell_reports_match(self):
        self.assertTrue(PWSH, "pwsh is required for the parity test but was not found")
        with tempfile.TemporaryDirectory() as root:
            data = os.path.join(root, "data")
            for path, content, offset in TREE:
                add_file(data, path, content, SAVED + timedelta(seconds=offset))
            ps_report = os.path.join(root, "ps.xlsx")
            py_report = os.path.join(root, "py.xlsx")

            ps = subprocess.run(
                [PWSH, "-NoProfile", "-NonInteractive", "-File", PS_SCRIPT,
                 "-Path", data, "-OutputFile", ps_report],
                capture_output=True, text=True,
            )
            self.assertEqual(ps.returncode, 0, f"PowerShell failed:\n{ps.stdout}\n{ps.stderr}")
            with contextlib.redirect_stdout(io.StringIO()):
                self.assertEqual(cli.main([data, py_report]), 0)

            ps_rows, py_rows = read_worksheet(ps_report), read_worksheet(py_report)
            self.assertEqual(len(py_rows), len(ps_rows))
            self.assertGreater(len(ps_rows), 1, "the fixture should produce duplicates")
            for number, (ps_row, py_row) in enumerate(zip(ps_rows, py_rows), start=1):
                with self.subTest(row=number):
                    if number > 1:
                        # Both are the same instant; allow float rounding (well under 1 ms).
                        self.assertTrue(math.isclose(float(ps_row[DATE_COLUMN]), float(py_row[DATE_COLUMN]),
                                                     abs_tol=1e-8))
                        ps_row, py_row = _without(ps_row, DATE_COLUMN), _without(py_row, DATE_COLUMN)
                    self.assertEqual(py_row, ps_row)

            self.assertEqual(_parts(py_report, skip="xl/worksheets/sheet1.xml"),
                             _parts(ps_report, skip="xl/worksheets/sheet1.xml"))


def _without(row, index):
    return row[:index] + row[index + 1:]


def _parts(path, skip):
    """Every package part except ``skip``, parsed and re-serialised for comparison."""
    with zipfile.ZipFile(path) as archive:
        return {
            name: ElementTree.tostring(ElementTree.fromstring(archive.read(name)))
            for name in sorted(archive.namelist()) if name != skip
        }


if __name__ == "__main__":
    unittest.main()
