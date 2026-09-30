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
    ("[set]/amp & semi;.txt", "esc", 0),    # wildcard and XML characters
    ("other/amp & semi;.txt", "esc", 0),
    ("x/Holiday/p1.jpg", "h1", 0),          # duplicate folders, with a nested duplicate
    ("x/Holiday/inner/p2.jpg", "h2", 0),
    ("y/Holiday/p1.jpg", "h1", 0),
    ("y/Holiday/inner/p2.jpg", "h2", 0),
    ("z/inner/p2.jpg", "h2", 0),            # a third copy of "inner", outside Holiday
]
if os.name != "nt":  # characters Windows does not allow in file names
    TREE += [("a/less <than>.txt", "lt", 0), ("b/less <than>.txt", "lt", 0)]

# Removed before validating: one copy of a three-copy set, one of a two-copy set,
# and a file inside one copy of the duplicate "Holiday" folder.
REMOVED_BEFORE_VALIDATE = ["b/report.doc", "b/photo.jpg", "y/Holiday/p1.jpg"]
FOLDER_SHEET = "xl/worksheets/sheet2.xml"


@unittest.skipUnless(PWSH or REQUIRED, "PowerShell 7 (pwsh) is not installed")
class ParityTests(unittest.TestCase):
    def setUp(self):
        self.assertTrue(PWSH, "pwsh is required for the parity test but was not found")
        self._temp = tempfile.TemporaryDirectory()
        self.root = self._temp.name
        self.data = os.path.join(self.root, "data")
        for path, content, offset in TREE:
            add_file(self.data, path, content, SAVED + timedelta(seconds=offset))

    def tearDown(self):
        self._temp.cleanup()

    def report(self, name):
        return os.path.join(self.root, name)

    def run_powershell(self, *args):
        ps = subprocess.run(
            [PWSH, "-NoProfile", "-NonInteractive", "-File", PS_SCRIPT, *args],
            capture_output=True, text=True,
        )
        self.assertEqual(ps.returncode, 0, f"PowerShell failed:\n{ps.stdout}\n{ps.stderr}")

    def run_python(self, *args):
        with contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(cli.main(list(args)), 0)

    def assert_same_report(self, ps_report, py_report):
        self.assertEqual(read_worksheet(py_report, FOLDER_SHEET), read_worksheet(ps_report, FOLDER_SHEET))
        ps_rows, py_rows = read_worksheet(ps_report), read_worksheet(py_report)
        self.assertEqual(len(py_rows), len(ps_rows))
        for number, (ps_row, py_row) in enumerate(zip(ps_rows, py_rows), start=1):
            with self.subTest(row=number):
                if number > 1:
                    # Both are the same instant; allow float formatting differences.
                    self.assertTrue(math.isclose(float(ps_row[DATE_COLUMN]), float(py_row[DATE_COLUMN]), abs_tol=1e-8))
                    ps_row, py_row = _without(ps_row, DATE_COLUMN), _without(py_row, DATE_COLUMN)
                self.assertEqual(py_row, ps_row)
        self.assertEqual(_parts(py_report), _parts(ps_report))

    def test_scan_reports_match(self):
        self.run_powershell("-Path", self.data, "-OutputFile", self.report("ps.xlsx"), "-IncludeFolders")
        self.run_python(self.data, self.report("py.xlsx"), "--folders")
        self.assertGreater(len(read_worksheet(self.report("ps.xlsx"))), 1, "the fixture should produce duplicates")
        folder_names = [row[0] for row in read_worksheet(self.report("ps.xlsx"), FOLDER_SHEET)[1:]]
        self.assertIn("Holiday", folder_names)
        self.assertIn("inner", folder_names, "the nested set with a copy outside Holiday is kept")
        self.assert_same_report(self.report("ps.xlsx"), self.report("py.xlsx"))

    def test_parallel_hashing_reports_match(self):
        self.run_powershell("-Path", self.data, "-OutputFile", self.report("ps.xlsx"), "-IncludeFolders",
                            "-ThrottleLimit", "4")
        self.run_python(self.data, self.report("py.xlsx"), "--folders", "--throttle-limit", "4")
        self.assert_same_report(self.report("ps.xlsx"), self.report("py.xlsx"))

    def test_validated_reports_match(self):
        self.run_powershell("-Path", self.data, "-OutputFile", self.report("ps.xlsx"), "-IncludeFolders")
        self.run_python(self.data, self.report("py.xlsx"), "--folders")
        for relative in REMOVED_BEFORE_VALIDATE:
            os.remove(os.path.join(self.data, *relative.split("/")))

        self.run_powershell("-Validate", "-OutputFile", self.report("ps.xlsx"))
        self.run_python("--validate", self.report("py.xlsx"))

        self.assert_same_report(self.report("ps.xlsx"), self.report("py.xlsx"))
        names = [row[0] for row in read_worksheet(self.report("py.xlsx"))[1:]]
        self.assertNotIn("Photo.JPG", names, "a row left with one copy is removed")


def _without(row, index):
    return row[:index] + row[index + 1:]


def _parts(path):
    """Every package part except the worksheets (compared cell by cell), parsed and re-serialised."""
    with zipfile.ZipFile(path) as archive:
        return {
            name: ElementTree.tostring(ElementTree.fromstring(archive.read(name)))
            for name in sorted(archive.namelist()) if not name.startswith("xl/worksheets/")
        }


if __name__ == "__main__":
    unittest.main()
