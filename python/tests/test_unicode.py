"""Unicode file and folder names. Mirrors the 'Unicode names' Describe block in
tests/DuplicateFinder.Tests.ps1, plus checks specific to Python."""

from __future__ import annotations

import os
import subprocess
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from find_duplicates import scanner  # noqa: E402
from find_duplicates.folders import find_duplicate_folders  # noqa: E402
from find_duplicates.matcher import find_duplicate_files  # noqa: E402
from find_duplicates.names import name_key, path_sort_key, sort_key  # noqa: E402
from find_duplicates.scanner import iter_files  # noqa: E402
from find_duplicates.validate import validate_report  # noqa: E402
from find_duplicates.xlsx import (  # noqa: E402
    export_duplicate_report, read_duplicate_folder_report, read_duplicate_report,
)
from tests.helpers import add_file  # noqa: E402

PYTHON_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def from_code_points(*code_points: int) -> str:
    # Built from code points, like the Pester tests, so both suites use identical names.
    return "".join(chr(c) for c in code_points)


NAMES = {
    "Japanese": from_code_points(0x65E5, 0x672C, 0x8A9E),
    "Arabic": from_code_points(0x0645, 0x0644, 0x0641),
    "Cyrillic": from_code_points(0x0444, 0x0430, 0x0439, 0x043B),
    "Emoji": from_code_points(0x1F600, 0x1F4F7),
}
COMPOSED = from_code_points(0x63, 0x61, 0x66, 0xE9)           # cafe with e-acute as one character
DECOMPOSED = from_code_points(0x63, 0x61, 0x66, 0x65, 0x301)  # e followed by a combining accent


class UnicodeNameTests(unittest.TestCase):
    def setUp(self):
        self._temp = tempfile.TemporaryDirectory()
        self.root = scanner.full_path(self._temp.name)

    def tearDown(self):
        self._temp.cleanup()

    def new_root(self, name):
        path = os.path.join(self.root, name)
        os.makedirs(path)
        return path

    def test_finds_saves_reads_back_and_validates_duplicates_named_in_many_scripts(self):
        for script, name in NAMES.items():
            with self.subTest(script):
                root = self.new_root(f"tree-{script}")
                add_file(root, f"{name}/{name}.txt")
                add_file(root, f"copy/{name}.txt")
                report = os.path.join(self.root, f"{script}-{name}.xlsx")

                found = find_duplicate_files(list(iter_files(root)))
                export_duplicate_report(found, report)
                (row,) = read_duplicate_report(report)
                result = validate_report(report)

                self.assertEqual(len(found), 1)
                self.assertEqual(row.file_name, f"{name}.txt")
                self.assertIn(os.path.join(root, name), row.folders)
                self.assertEqual(result.copies_removed, 0)

    def test_matches_names_stored_in_different_unicode_forms(self):
        first = add_file(self.root, f"a/{COMPOSED}.txt")
        second = add_file(self.root, f"b/{DECOMPOSED}.txt")
        records = list(iter_files(self.root))
        if {r.name for r in records} == {os.path.basename(first)}:
            self.skipTest("this file system normalises names itself")
        self.assertTrue(os.path.exists(second))
        self.assertEqual(len(find_duplicate_files(records)), 1)

    def test_keeps_a_copy_whose_name_is_stored_in_another_unicode_form_when_validating(self):
        add_file(self.root, f"a/{COMPOSED}.txt")
        add_file(self.root, f"b/{DECOMPOSED}.txt")
        if len({r.name for r in iter_files(self.root)}) == 1:
            self.skipTest("this file system normalises names itself")
        report = os.path.join(self.root, "forms.xlsx")
        export_duplicate_report(find_duplicate_files(list(iter_files(self.root))), report)
        self.assertEqual(validate_report(report).copies_removed, 0)

    def test_finds_duplicate_folders_with_unicode_names(self):
        for parent in ("one", "two"):
            add_file(self.root, f"{parent}/{NAMES['Japanese']}/{NAMES['Emoji']}.jpg")
        records = []
        files = list(iter_files(self.root, folders=records))
        (result,) = find_duplicate_folders(files, records)
        self.assertEqual(result.folder_name, NAMES["Japanese"])

    def test_handles_unicode_names_end_to_end_through_the_command_line(self):
        data = self.new_root("data")
        for parent in ("one", "two"):
            add_file(data, f"{parent}/{NAMES['Arabic']}/{NAMES['Cyrillic']}.txt")
        report = os.path.join(self.root, f"{NAMES['Emoji']}.xlsx")
        completed = self.run_cli(data, report, "--folders")
        self.assertEqual(completed.returncode, 0, completed.stderr)
        self.assertEqual(read_duplicate_report(report)[0].file_name, f"{NAMES['Cyrillic']}.txt")
        self.assertEqual(read_duplicate_folder_report(report)[0].folder_name, NAMES["Arabic"])

    def test_does_not_crash_when_the_console_cannot_show_unicode(self):
        data = self.new_root(NAMES["Japanese"])
        add_file(data, "a/x.txt")
        add_file(data, "b/x.txt")
        completed = self.run_cli(data, os.path.join(self.root, "ascii.xlsx"), "--verbose", encoding="ascii")
        self.assertEqual(completed.returncode, 0, completed.stderr)
        self.assertIn("\\u65e5", completed.stdout, "characters the console cannot show are escaped")

    @unittest.skipIf(sys.platform in ("win32", "darwin"), "only Linux allows file names that are not valid UTF-8")
    def test_reports_file_names_that_are_not_valid_utf8(self):
        bad = os.fsdecode(b"bad-\xff.txt")  # arrives as a lone surrogate
        for folder in ("a", "b"):
            os.makedirs(os.path.join(self.root, folder))
            with open(os.path.join(self.root, folder, bad), "w") as stream:
                stream.write("x")
        report = os.path.join(self.root, "bad.xlsx")
        export_duplicate_report(find_duplicate_files(list(iter_files(self.root))), report)
        self.assertEqual(read_duplicate_report(report)[0].file_name, "bad-" + chr(0xFFFD) + ".txt")

    def run_cli(self, *args, encoding="utf-8"):
        env = dict(os.environ, PYTHONIOENCODING=encoding, PYTHONPATH=PYTHON_ROOT)
        return subprocess.run(
            [sys.executable, "-m", "find_duplicates", *args],
            capture_output=True, env=env, cwd=self.root, text=True, encoding="utf-8", errors="replace",
        )


class NameKeyTests(unittest.TestCase):
    def test_compares_names_ignoring_case_and_unicode_form(self):
        self.assertEqual(name_key(COMPOSED), name_key(DECOMPOSED.upper()))

    def test_does_not_expand_characters_whose_upper_case_is_longer(self):
        self.assertNotEqual(name_key("stra" + chr(0xDF) + "e"), name_key("STRASSE"))

    def test_orders_like_dotnet_ordinal_comparison(self):
        # U+FF21 (full-width A) comes after an emoji in UTF-16 order, the opposite of
        # code-point order; .NET, and so PowerShell, uses UTF-16 order.
        emoji, full_width_a = NAMES["Emoji"], chr(0xFF21)
        self.assertEqual(sorted([full_width_a, emoji], key=sort_key), [emoji, full_width_a])
        self.assertEqual(sorted([full_width_a, emoji], key=lambda t: sort_key(t, True)), [emoji, full_width_a])

    def test_orders_paths_differing_only_in_case_in_a_fixed_order(self):
        self.assertEqual(sorted(["/x/photos", "/x/Photos"], key=path_sort_key), ["/x/Photos", "/x/photos"])


if __name__ == "__main__":
    unittest.main()
