"""Tests for duplicate folders. Mirrors the folder Describe blocks in tests/DuplicateFinder.Tests.ps1."""

from __future__ import annotations

import contextlib
import io
import os
import shutil
import sys
import tempfile
import unittest
import zipfile
from datetime import timedelta
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from find_duplicates import cli, folders, matcher, scanner, validate  # noqa: E402
from find_duplicates.folders import DuplicateFolderSet, find_duplicate_folders  # noqa: E402
from find_duplicates.matcher import find_duplicate_files  # noqa: E402
from find_duplicates.scanner import FolderRecord, iter_files  # noqa: E402
from find_duplicates.validate import validate_report  # noqa: E402
from find_duplicates.xlsx import (  # noqa: E402
    export_duplicate_report, read_duplicate_folder_report, read_duplicate_report,
)
from tests.helpers import SAVED, add_file  # noqa: E402


def folder_scan(root):
    """Files and folder records for a tree, as the command line collects them."""
    records = []
    files = list(iter_files(root, folders=records))
    return files, records


def add_photo_folder(root, folder, content="photo"):
    """A small tree: two files, one in a sub folder, and an empty sub folder."""
    add_file(root, f"{folder}/a.jpg", f"a {content}")
    add_file(root, f"{folder}/sub/b.jpg", f"b {content}")
    os.makedirs(os.path.join(root, *folder.split("/"), "empty"), exist_ok=True)


class TempRootTestCase(unittest.TestCase):
    def setUp(self):
        self._temp = tempfile.TemporaryDirectory()
        # Long form: the Windows temp folder is often a short (8.3) path, which the tool expands.
        self.root = scanner.full_path(self._temp.name)

    def tearDown(self):
        self._temp.cleanup()

    def path(self, relative):
        return os.path.join(self.root, *relative.split("/"))

    def find(self, **options):
        files, records = folder_scan(self.root)
        return find_duplicate_folders(files, records, **options)


class FindDuplicateFoldersTests(TempRootTestCase):
    def test_finds_folders_with_the_same_name_and_the_same_contents(self):
        add_photo_folder(self.root, "one/Photos")
        add_photo_folder(self.root, "two/Photos")
        (result,) = self.find()
        self.assertEqual(result.folder_name, "Photos")
        self.assertEqual((result.file_count, result.folder_count, result.count), (2, 2, 2))
        self.assertEqual(result.size_bytes, len("a photo") + len("b photo"))
        self.assertEqual(result.folders, [self.path("one/Photos"), self.path("two/Photos")])

    def test_matches_folder_names_that_differ_only_by_case(self):
        add_photo_folder(self.root, "one/Photos")
        add_photo_folder(self.root, "two/photos")
        self.assertEqual(len(self.find()), 1)

    def test_ignores_folders_that_differ_but_still_finds_their_identical_sub_folders(self):
        cases = {
            "names differ": ("two/Pictures", lambda r: None),
            "file contents differ at the same size and date": (
                "two/Photos", lambda r: add_file(r, "two/Photos/a.jpg", "a PHOTO")),
            "file saved dates differ": (
                "two/Photos", lambda r: add_file(r, "two/Photos/a.jpg", "a photo", SAVED + timedelta(minutes=1))),
            "files differ (an extra file)": ("two/Photos", lambda r: add_file(r, "two/Photos/extra.txt")),
            "sub folders differ (an extra empty folder)": (
                "two/Photos", lambda r: os.makedirs(os.path.join(r, "two", "Photos", "more"))),
            "file names differ": (
                "two/Photos", lambda r: os.rename(os.path.join(r, "two", "Photos", "a.jpg"),
                                                  os.path.join(r, "two", "Photos", "c.jpg"))),
        }
        for case, (second, change) in cases.items():
            with self.subTest(case), tempfile.TemporaryDirectory() as root:
                add_photo_folder(root, "one/Photos")
                add_photo_folder(root, second)
                change(root)
                files, records = folder_scan(root)
                # The Photos folders differ, but their untouched "sub" folders are still duplicates.
                self.assertEqual([r.folder_name for r in find_duplicate_folders(files, records)], ["sub"])

    def test_reports_only_the_top_most_duplicate_folders(self):
        add_photo_folder(self.root, "one/Photos/2024")
        add_photo_folder(self.root, "two/Photos/2024")
        self.assertEqual([r.folder_name for r in self.find()], ["Photos"])

    def test_keeps_a_nested_set_that_also_has_a_copy_somewhere_else(self):
        add_photo_folder(self.root, "one/Photos/2024")
        add_photo_folder(self.root, "two/Photos/2024")
        add_photo_folder(self.root, "three/2024")
        result = self.find()
        self.assertEqual([r.folder_name for r in result], ["2024", "Photos"])
        self.assertEqual(result[0].count, 3)

    def test_does_not_report_folders_that_contain_no_files(self):
        for folder in ("one/Empty/inner", "two/Empty/inner"):
            os.makedirs(self.path(folder))
        self.assertEqual(self.find(), [])

    def test_does_not_report_a_folder_that_has_an_unreadable_sub_folder(self):
        add_photo_folder(self.root, "one/Photos")
        add_photo_folder(self.root, "two/Photos")
        files, records = folder_scan(self.root)
        unreadable = self.path("two/Photos/sub")
        records = [FolderRecord(r.path, r.path != unreadable) for r in records]
        self.assertEqual(find_duplicate_folders(files, records), [])

    def test_does_not_read_files_again_that_the_file_scan_already_hashed(self):
        add_photo_folder(self.root, "one/Photos")
        add_photo_folder(self.root, "two/Photos")
        files, records = folder_scan(self.root)
        cache = {}
        find_duplicate_files(files, md5_cache=cache)
        with mock.patch.object(matcher, "md5_file", side_effect=AssertionError("should not be called")) as md5:
            result = find_duplicate_folders(files, records, md5_cache=cache)
        md5.assert_not_called()
        self.assertEqual(len(result), 1)

    def test_does_not_report_folders_holding_online_only_files_with_skip_cloud_only(self):
        for folder in ("one/Photos", "two/Photos", "cloud/Photos"):
            add_photo_folder(self.root, folder)
        with mock.patch.object(folders, "is_cloud_only", side_effect=lambda r: "cloud" in r.path), \
                self.assertLogs("find_duplicates", "WARNING") as logs:
            (result,) = self.find(skip_cloud_only=True)
        self.assertFalse(any("cloud" in f for f in result.folders))
        self.assertRegex(logs.output[0], "2 online-only.*folders")

    def test_finds_the_same_folders_hashing_several_files_at_a_time(self):
        add_photo_folder(self.root, "one/Photos")
        add_photo_folder(self.root, "two/Photos")
        add_photo_folder(self.root, "one/Other", "other")
        add_photo_folder(self.root, "two/Other", "other")
        self.assertEqual([r.folder_name for r in self.find(throttle_limit=4)], ["Other", "Photos"])


class DuplicateFoldersInTheReportTests(TempRootTestCase):
    SETS = [DuplicateFolderSet("Photos", 12, 3, 123456, 2, ["C:\\one\\Photos", "D:\\two\\Photos"])]

    def sheet_parts(self, path):
        with zipfile.ZipFile(path) as archive:
            return archive.namelist()

    def test_writes_and_reads_back_a_duplicate_folders_sheet(self):
        path = self.path("report.xlsx")
        export_duplicate_report([], path, self.SETS)
        self.assertEqual(read_duplicate_folder_report(path), self.SETS)

    def test_writes_no_folder_sheet_unless_folder_sets_are_given(self):
        path = self.path("report.xlsx")
        export_duplicate_report([], path)
        self.assertEqual(read_duplicate_folder_report(path), [])
        self.assertNotIn("xl/worksheets/sheet2.xml", self.sheet_parts(path))

    def test_writes_an_empty_folder_sheet_when_no_duplicate_folders_were_found(self):
        path = self.path("report.xlsx")
        export_duplicate_report([], path, [])
        self.assertIn("xl/worksheets/sheet2.xml", self.sheet_parts(path))
        self.assertEqual(read_duplicate_folder_report(path), [])


class ValidatingDuplicateFoldersTests(TempRootTestCase):
    def folder_report(self):
        """Scan the tree (files and folders) into a report next to it and return its path."""
        files, records = folder_scan(self.root)
        path = self.root + ".xlsx"
        self.addCleanup(lambda: os.path.exists(path) and os.remove(path))
        export_duplicate_report(find_duplicate_files(files), path, find_duplicate_folders(files, records))
        return path

    def add_three_copies(self):
        for folder in ("one", "two", "three"):
            add_file(self.root, f"{folder}/Photos/a.jpg", "a")
            add_file(self.root, f"{folder}/Photos/sub/b.jpg", "b")

    def test_removes_a_folder_copy_that_no_longer_exists(self):
        self.add_three_copies()
        report = self.folder_report()
        shutil.rmtree(self.path("two/Photos"))
        result = validate_report(report)
        self.assertEqual(result.folder_copies_removed, 1)
        (row,) = read_duplicate_folder_report(report)
        self.assertEqual(row.count, 2)
        self.assertEqual(row.folders, [self.path("one/Photos"), self.path("three/Photos")])

    def test_removes_a_folder_copy_whose_contents_changed(self):
        self.add_three_copies()
        report = self.folder_report()
        add_file(self.root, "three/Photos/sub/new.jpg")
        result = validate_report(report)
        self.assertEqual((result.folder_copies_removed, result.folder_rows_remaining), (1, 1))

    def test_keeps_an_empty_folder_sheet_when_every_folder_row_is_removed(self):
        self.add_three_copies()
        report = self.folder_report()
        for folder in ("one", "two"):
            shutil.rmtree(self.path(f"{folder}/Photos"))
        result = validate_report(report)
        self.assertEqual(result.folder_rows_removed, 1)
        self.assertEqual(result.duplicate_folders, [])
        with zipfile.ZipFile(report) as archive:
            self.assertIn("xl/worksheets/sheet2.xml", archive.namelist())

    def test_keeps_folder_copies_that_cannot_be_reached(self):
        self.add_three_copies()
        report = self.folder_report()
        with mock.patch.object(validate, "check_folder_copy", return_value=validate.UNAVAILABLE), \
                self.assertLogs("find_duplicates", "WARNING"):
            result = validate_report(report)
        self.assertEqual((result.folder_copies_unavailable, result.folder_rows_remaining), (3, 1))

    def test_leaves_reports_without_a_folder_sheet_without_one(self):
        self.add_three_copies()
        report = self.root + ".xlsx"
        self.addCleanup(lambda: os.path.exists(report) and os.remove(report))
        export_duplicate_report(find_duplicate_files(list(iter_files(self.root))), report)
        os.remove(self.path("one/Photos/a.jpg"))
        result = validate_report(report)
        self.assertTrue(result.saved)
        self.assertIsNone(result.duplicate_folders)
        with zipfile.ZipFile(report) as archive:
            self.assertNotIn("xl/worksheets/sheet2.xml", archive.namelist())


class FolderCliTests(TempRootTestCase):
    def run_cli(self, *args):
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            return cli.main(list(args))

    def test_adds_duplicate_folders_to_the_report_with_folders(self):
        for folder in ("one", "two"):
            add_file(self.root, f"{folder}/Photos/a.jpg", "a")
            add_file(self.root, f"{folder}/Photos/b.jpg", "b")
        with tempfile.TemporaryDirectory() as out_dir:
            out = os.path.join(out_dir, "folders.xlsx")
            self.assertEqual(self.run_cli(self.root, out, "--folders"), 0)
            self.assertEqual([r.folder_name for r in read_duplicate_folder_report(out)], ["Photos"])
            self.assertEqual(len(read_duplicate_report(out)), 2)

    def test_re_checks_duplicate_folders_with_validate(self):
        for folder in ("one", "two", "three"):
            add_file(self.root, f"{folder}/Photos/a.jpg", "a")
        with tempfile.TemporaryDirectory() as out_dir:
            out = os.path.join(out_dir, "folders.xlsx")
            self.run_cli(self.root, out, "--folders")
            shutil.rmtree(self.path("one/Photos"))
            self.run_cli("--validate", out)
            self.assertEqual(read_duplicate_folder_report(out)[0].count, 2)

    def test_rejects_folders_with_validate(self):
        with self.assertRaises(SystemExit), contextlib.redirect_stderr(io.StringIO()):
            cli.main(["--validate", "--folders"])


if __name__ == "__main__":
    unittest.main()
