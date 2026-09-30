"""Tests for the Python port. Mirrors tests/DuplicateFinder.Tests.ps1 case for case."""

from __future__ import annotations

import contextlib
import io
import os
import stat
import sys
import tempfile
import unittest
import zipfile
from datetime import datetime, timedelta
from types import SimpleNamespace
from unittest import mock
from xml.etree import ElementTree

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from find_duplicates import cli, matcher, scanner  # noqa: E402
from find_duplicates.matcher import DuplicateSet, find_duplicate_files, md5_file  # noqa: E402
from find_duplicates.scanner import FileRecord, iter_files  # noqa: E402
from find_duplicates.xlsx import column_name, export_duplicate_report  # noqa: E402
from tests.helpers import SAVED, add_file, read_worksheet  # noqa: E402


class TempDirTestCase(unittest.TestCase):
    def setUp(self) -> None:
        self._temp = tempfile.TemporaryDirectory()
        self.root = self._temp.name

    def tearDown(self) -> None:
        self._temp.cleanup()

    def new_dir(self, name: str) -> str:
        path = os.path.join(self.root, name)
        os.makedirs(path)
        return path

    def records(self, *paths: str):
        """FileRecords for the given paths, as the scanner would produce them."""
        by_path = {r.path: r for r in iter_files(self.root)}
        return [by_path[p] for p in paths]


class ColumnNameTests(unittest.TestCase):
    def test_converts_index_to_letters(self):
        for index, expected in [(1, "A"), (26, "Z"), (27, "AA"), (52, "AZ"),
                                (702, "ZZ"), (703, "AAA"), (16384, "XFD")]:
            with self.subTest(index=index):
                self.assertEqual(column_name(index), expected)

    def test_rejects_out_of_range(self):
        for index in (0, 16385):
            with self.subTest(index=index), self.assertRaises(ValueError):
                column_name(index)


class IterFilesTests(TempDirTestCase):
    def test_finds_files_in_every_sub_folder(self):
        add_file(self.root, "a.txt")
        add_file(self.root, "one/b.txt")
        add_file(self.root, "one/two/three/c.txt")
        self.assertEqual(sorted(r.name for r in iter_files(self.root)), ["a.txt", "b.txt", "c.txt"])

    def test_leaves_out_excluded_files(self):
        keep = add_file(self.root, "keep.txt")
        skip = add_file(self.root, "duplicates.xlsx")
        self.assertEqual([r.path for r in iter_files(self.root, exclude=[skip])], [keep])

    def test_lists_in_name_order_whatever_the_creation_order(self):
        for name in ("z", "a", "m"):
            add_file(self.root, f"{name}/{name}.txt")
        self.assertEqual([r.name for r in iter_files(self.root)], ["a.txt", "m.txt", "z.txt"])

    def test_handles_wildcard_characters_in_names(self):
        add_file(self.root, "[set]/file[1].txt")
        self.assertEqual([r.name for r in iter_files(self.root)], ["file[1].txt"])

    def test_does_not_follow_a_folder_link_that_loops_back(self):
        add_file(self.root, "sub/a.txt")
        try:
            os.symlink(self.root, os.path.join(self.root, "sub", "loop"), target_is_directory=True)
        except (OSError, NotImplementedError) as exc:
            self.skipTest(f"symbolic links cannot be created here: {exc}")
        self.assertEqual([r.name for r in iter_files(self.root)], ["a.txt"])

    def test_rejects_a_path_that_is_not_a_folder(self):
        path = add_file(self.root, "a.txt")
        with self.assertRaisesRegex(NotADirectoryError, "is not a folder"):
            list(iter_files(path))

    def test_reports_the_folder_being_scanned(self):
        add_file(self.root, "sub/a.txt")
        seen = []
        list(iter_files(self.root, on_folder=lambda folder, *_: seen.append(folder)))
        self.assertIn(os.path.join(self.root, "sub"), seen)

    def test_skips_an_unreadable_folder_with_a_warning(self):
        add_file(self.root, "ok/a.txt")
        add_file(self.root, "locked/b.txt")
        real_scandir = os.scandir

        def scandir(path):
            if os.path.basename(path) == "locked":
                raise PermissionError(13, "Permission denied")
            return real_scandir(path)

        with mock.patch.object(scanner.os, "scandir", scandir), self.assertLogs("find_duplicates", "WARNING") as logs:
            names = [r.name for r in iter_files(self.root)]
        self.assertEqual(names, ["a.txt"])
        self.assertIn("locked", logs.output[0])


class FolderAndCloudDetectionTests(unittest.TestCase):
    @staticmethod
    def entry(symlink=False, junction=None):
        entry = SimpleNamespace(is_symlink=lambda: symlink)
        if junction is not None:
            entry.is_junction = lambda: junction
        return entry

    def test_follows_plain_and_cloud_synced_folders(self):
        self.assertFalse(scanner.is_folder_link(self.entry(junction=False)))

    def test_does_not_follow_symlinks_or_junctions(self):
        self.assertTrue(scanner.is_folder_link(self.entry(symlink=True)))
        self.assertTrue(scanner.is_folder_link(self.entry(junction=True)))

    def test_online_only_attributes(self):
        cases = [
            (0x20, False),      # Archive: a normal local file
            (0x420, False),     # Archive + ReparsePoint: pinned / locally available
            (0x1020, True),     # Offline
            (0x40020, True),    # RecallOnOpen
            (0x400420, True),   # RecallOnDataAccess (OneDrive Files On-Demand)
        ]
        for attributes, expected in cases:
            with self.subTest(attributes=hex(attributes)):
                record = FileRecord("p", "n", "f", 1, 0, attributes)
                self.assertEqual(scanner.is_cloud_only(record), expected)


class FindDuplicateFilesTests(TempDirTestCase):
    def test_records_every_copy_across_many_folders(self):
        paths = [add_file(self.root, p) for p in ("a/report.doc", "b/report.doc", "b/c/d/report.doc")]
        result = find_duplicate_files(self.records(*paths))
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0].file_name, "report.doc")
        self.assertEqual(result[0].count, 3)
        self.assertEqual(result[0].folders, sorted(os.path.dirname(p) for p in paths))
        self.assertEqual(result[0].md5, md5_file(paths[0]))

    def test_md5_matches_known_value(self):
        path = add_file(self.root, "x.txt", content="abc")
        self.assertEqual(md5_file(path), "900150983CD24FB0D6963F7D28E17F72")

    def test_ignores_files_whose_contents_differ(self):
        paths = [add_file(self.root, "a/x.txt", "aaaa"), add_file(self.root, "b/x.txt", "bbbb")]
        self.assertEqual(find_duplicate_files(self.records(*paths)), [])

    def test_ignores_files_whose_saved_dates_differ(self):
        paths = [add_file(self.root, "a/x.txt"),
                 add_file(self.root, "b/x.txt", saved=SAVED + timedelta(minutes=1))]
        self.assertEqual(find_duplicate_files(self.records(*paths)), [])

    def test_ignores_files_whose_names_differ(self):
        paths = [add_file(self.root, "a/x.txt"), add_file(self.root, "b/y.txt")]
        self.assertEqual(find_duplicate_files(self.records(*paths)), [])

    def test_names_differing_only_by_case_match(self):
        paths = [add_file(self.root, "a/Photo.JPG"), add_file(self.root, "b/photo.jpg")]
        self.assertEqual(len(find_duplicate_files(self.records(*paths))), 1)

    def test_names_that_differ_beyond_case_do_not_match(self):
        paths = [add_file(self.root, "a/stra" + chr(0xDF) + "e.txt"), add_file(self.root, "b/STRASSE.txt")]
        self.assertEqual(find_duplicate_files(self.records(*paths)), [])

    def test_ignores_sub_second_differences(self):
        paths = [add_file(self.root, "a/x.txt"),
                 add_file(self.root, "b/x.txt", saved=SAVED + timedelta(milliseconds=400))]
        self.assertEqual(len(find_duplicate_files(self.records(*paths))), 1)

    def test_splits_same_name_same_date_files_by_content(self):
        paths = [add_file(self.root, "a/x.txt", "first"), add_file(self.root, "b/x.txt", "first"),
                 add_file(self.root, "c/x.txt", "other"), add_file(self.root, "d/x.txt", "other"),
                 add_file(self.root, "e/x.txt", "unique!")]
        result = find_duplicate_files(self.records(*paths))
        self.assertEqual([s.count for s in result], [2, 2])
        self.assertEqual(sorted(f for s in result for f in s.folders),
                         sorted(os.path.dirname(p) for p in paths[:4]))

    def test_names_a_duplicate_after_the_copy_in_the_first_folder(self):
        add_file(self.root, "b/photo.jpg")
        add_file(self.root, "a/Photo.JPG")
        result = find_duplicate_files(list(iter_files(self.root)))
        self.assertEqual(result[0].file_name, "Photo.JPG")

    def test_returns_nothing_for_an_empty_list(self):
        self.assertEqual(find_duplicate_files([]), [])

    def test_md5_only_when_name_date_and_size_match(self):
        cases = {
            "different names": [("a/x.txt", "c", SAVED), ("b/y.txt", "c", SAVED)],
            "different dates": [("a/x.txt", "c", SAVED), ("b/x.txt", "c", SAVED + timedelta(days=1))],
            "different sizes": [("a/x.txt", "short", SAVED), ("b/x.txt", "much longer", SAVED)],
        }
        for label, files in cases.items():
            with self.subTest(label), tempfile.TemporaryDirectory() as root:
                self.root = root
                paths = [add_file(root, p, c, s) for p, c, s in files]
                with mock.patch.object(matcher, "md5_file", return_value="ABC") as md5:
                    find_duplicate_files(self.records(*paths))
                md5.assert_not_called()

    def test_hashes_only_the_matching_candidates(self):
        paths = [add_file(self.root, "a/x.txt"), add_file(self.root, "b/x.txt"), add_file(self.root, "c/y.txt")]
        with mock.patch.object(matcher, "md5_file", return_value="ABC") as md5:
            find_duplicate_files(self.records(*paths))
        self.assertEqual(md5.call_count, 2)

    def test_downloads_online_only_files_by_default(self):
        paths = [add_file(self.root, "cloud/x.txt"), add_file(self.root, "local/x.txt")]
        with mock.patch.object(matcher, "is_cloud_only", return_value=True):
            self.assertEqual(len(find_duplicate_files(self.records(*paths))), 1)

    def test_skip_cloud_only_does_not_download(self):
        paths = [add_file(self.root, "cloud/x.txt"), add_file(self.root, "local1/x.txt"),
                 add_file(self.root, "local2/x.txt")]
        with mock.patch.object(matcher, "is_cloud_only", side_effect=lambda r: "cloud" in r.path), \
                mock.patch.object(matcher, "md5_file", return_value="SAME") as md5, \
                self.assertLogs("find_duplicates", "WARNING") as logs:
            result = find_duplicate_files(self.records(*paths), skip_cloud_only=True)
        self.assertEqual(md5.call_count, 2)
        self.assertFalse(any("cloud" in c.args[0] for c in md5.call_args_list))
        self.assertEqual(result[0].folders, sorted(os.path.dirname(p) for p in paths[1:]))
        self.assertIn("1 online-only", logs.output[0])

    def test_skips_a_file_it_cannot_hash_and_keeps_the_rest(self):
        paths = [add_file(self.root, "locked/x.txt"), add_file(self.root, "b/x.txt"), add_file(self.root, "c/x.txt")]

        def fake_md5(path):
            if "locked" in path:
                raise PermissionError(13, "file is locked")
            return "SAME"

        with mock.patch.object(matcher, "md5_file", side_effect=fake_md5), \
                self.assertLogs("find_duplicates", "WARNING") as logs:
            result = find_duplicate_files(self.records(*paths))
        self.assertEqual(len(logs.output), 1)
        self.assertIn("locked", logs.output[0])
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0].folders, sorted(os.path.dirname(p) for p in paths[1:]))


class ExportDuplicateReportTests(TempDirTestCase):
    SETS = [
        DuplicateSet("a & b <1>.txt", datetime(2024, 1, 2, 3, 4, 5), 1234, "AAAA", 3,
                     ["C:\\one", "C:\\two & more", "D:\\three"]),
        DuplicateSet("z.txt", datetime(2023, 6, 7, 8, 9, 10), 5, "BBBB", 2,
                     ["C:\\x", "C:\\bad" + chr(1) + "name"]),
    ]

    def export(self, name="report.xlsx", sets=None):
        path = os.path.join(self.root, name)
        export_duplicate_report(self.SETS if sets is None else sets, path)
        return path

    def test_one_row_per_file_with_a_column_per_location(self):
        rows = read_worksheet(self.export())
        self.assertEqual(len(rows), 3)
        self.assertEqual(rows[0], ["File Name", "Last Modified", "Size (bytes)", "MD5", "Copies",
                                   "Location 1", "Location 2", "Location 3"])
        self.assertEqual(rows[1][0], "a & b <1>.txt")
        self.assertEqual(rows[1][2:5], ["1234", "AAAA", "3"])
        self.assertEqual(rows[1][5:8], ["C:\\one", "C:\\two & more", "D:\\three"])
        self.assertEqual(rows[2][5:7], ["C:\\x", "C:\\bad" + chr(0xFFFD) + "name"])

    def test_stores_the_saved_date_as_a_real_excel_date(self):
        serial = float(read_worksheet(self.export())[1][1])
        self.assertEqual(datetime(1899, 12, 30) + timedelta(days=serial), self.SETS[0].last_write_time)

    def test_package_has_every_required_part_as_valid_xml(self):
        with zipfile.ZipFile(self.export()) as archive:
            self.assertEqual(sorted(archive.namelist()), sorted([
                "[Content_Types].xml", "_rels/.rels", "xl/_rels/workbook.xml.rels",
                "xl/styles.xml", "xl/workbook.xml", "xl/worksheets/sheet1.xml"]))
            for name in archive.namelist():
                with self.subTest(part=name):
                    ElementTree.fromstring(archive.read(name))

    def test_replaces_unpaired_surrogates(self):
        # Undecodable bytes in Linux file names arrive as lone surrogates.
        bad = DuplicateSet("f" + chr(0xDC80) + ".txt", datetime(2024, 1, 1), 1, "A", 2, ["/x", "/y"])
        rows = read_worksheet(self.export(sets=[bad]))
        self.assertEqual(rows[1][0], "f" + chr(0xFFFD) + ".txt")

    def test_header_only_workbook_when_no_duplicates(self):
        rows = read_worksheet(self.export("empty.xlsx", sets=[]))
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0][-1], "Location 1")

    def test_overwrites_existing_report_and_leaves_no_temp_files(self):
        path = os.path.join(self.root, "report.xlsx")
        with open(path, "w") as stream:
            stream.write("old")
        export_duplicate_report(self.SETS, path)
        self.assertEqual(len(read_worksheet(path)), 3)
        self.assertEqual(os.listdir(self.root), ["report.xlsx"])

    def test_resolves_a_relative_path_against_the_current_folder(self):
        with contextlib.chdir(self.root) if hasattr(contextlib, "chdir") else _chdir(self.root):
            export_duplicate_report(self.SETS, "relative.xlsx")
        self.assertTrue(os.path.exists(os.path.join(self.root, "relative.xlsx")))


@contextlib.contextmanager
def _chdir(path):
    previous = os.getcwd()
    os.chdir(path)
    try:
        yield
    finally:
        os.chdir(previous)


class CliTests(TempDirTestCase):
    def setUp(self):
        super().setUp()
        self.data = os.path.join(self.root, "data")
        add_file(self.data, "2023/invoice.pdf", "invoice")
        add_file(self.data, "backup/invoice.pdf", "invoice")
        add_file(self.data, "old/copy/invoice.pdf", "invoice")
        add_file(self.data, "notes.txt", "one")
        add_file(self.data, "other/notes.txt", "two")

    def run_cli(self, *args):
        out = io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(io.StringIO()):
            code = cli.main(list(args))
        return code, out.getvalue()

    def test_saves_to_duplicates_xlsx_in_the_current_folder_by_default(self):
        work = self.new_dir("work")
        with _chdir(work):
            code, out = self.run_cli(self.data)
        self.assertEqual(code, 0)
        report = os.path.join(work, "duplicates.xlsx")
        rows = read_worksheet(report)
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[1][0], "invoice.pdf")
        self.assertEqual(rows[1][4], "3")
        self.assertEqual(rows[1][5:8], [os.path.join(self.data, p) for p in ("2023", "backup", "old/copy".replace("/", os.sep))])
        self.assertIn("Found 1 duplicated files (3 copies in total).", out)

    def test_uses_the_given_output_name_and_adds_xlsx(self):
        work = self.new_dir("work")
        self.run_cli(self.data, os.path.join(work, "my-report"))
        self.assertTrue(os.path.exists(os.path.join(work, "my-report.xlsx")))
        self.run_cli(self.data, "-o", os.path.join(work, "named"))
        self.assertTrue(os.path.exists(os.path.join(work, "named.xlsx")))

    def test_does_not_scan_its_own_report(self):
        report = os.path.join(self.data, "duplicates.xlsx")
        self.run_cli(self.data, report)
        self.run_cli(self.data, report)
        rows = read_worksheet(report)
        self.assertEqual([r[0] for r in rows[1:]], ["invoice.pdf"])

    def test_rejects_a_missing_folder(self):
        code, _ = self.run_cli(os.path.join(self.root, "missing"))
        self.assertEqual(code, 1)


if __name__ == "__main__":
    unittest.main()
