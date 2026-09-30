"""Tests for the Python port. Mirrors tests/DuplicateFinder.Tests.ps1 case for case."""

from __future__ import annotations

import contextlib
import io
import os
import stat
import sys
import tempfile
import time
import unittest
import zipfile
from datetime import datetime, timedelta, timezone
from types import SimpleNamespace
from unittest import mock
from xml.etree import ElementTree

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from find_duplicates import cli, matcher, scanner, validate  # noqa: E402
from find_duplicates.matcher import DuplicateSet, find_duplicate_files, md5_file  # noqa: E402
from find_duplicates.scanner import FileRecord, iter_files  # noqa: E402
from find_duplicates.validate import validate_report  # noqa: E402
from find_duplicates.xlsx import (  # noqa: E402
    column_name, excel_serial, export_duplicate_report, from_excel_serial, read_duplicate_report,
)
from tests.helpers import SAVED, add_file, read_worksheet  # noqa: E402


class TempDirTestCase(unittest.TestCase):
    def setUp(self) -> None:
        self._temp = tempfile.TemporaryDirectory()
        # Long form: the Windows temp folder is often a short (8.3) path, which the tool expands.
        self.root = scanner.full_path(self._temp.name)

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

    def test_leaves_out_excluded_files_given_in_another_form(self):
        keep = add_file(self.root, "keep.txt")
        skip = add_file(self.root, "sub/duplicates.xlsx")
        other_form = os.path.join(self.root, "sub", "..", "sub", "duplicates.xlsx")
        self.assertEqual([r.path for r in iter_files(self.root, exclude=[other_form])], [keep])
        self.assertTrue(os.path.exists(skip))

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

    @unittest.skipUnless(sys.platform == "win32", "short (8.3) names are Windows-only")
    def test_reports_long_folder_names_when_given_a_short_path(self):
        import ctypes

        add_file(self.root, "a long folder name/x.txt")
        long_root = scanner.full_path(self.root)  # the temp folder itself may be a short path
        buffer = ctypes.create_unicode_buffer(32_768)
        ctypes.windll.kernel32.GetShortPathNameW(long_root, buffer, len(buffer))
        if not buffer.value or buffer.value == long_root:
            self.skipTest("short names are disabled on this volume")
        (record,) = iter_files(buffer.value)
        self.assertEqual(record.folder, os.path.join(long_root, "a long folder name"))
        self.assertNotIn("~", record.folder)

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

    def test_includes_files_of_0_bytes_by_default(self):
        paths = [add_file(self.root, "a/empty.txt", ""), add_file(self.root, "b/empty.txt", "")]
        self.assertEqual(len(find_duplicate_files(self.records(*paths))), 1)

    def test_leaves_files_of_0_bytes_out_with_ignore_empty_files(self):
        paths = [add_file(self.root, "a/empty.txt", ""), add_file(self.root, "b/empty.txt", ""),
                 add_file(self.root, "a/full.txt"), add_file(self.root, "b/full.txt")]
        result = find_duplicate_files(self.records(*paths), ignore_empty_files=True)
        self.assertEqual([d.file_name for d in result], ["full.txt"])

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
                root = self.root = scanner.full_path(root)
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

    def test_writes_the_matching_rules_above_the_table(self):
        path = self.export()
        rows = read_worksheet(path, include_rules=True)
        self.assertEqual(rows[0][0], "Duplicate files")
        self.assertRegex(rows[1][0], "^A file is listed when another file has ALL of.*MD5")
        self.assertEqual(rows[4][0], "File Name", "rules, then one blank row (not written), then the table")
        with zipfile.ZipFile(path) as archive:
            sheet = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
        self.assertIn('<pane ySplit="6" topLeftCell="A7"', sheet)
        self.assertIn('<autoFilter ref="A6:H8"', sheet)

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


class ThrottleLimitTests(TempDirTestCase):
    """Hashing several files at a time (-ThrottleLimit / --throttle-limit)."""

    def setUp(self):
        super().setUp()
        for folder in ("a", "b", "c", "d"):
            add_file(self.root, f"{folder}/same.txt", "same")
            add_file(self.root, f"{folder}/split.txt", f"half {int(folder in ('a', 'b'))}")
        self.files = list(iter_files(self.root))
        self.sequential = find_duplicate_files(self.files)

    def test_finds_the_same_duplicates_hashing_several_at_a_time(self):
        for limit in (2, 8):
            with self.subTest(limit=limit):
                parallel = find_duplicate_files(self.files, throttle_limit=limit)
                self.assertEqual(len(parallel), 3)
                self.assertEqual(parallel, self.sequential)

    def test_reports_a_file_it_cannot_read_and_keeps_the_rest(self):
        for limit in (1, 4):
            with self.subTest(limit=limit), tempfile.TemporaryDirectory() as root:
                for folder in ("a", "b", "c"):
                    add_file(root, f"{folder}/x.txt")
                files = list(iter_files(root))
                os.remove(files[0].path)  # gone before it could be hashed
                with self.assertLogs("find_duplicates", "WARNING") as logs:
                    result = find_duplicate_files(files, throttle_limit=limit)
                self.assertEqual(len(logs.output), 1)
                self.assertIn(files[0].path, logs.output[0])
                self.assertEqual(result[0].folders, [f.folder for f in files[1:]])

    def test_rejects_a_throttle_limit_outside_1_to_64(self):
        for limit in (0, 65):
            with self.subTest(limit=limit), self.assertRaises(ValueError):
                find_duplicate_files([], throttle_limit=limit)


class ExcelDateTests(unittest.TestCase):
    def test_truncates_to_the_millisecond_like_dotnet(self):
        value = datetime(2024, 1, 2, 3, 4, 5, 678_999)
        self.assertEqual(from_excel_serial(excel_serial(value)), datetime(2024, 1, 2, 3, 4, 5, 678_000))


def write_excel_saved_workbook(path, rows):
    """A workbook shaped like one Excel has re-saved: shared strings, a renamed
    worksheet part, and cells without explicit types."""
    strings, sheet_rows = [], []
    for r, row in enumerate(rows, start=1):
        cells = []
        for c, value in enumerate(row, start=1):
            ref = f"{column_name(c)}{r}"
            if value is None:
                continue  # Excel leaves blank cells out
            if isinstance(value, str):
                strings.append(value)
                cells.append(f'<c r="{ref}" t="s"><v>{len(strings) - 1}</v></c>')
            else:
                cells.append(f'<c r="{ref}"><v>{value!r}</v></c>')
        sheet_rows.append(f'<row r="{r}">{"".join(cells)}</row>')
    main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    parts = {
        "[Content_Types].xml": '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
        "xl/workbook.xml": f'<workbook xmlns="{main}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                           '<sheets><sheet name="Duplicates" sheetId="1" r:id="rId7"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                                      '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/data.xml"/>'
                                      "</Relationships>",
        "xl/sharedStrings.xml": f'<sst xmlns="{main}">' + "".join(f"<si><t>{s}</t></si>" for s in strings) + "</sst>",
        "xl/worksheets/data.xml": f'<worksheet xmlns="{main}"><sheetData>{"".join(sheet_rows)}</sheetData></worksheet>',
    }
    with zipfile.ZipFile(path, "w") as archive:
        for name, content in parts.items():
            archive.writestr(name, content)


class ReadDuplicateReportTests(TempDirTestCase):
    ROUND_TRIP = [
        DuplicateSet("a & b.txt", datetime(2024, 1, 2, 3, 4, 5, 678_000), 1234, "AAAA", 3,
                     ["C:\\one", "C:\\two", "D:\\three"]),
        DuplicateSet("z.txt", datetime(2023, 6, 7, 8, 9, 10), 5, "BBBB", 2, ["C:\\x", "C:\\y"]),
    ]

    def test_reads_back_what_export_wrote(self):
        path = os.path.join(self.root, "report.xlsx")
        export_duplicate_report(self.ROUND_TRIP, path)
        self.assertEqual(read_duplicate_report(path), self.ROUND_TRIP)

    def test_reads_a_report_after_excel_has_saved_it(self):
        path = os.path.join(self.root, "excel.xlsx")
        write_excel_saved_workbook(path, [
            ["File Name", "Last Modified", "Size (bytes)", "MD5", "Copies", "Location 1", "Location 2"],
            ["x.txt", 45292.5, 10, "CCCC", 2, "C:\\a", "C:\\b"],
        ])
        (dup,) = read_duplicate_report(path)
        self.assertEqual(dup.file_name, "x.txt")
        self.assertEqual(dup.last_write_time, datetime(2024, 1, 1, 12, 0, 0))
        self.assertEqual(dup.size_bytes, 10)
        self.assertEqual(dup.folders, ["C:\\a", "C:\\b"])

    def test_rejects_a_workbook_that_is_not_a_duplicates_report(self):
        path = os.path.join(self.root, "other.xlsx")
        write_excel_saved_workbook(path, [["Name", "Amount"]])
        with self.assertRaisesRegex(ValueError, "is not a duplicates report"):
            read_duplicate_report(path)

    def test_rejects_a_file_that_is_not_a_workbook(self):
        path = add_file(self.root, "notes.xlsx", "not a zip")
        with self.assertRaisesRegex(ValueError, "is not an Excel workbook"):
            read_duplicate_report(path)

    def test_rejects_a_header_row_with_a_blank_cell(self):
        path = os.path.join(self.root, "gap.xlsx")
        write_excel_saved_workbook(path, [["File Name", None, "Size (bytes)", "MD5", "Copies"]])
        with self.assertRaisesRegex(ValueError, "is not a duplicates report"):
            read_duplicate_report(path)

    def test_reports_a_blank_number_cell_as_a_clear_error(self):
        path = os.path.join(self.root, "blank.xlsx")
        write_excel_saved_workbook(path, [
            ["File Name", "Last Modified", "Size (bytes)", "MD5", "Copies", "Location 1", "Location 2"],
            ["x.txt", None, 10, "CCCC", 2, "C:\\a", "C:\\b"],
        ])
        with self.assertRaisesRegex(ValueError, "could not be read as a duplicates report"):
            read_duplicate_report(path)

    def test_rejects_a_report_containing_a_dtd(self):
        path = os.path.join(self.root, "dtd.xlsx")
        export_duplicate_report(self.ROUND_TRIP, path)
        add_doctype_to_workbook_part(path)
        with self.assertRaisesRegex(ValueError, "DTD"):
            read_duplicate_report(path)

    def test_reports_a_damaged_workbook_as_an_error(self):
        path = os.path.join(self.root, "damaged.xlsx")
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("xl/workbook.xml", "<workbook")  # truncated XML
            archive.writestr("xl/_rels/workbook.xml.rels", "<Relationships/>")
        with self.assertRaisesRegex(ValueError, "could not be read"):
            read_duplicate_report(path)

    def test_reads_a_report_that_another_program_has_open(self):
        path = os.path.join(self.root, "open.xlsx")
        export_duplicate_report(self.ROUND_TRIP, path)
        with open(path, "r+b"):  # like Excel holding the file
            self.assertEqual(len(read_duplicate_report(path)), 2)

    def test_reports_a_missing_report(self):
        with self.assertRaisesRegex(FileNotFoundError, "was not found"):
            read_duplicate_report(os.path.join(self.root, "missing.xlsx"))


def add_doctype_to_workbook_part(path):
    """Rewrite the workbook part with a DTD, as a malicious file might."""
    with zipfile.ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    parts["xl/workbook.xml"] = parts["xl/workbook.xml"].replace(
        b"<workbook", b'<!DOCTYPE workbook [<!ENTITY x "x">]><workbook', 1)
    with zipfile.ZipFile(path, "w") as archive:
        for name, data in parts.items():
            archive.writestr(name, data)


class ValidateReportTests(TempDirTestCase):
    def scanned_report(self, root):
        """Scan root into a report next to it and return the report path."""
        path = root + ".xlsx"
        export_duplicate_report(find_duplicate_files(list(iter_files(root))), path)
        return path

    def tree(self, *relative, content="same content"):
        root = self.new_dir("tree")
        for rel in relative:
            add_file(root, rel, content)
        return root

    def test_leaves_the_report_untouched_when_every_copy_exists(self):
        root = self.tree("a/x.txt", "b/x.txt")
        report = self.scanned_report(root)
        before = os.stat(report).st_mtime_ns
        result = validate_report(report)
        self.assertFalse(result.saved)
        self.assertEqual((result.copies_checked, result.rows_remaining), (2, 1))
        self.assertEqual(os.stat(report).st_mtime_ns, before)

    def test_removes_a_copy_that_no_longer_exists(self):
        root = self.tree("a/x.txt", "b/x.txt", "c/x.txt")
        report = self.scanned_report(root)
        os.remove(os.path.join(root, "b", "x.txt"))
        result = validate_report(report)
        self.assertEqual(result.copies_removed, 1)
        self.assertTrue(result.saved)
        (dup,) = read_duplicate_report(report)
        self.assertEqual(dup.count, 2)
        self.assertEqual(dup.folders, [os.path.join(root, "a"), os.path.join(root, "c")])

    def test_removes_a_row_left_with_fewer_than_two_copies(self):
        root = self.new_dir("tree")
        for folder in ("a", "b"):
            add_file(root, f"{folder}/gone.txt", "gone")
            add_file(root, f"{folder}/kept.txt", "kept")
        report = self.scanned_report(root)
        os.remove(os.path.join(root, "a", "gone.txt"))
        result = validate_report(report)
        self.assertEqual(result.rows_removed, 1)
        self.assertEqual([d.file_name for d in read_duplicate_report(report)], ["kept.txt"])

    def test_removes_a_copy_that_was_changed_since_the_scan(self):
        root = self.tree("a/x.txt", "b/x.txt", "c/x.txt")
        report = self.scanned_report(root)
        add_file(root, "c/x.txt", "edited", SAVED + timedelta(hours=1))
        result = validate_report(report)
        self.assertEqual(result.copies_removed, 1)
        self.assertEqual(read_duplicate_report(report)[0].folders, [os.path.join(root, "a"), os.path.join(root, "b")])

    def test_finds_a_copy_whose_name_differs_only_by_case(self):
        root = self.tree("a/Photo.JPG", "b/photo.jpg")
        result = validate_report(self.scanned_report(root))
        self.assertEqual((result.copies_removed, result.rows_remaining), (0, 1))

    def test_keeps_copies_on_a_drive_or_share_that_cannot_be_reached(self):
        root = self.tree("a/x.txt", "b/x.txt")
        report = self.scanned_report(root)
        with mock.patch.object(validate, "check_copy", return_value=validate.UNAVAILABLE), \
                self.assertLogs("find_duplicates", "WARNING") as logs:
            result = validate_report(report)
        self.assertEqual(result.copies_unavailable, 2)
        self.assertFalse(result.saved)
        self.assertEqual(len(logs.output), 2)
        self.assertEqual(result.rows_remaining, 1)

    @unittest.skipUnless(sys.platform == "win32", "drive letters are Windows-only")
    def test_treats_a_copy_on_a_missing_drive_letter_as_unreachable(self):
        free = next((d for d in "QRSTUVWXYZ" if not os.path.exists(f"{d}:\\")), None)
        if free is None:
            self.skipTest("no free drive letter")
        state = validate.check_copy(f"{free}:\\photos", "x.txt", 1, datetime.now())
        self.assertEqual(state, validate.UNAVAILABLE)

    @unittest.skipIf(sys.platform == "win32", "time.tzset is not available on Windows")
    def test_keeps_a_copy_saved_in_the_repeated_hour_when_daylight_saving_ends(self):
        previous = os.environ.get("TZ")
        os.environ["TZ"] = "America/New_York"
        time.tzset()
        try:
            # 06:30 UTC on 3 Nov 2024 is 01:30 local, in the hour that happens twice.
            saved = datetime(2024, 11, 3, 6, 30, tzinfo=timezone.utc)
            path = add_file(self.root, "dst/x.txt", saved=saved)
            recorded = datetime.fromtimestamp(saved.timestamp())
            state = validate.check_copy(os.path.dirname(path), "x.txt", os.path.getsize(path), recorded)
            self.assertEqual(state, validate.PRESENT)
        finally:
            if previous is None:
                del os.environ["TZ"]
            else:
                os.environ["TZ"] = previous
            time.tzset()

    def test_checks_each_drive_or_share_only_once(self):
        root_cache = {}
        with mock.patch.object(validate, "_path_root", return_value="Z:\\"), \
                mock.patch.object(validate.os.path, "isdir", return_value=False) as isdir:
            states = {validate.check_copy(f"Z:\\{n}", "x.txt", 1, datetime.now(), root_cache) for n in range(5)}
        self.assertEqual(states, {validate.UNAVAILABLE})
        self.assertEqual(isdir.call_count, 1)

    def test_changes_nothing_with_dry_run(self):
        root = self.tree("a/x.txt", "b/x.txt", "c/x.txt")
        report = self.scanned_report(root)
        os.remove(os.path.join(root, "a", "x.txt"))
        result = validate_report(report, dry_run=True)
        self.assertEqual(result.copies_removed, 1)
        self.assertFalse(result.saved)
        self.assertEqual(read_duplicate_report(report)[0].count, 3)


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

    def test_hashes_several_files_at_a_time_with_throttle_limit(self):
        report = os.path.join(self.new_dir("work"), "parallel.xlsx")
        code, _ = self.run_cli(self.data, report, "--throttle-limit", "4")
        self.assertEqual(code, 0)
        self.assertEqual(read_duplicate_report(report)[0].count, 3)

    def test_leaves_files_of_0_bytes_out_with_ignore_empty_files(self):
        root = self.new_dir("empty")
        for folder in ("a", "b"):
            add_file(root, f"{folder}/empty.txt", "")
        report = os.path.join(self.new_dir("work"), "empty.xlsx")
        self.run_cli(root, report)
        self.assertEqual(len(read_duplicate_report(report)), 1)
        self.run_cli(root, report, "--ignore-empty-files")
        self.assertEqual(read_duplicate_report(report), [])

    def test_saves_no_report_with_dry_run(self):
        report = os.path.join(self.new_dir("work"), "dry.xlsx")
        code, out = self.run_cli(self.data, report, "--dry-run")
        self.assertEqual(code, 0)
        self.assertIn("Found 1 duplicated files", out)
        self.assertFalse(os.path.exists(report))

    def test_validates_duplicates_xlsx_in_the_current_folder(self):
        work = self.new_dir("work")
        with _chdir(work):
            self.run_cli(self.data)
            os.remove(os.path.join(self.data, "backup", "invoice.pdf"))
            code, out = self.run_cli("--validate")
        self.assertEqual(code, 0)
        self.assertIn("1 missing or changed", out)
        (dup,) = read_duplicate_report(os.path.join(work, "duplicates.xlsx"))
        self.assertEqual(dup.count, 2)

    def test_takes_the_report_to_validate_as_its_first_argument(self):
        report = os.path.join(self.new_dir("work"), "named")
        self.run_cli(self.data, report)
        os.remove(os.path.join(self.data, "2023", "invoice.pdf"))
        os.remove(os.path.join(self.data, "backup", "invoice.pdf"))
        self.run_cli("--validate", report)
        self.assertEqual(read_duplicate_report(report + ".xlsx"), [])

    def test_prints_every_folder_with_verbose(self):
        err = io.StringIO()
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(err):
            cli.main([self.data, os.path.join(self.new_dir("work"), "v.xlsx"), "--verbose"])
        self.assertIn(f"Scanning {os.path.join(self.data, 'backup')}", err.getvalue())

    def test_reports_a_damaged_report_without_a_traceback(self):
        path = os.path.join(self.new_dir("work"), "damaged.xlsx")
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("xl/workbook.xml", "<workbook")
            archive.writestr("xl/_rels/workbook.xml.rels", "<Relationships/>")
        err = io.StringIO()
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(err):
            code = cli.main(["--validate", path])
        self.assertEqual(code, 1)
        self.assertIn("error:", err.getvalue())
        self.assertNotIn("Traceback", err.getvalue())

    def test_rejects_scan_options_with_validate(self):
        with self.assertRaises(SystemExit), contextlib.redirect_stderr(io.StringIO()):
            cli.main(["--validate", "-j", "4"])

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
