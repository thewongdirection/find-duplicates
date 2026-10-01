"""Tests for the Python port. Mirrors tests/DuplicateFinder.Tests.ps1 case for case."""

from __future__ import annotations

import contextlib
import dataclasses
import io
import logging
import os
import shutil
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
from xml.sax.saxutils import escape

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from find_duplicates import cli, matcher, scanner, validate  # noqa: E402
from find_duplicates.folders import find_duplicate_folders  # noqa: E402
from find_duplicates.matcher import DuplicateSet, find_duplicate_files, md5_file  # noqa: E402
from find_duplicates.names import path_depth  # noqa: E402
from find_duplicates.scanner import FileRecord, iter_files, local_time, utc_offset  # noqa: E402
from find_duplicates.validate import validate_report  # noqa: E402
from find_duplicates.xlsx import (  # noqa: E402
    ReportLockedError, ScanSettings, column_name, excel_serial, export_duplicate_report, from_excel_serial, parse_utc_offset,
    read_duplicate_report, read_duplicate_workbook, utc_offset_text,
)
from tests.helpers import (  # noqa: E402
    SAVED, add_file, as_if_on_a_network_drive, locked_like_excel, locked_message, read_worksheet, sheet_names,
    time_zone,
)


MINIMUM_SIZE_LABEL = (
    "Smallest file listed, in bytes (-MinimumSize or -IgnoreEmptyFiles; Python: --minimum-size or --ignore-empty-files)"
)
EXCLUDE_NAMES_LABEL = "Names left out (-Exclude; Python: --exclude)"


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

    def test_leaves_out_an_excluded_file_when_the_folder_path_ends_in_a_separator(self):
        keep = add_file(self.root, "keep.txt")
        skip = add_file(self.root, "duplicates.xlsx")
        self.assertEqual([r.path for r in iter_files(self.root + os.sep, exclude=[skip])], [keep])

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

    def test_lists_several_folders_at_a_time_and_returns_the_same_files_in_the_same_order(self):
        for path in ("b/x.txt", "a/y.txt", "a/deep/er/z.txt", "c/w.txt", "a/b.txt", "top.txt"):
            add_file(self.root, path)
        os.makedirs(os.path.join(self.root, "a", "empty"))
        with contextlib.suppress(OSError, NotImplementedError):  # the rest is still worth checking
            os.symlink(self.root, os.path.join(self.root, "a", "loop"), target_is_directory=True)
        one, many = [], []

        sequential = [f.path for f in iter_files(self.root, folders=one)]
        with as_if_on_a_network_drive():
            parallel = [f.path for f in iter_files(self.root, folders=many, throttle_limit=4)]

        self.assertEqual(parallel, sequential)
        self.assertEqual(many, one)

    def test_leaves_out_files_and_folders_whose_names_match_exclude_names_ignoring_case(self):
        for path in ("keep.txt", "Thumbs.db", "a.TMP", ".git/obj", "sub/.Git/x", "sub/keep2.txt"):
            add_file(self.root, path)
        folders = []

        found = list(iter_files(self.root, exclude_names=["thumbs.db", "*.tmp", ".GIT"], folders=folders))

        self.assertEqual([f.name for f in found], ["keep.txt", "keep2.txt"])
        self.assertFalse([f for f in folders if "git" in f.path.lower()], "left-out folders are not listed")
        self.assertTrue(all(f.readable for f in folders), "leaving names out does not make a folder unreadable")

    def test_treats_characters_other_than_star_and_question_mark_in_exclude_names_literally(self):
        for name in ("a.b", "aXb", "[ab]", "a", "file(1).txt", "x+y"):
            add_file(self.root, name)

        found = list(iter_files(self.root, exclude_names=["a.b", "[ab]", "?", "x+y"]))

        self.assertEqual(sorted(f.name for f in found), ["aXb", "file(1).txt"])

    def test_matches_exclude_names_patterns_against_the_whole_name(self):
        for name in ("ab", "x.tmp.bak", "y.tmp", "abab", "aba"):
            add_file(self.root, name)

        found = list(iter_files(self.root, exclude_names=["a", "*.tmp", "*ab*ab"]))

        self.assertEqual(sorted(f.name for f in found), ["ab", "aba", "x.tmp.bak"])

    def test_matches_a_pattern_with_many_stars_quickly(self):
        add_file(self.root, "a" * 200)
        started = time.monotonic()

        found = list(iter_files(self.root, exclude_names=["*a*a*a*a*a*a*a*a*b"]))

        self.assertEqual(len(found), 1)
        self.assertLess(time.monotonic() - started, 5, "patterns must never backtrack exponentially")

    def test_rejects_an_exclude_name_pattern_holding_a_slash_or_backslash(self):
        for pattern in ("photos/raw", "photos\\raw"):
            with self.subTest(pattern=pattern), self.assertRaisesRegex(ValueError, "not paths"):
                list(iter_files(self.root, exclude_names=[pattern]))

    def test_leaves_out_a_folder_link_whose_name_matches_exclude_names_without_marking_its_folder_unreadable(self):
        add_file(self.root, "a/x.txt")
        try:
            os.symlink(self.root, os.path.join(self.root, "a", ".git"), target_is_directory=True)
        except (OSError, NotImplementedError) as exc:
            self.skipTest(f"symbolic links cannot be created here: {exc}")
        folders = []
        list(iter_files(self.root, exclude_names=[".git"], folders=folders))
        self.assertTrue(all(f.readable for f in folders))

    def test_matches_exclude_names_patterns_in_any_unicode_form_and_a_character_beyond_u_ffff_with_one_question_mark(self):
        composed = "caf\u00e9.txt"  # e-acute as one character
        decomposed = "CAFE\u0301.*"  # E + combining acute, upper case
        emoji = "x\U0001F600.txt"
        for name in (composed, emoji, "keep.txt"):
            add_file(self.root, name)

        found = list(iter_files(self.root, exclude_names=[decomposed, "x?.txt"]))

        self.assertEqual([f.name for f in found], ["keep.txt"])

    def test_does_not_match_a_character_beyond_u_ffff_with_two_question_marks(self):
        add_file(self.root, "x\U0001F600.txt")
        self.assertEqual(len(list(iter_files(self.root, exclude_names=["x??.txt"]))), 1)

    def test_leaves_out_the_same_names_listing_several_folders_at_a_time(self):
        for path in ("a/x.txt", "a/cache/y.txt", "b/cache/deep/z.txt", "b/w.tmp", "c/v.txt"):
            add_file(self.root, path)
        one, many = [], []

        sequential = [f.path for f in iter_files(self.root, exclude_names=["cache", "*.tmp"], folders=one)]
        with as_if_on_a_network_drive():
            parallel = [f.path for f in iter_files(self.root, exclude_names=["cache", "*.tmp"], folders=many, throttle_limit=4)]

        self.assertEqual([os.path.basename(p) for p in sequential], ["x.txt", "v.txt"])
        self.assertEqual(parallel, sequential)
        self.assertEqual([f.path for f in many], [f.path for f in one])

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

    def test_treats_a_local_folder_as_not_on_a_network_drive(self):
        with tempfile.TemporaryDirectory() as root:
            self.assertFalse(scanner.on_network_drive(root))

    def test_uses_4_at_a_time_by_default_on_a_network_drive_and_1_elsewhere(self):
        with mock.patch.object(scanner, "on_network_drive", return_value=True):
            self.assertEqual(scanner.default_throttle_limit("."), 4)
        with mock.patch.object(scanner, "on_network_drive", return_value=False):
            self.assertEqual(scanner.default_throttle_limit("."), 1)

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

    def test_computes_the_md5_of_the_start_of_a_file(self):
        path = add_file(self.root, "abcdef.txt", content="abcdef")
        self.assertEqual(md5_file(path, 3), "900150983CD24FB0D6963F7D28E17F72")

    def add_large_file(self, relative, at, value):
        """A 3 MB file of zeros with one byte set, saved at the same date as the others."""
        path = os.path.join(self.root, *relative.split("/"))
        os.makedirs(os.path.dirname(path), exist_ok=True)
        content = bytearray(3 << 20)
        content[at] = value
        with open(path, "wb") as stream:
            stream.write(content)
        ns = int(SAVED.timestamp()) * 1_000_000_000
        os.utime(path, ns=(ns, ns))

    def test_reads_large_files_in_full_only_when_their_start_is_the_same(self):
        self.add_large_file("a/same.bin", 0, 1)
        self.add_large_file("b/same.bin", 0, 1)
        self.add_large_file("a/start.bin", 0, 1)
        self.add_large_file("b/start.bin", 0, 2)
        self.add_large_file("a/end.bin", 2 << 20, 1)
        self.add_large_file("b/end.bin", 2 << 20, 2)
        for limit in (1, 4):
            with self.subTest(limit=limit), as_if_on_a_network_drive(), \
                    mock.patch.object(matcher, "FIRST_BYTES_MIN_SIZE", 2 << 20):  # large: 2 MB, to keep the files small
                cache = {}
                result = find_duplicate_files(list(iter_files(self.root)), md5_cache=cache, throttle_limit=limit)
                self.assertEqual([d.file_name for d in result], ["same.bin"])
                self.assertEqual(sorted(os.path.basename(p) for p in cache), ["end.bin", "end.bin", "same.bin", "same.bin"],
                                 "files that differ at the start are never read in full")

    def test_does_not_compare_the_start_of_large_files_the_previous_report_already_hashed(self):
        self.add_large_file("a/start.bin", 0, 1)
        self.add_large_file("b/start.bin", 0, 2)
        files = list(iter_files(self.root))
        cache = {f.path: "0123456789ABCDEF0123456789ABCDEF" for f in files}
        with mock.patch.object(matcher, "FIRST_BYTES_MIN_SIZE", 2 << 20):
            result = find_duplicate_files(files, md5_cache=cache)
        self.assertEqual([d.file_name for d in result], ["start.bin"], "the recorded hashes are trusted")

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

    def test_leaves_files_smaller_than_minimum_size_out(self):
        paths = []
        for folder in ("a", "b"):
            paths += [add_file(self.root, f"{folder}/small.txt", "s" * 9), add_file(self.root, f"{folder}/exact.txt", "e" * 10),
                      add_file(self.root, f"{folder}/large.txt", "l" * 11)]
        result = find_duplicate_files(self.records(*paths), minimum_size=10)
        self.assertEqual([d.file_name for d in result], ["exact.txt", "large.txt"])

    def test_orders_each_rows_locations_from_the_least_to_the_most_nested(self):
        paths = [add_file(self.root, f"{folder}/x.txt") for folder in ("z", "a/b/c", "a/b", "B", "a/C/c")]

        (result,) = find_duplicate_files(self.records(*paths))

        expected = [os.path.join(self.root, *folder.split("/")) for folder in ("B", "z", "a/b", "a/b/c", "a/C/c")]
        self.assertEqual(result.folders, expected, "equally deep folders keep their alphabetical order")

    def test_counts_how_many_folders_deep_a_path_is(self):
        self.assertEqual(path_depth("C:\\a\\b"), 3)
        self.assertEqual(path_depth("/home/a/"), 2)
        self.assertEqual(path_depth("\\\\server\\share\\a"), 3)

    def test_returns_nothing_for_an_empty_list(self):
        self.assertEqual(find_duplicate_files([]), [])

    def test_never_reads_the_size_or_saved_date_of_a_file_whose_name_no_other_file_has(self):
        paths = [add_file(self.root, "a/x.txt"), add_file(self.root, "b/x.txt")]
        # Never created: reading its size or saved date would fail with a warning.
        unique = FileRecord(os.path.join(self.root, "c", "unique.txt"), "unique.txt", os.path.join(self.root, "c"))

        with mock.patch.object(matcher.log, "warning") as warning:
            result = find_duplicate_files(self.records(*paths) + [unique])

        warning.assert_not_called()
        self.assertEqual(len(result), 1)

    def test_skips_a_file_that_has_gone_since_the_scan_with_a_warning(self):
        paths = [add_file(self.root, "a/x.txt"), add_file(self.root, "b/x.txt")]
        # Listed by the scan, deleted before it was compared.
        gone = FileRecord(os.path.join(self.root, "c", "x.txt"), "x.txt", os.path.join(self.root, "c"))

        with self.assertLogs("find_duplicates", "WARNING") as logs:
            result = find_duplicate_files(self.records(*paths) + [gone])

        self.assertIn(f"Skipping '{gone.path}'", logs.output[0])
        self.assertEqual(result[0].count, 2)

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
                     ["C:\\one", "C:\\two & more", "D:\\three"], utc_offset=timedelta(minutes=330)),
        DuplicateSet("z.txt", datetime(2023, 6, 7, 8, 9, 10), 5, "BBBB", 2,
                     ["C:\\x", "C:\\bad" + chr(1) + "name"]),
    ]

    def export(self, name="report.xlsx", sets=None):
        path = os.path.join(self.root, name)
        export_duplicate_report(self.SETS if sets is None else sets, path)
        return path

    @unittest.skipUnless(sys.platform == "win32", "other programs lock files against writing only on Windows")
    def test_says_the_report_is_locked_when_another_program_holds_it_open(self):
        report = self.export("locked.xlsx")
        with open(report, "rb") as stream:
            before = stream.read()

        with locked_like_excel(report), self.assertRaises(ReportLockedError) as caught:
            export_duplicate_report([], report)

        self.assertEqual(str(caught.exception), locked_message(report))
        with open(report, "rb") as stream:
            self.assertEqual(stream.read(), before, "the report is left as it was")
        self.assertEqual([name for name in os.listdir(self.root) if name.endswith(".tmp")], [])

    def test_one_row_per_file_with_a_column_per_location(self):
        rows = read_worksheet(self.export())
        self.assertEqual(len(rows), 3)
        self.assertEqual(rows[0], ["File Name", "Last Modified", "UTC Offset", "Size (bytes)", "MD5", "Copies",
                                   "Location 1", "Location 2", "Location 3"])
        self.assertEqual(rows[1][0], "a & b <1>.txt")
        self.assertEqual(rows[1][2:6], ["+05:30", "1234", "AAAA", "3"])
        self.assertEqual(rows[1][6:9], ["C:\\one", "C:\\two & more", "D:\\three"])
        self.assertRegex(rows[2][2], r"^[+-]\d{2}:\d{2}", "a set without an offset gets this computer's")
        self.assertEqual(rows[2][6:8], ["C:\\x", "C:\\bad" + chr(0xFFFD) + "name"])

    def test_stores_the_saved_date_as_a_real_excel_date(self):
        serial = float(read_worksheet(self.export())[1][1])
        self.assertEqual(datetime(1899, 12, 30) + timedelta(days=serial), self.SETS[0].last_write_time)

    def test_package_has_every_required_part_as_valid_xml(self):
        with zipfile.ZipFile(self.export()) as archive:
            self.assertEqual(sorted(archive.namelist()), sorted([
                "[Content_Types].xml", "_rels/.rels", "xl/_rels/workbook.xml.rels",
                "xl/styles.xml", "xl/workbook.xml", "xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"]))
            for name in archive.namelist():
                with self.subTest(part=name):
                    ElementTree.fromstring(archive.read(name))

    def test_replaces_unpaired_surrogates(self):
        # Undecodable bytes in Linux file names arrive as lone surrogates.
        bad = DuplicateSet("f" + chr(0xDC80) + ".txt", datetime(2024, 1, 1), 1, "A", 2, ["/x", "/y"])
        rows = read_worksheet(self.export(sets=[bad]))
        self.assertEqual(rows[1][0], "f" + chr(0xFFFD) + ".txt")

    def test_writes_the_matching_rules_on_a_rules_sheet(self):
        path = self.export()
        self.assertEqual(sheet_names(path), ["Duplicates", "Rules"])
        rules = [row[0] for row in read_worksheet(path, "xl/worksheets/sheet2.xml", all_rows=True)]
        self.assertEqual(rules[0], "Matching rules")
        self.assertIn("Sheet 'Duplicates': duplicate files", rules)
        self.assertTrue(any(line.startswith("A file is listed when another file has ALL of") for line in rules))
        self.assertNotIn("Sheet 'Duplicate Folders': duplicate folders", rules, "no folder sheet, no folder rules")
        self.assertNotIn("Scan settings", rules, "the scan left nothing out")

    def test_records_the_scan_settings_on_the_rules_sheet(self):
        path = os.path.join(self.root, "report.xlsx")
        export_duplicate_report(self.SETS, path, settings=ScanSettings(["*.tmp", "Thumbs.db"], 1024))

        rows = read_worksheet(path, "xl/worksheets/sheet2.xml", all_rows=True)
        self.assertIn("Scan settings", [row[0] for row in rows])
        (exclude,) = [row for row in rows if row[0] == "Names left out (-Exclude; Python: --exclude)"]
        self.assertEqual(exclude[1:3], ["*.tmp", "Thumbs.db"])
        (size,) = [row for row in rows if row[0] == MINIMUM_SIZE_LABEL]
        self.assertEqual(size[1], "1024")

    def test_starts_the_table_on_row_1(self):
        path = self.export()
        self.assertEqual(read_worksheet(path, all_rows=True)[0][0], "File Name")
        with zipfile.ZipFile(path) as archive:
            sheet = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
        self.assertIn('<pane ySplit="1" topLeftCell="A2"', sheet)
        self.assertIn('<autoFilter ref="A1:I3"', sheet)

    def test_reads_a_report_that_had_the_rules_above_the_table(self):
        path = os.path.join(self.root, "older.xlsx")
        write_excel_saved_workbook(path, [
            ["Duplicate files"], ["Some rule."], [],
            ["File Name", "Last Modified", "Size (bytes)", "MD5", "Copies", "Location 1", "Location 2"],
            ["x.txt", 45292.5, 10, "CCCC", 2, "C:\\a", "C:\\b"],
        ])
        self.assertEqual([d.file_name for d in read_duplicate_report(path)], ["x.txt"])

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
            with self.subTest(limit=limit), as_if_on_a_network_drive():
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
                with as_if_on_a_network_drive(), self.assertLogs("find_duplicates", "WARNING") as logs:
                    result = find_duplicate_files(files, throttle_limit=limit)
                self.assertEqual(len(logs.output), 1)
                self.assertIn(files[0].path, logs.output[0])
                self.assertEqual(result[0].folders, [f.folder for f in files[1:]])

    def test_rejects_a_throttle_limit_outside_1_to_64(self):
        for limit in (0, 65):
            with self.subTest(limit=limit), self.assertRaises(ValueError):
                find_duplicate_files([], throttle_limit=limit)


class ThreadUseTests(TempDirTestCase):
    """Python-specific: threads (-j) are used only where they are faster (see
    scanner.on_network_drive)."""

    def test_hashes_on_threads_only_files_on_a_network_drive_or_of_1_mb_or_more(self):
        path = add_file(self.root, "a/x.txt")
        big = matcher.PARALLEL_HASH_MIN_BYTES
        self.assertFalse(matcher._worth_a_thread(path, {path: 10}))
        self.assertTrue(matcher._worth_a_thread(path, {path: big}))
        self.assertTrue(matcher._worth_a_thread(path, None), "without sizes, as md5_map always did")
        with as_if_on_a_network_drive():
            self.assertTrue(matcher._worth_a_thread(path, {path: 10}))

    def test_hashes_small_local_files_without_threads(self):
        paths = [add_file(self.root, f"{folder}/x.txt") for folder in "abc"]
        with mock.patch.object(matcher, "ThreadPoolExecutor", side_effect=AssertionError("threads used")):
            result = matcher.md5_map(paths, 4, sizes={p: os.path.getsize(p) for p in paths})
        self.assertEqual(len(result), 3)

    def test_lists_local_folders_one_at_a_time(self):
        add_file(self.root, "a/x.txt")
        with mock.patch.object(scanner, "_tree_listing", side_effect=AssertionError("threads used")):
            self.assertEqual(len(list(iter_files(self.root, throttle_limit=4))), 1)
        with as_if_on_a_network_drive(), mock.patch.object(scanner, "_tree_listing", wraps=scanner._tree_listing) as listing:
            list(iter_files(self.root, throttle_limit=4))
        listing.assert_called_once()

    def test_recognises_the_file_system_holding_a_path(self):
        mounts = [("/", "ext4"), ("/mnt/share", "cifs"), ("/mnt/shared", "nfs"), ("/mnt/my share", "nfs4")]
        self.assertEqual(scanner.mount_type("/mnt/share/photos", mounts), "cifs")
        self.assertEqual(scanner.mount_type("/mnt/shared2", mounts), "ext4", "a longer name is not inside the mount")
        self.assertEqual(scanner.mount_type("/mnt/my share/a", mounts), "nfs4")
        self.assertIn("cifs", scanner.NETWORK_FILE_SYSTEMS)
        self.assertNotIn("ext4", scanner.NETWORK_FILE_SYSTEMS)

    @unittest.skipUnless(sys.platform == "win32", "long-path forms are Windows-only")
    def test_treats_the_long_path_forms_of_a_path_as_what_they_point_to(self):
        self.assertTrue(scanner.on_network_drive("\\\\?\\UNC\\server\\share\\folder"))
        self.assertFalse(scanner.on_network_drive("\\\\?\\" + self.root))

    @unittest.skipUnless(sys.platform == "win32", "UNC paths are Windows-only")
    def test_treats_a_unc_path_as_a_network_drive(self):
        self.assertTrue(scanner.on_network_drive("\\\\server\\share\\folder"))


class ExcelDateTests(unittest.TestCase):
    def test_truncates_to_the_millisecond_like_dotnet(self):
        value = datetime(2024, 1, 2, 3, 4, 5, 678_999)
        self.assertEqual(from_excel_serial(excel_serial(value)), datetime(2024, 1, 2, 3, 4, 5, 678_000))


def write_excel_saved_workbook(path, rows, rules_rows=None):
    """A workbook shaped like one Excel has re-saved: shared strings, renamed worksheet
    parts, and cells without explicit types; with ``rules_rows``, a Rules sheet too."""
    strings = []

    def sheet(sheet_rows):
        xml_rows = []
        for r, row in enumerate(sheet_rows, start=1):
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
            xml_rows.append(f'<row r="{r}">{"".join(cells)}</row>')
        return f'<worksheet xmlns="{main}"><sheetData>{"".join(xml_rows)}</sheetData></worksheet>'

    main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rels = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    sheets = [("Duplicates", "data.xml", rows)] + ([("Rules", "notes.xml", rules_rows)] if rules_rows is not None else [])
    parts = {
        "[Content_Types].xml": '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
        "xl/workbook.xml": f'<workbook xmlns="{main}" xmlns:r="{rels}"><sheets>'
                           + "".join(f'<sheet name="{name}" sheetId="{i}" r:id="rId{i + 6}"/>' for i, (name, _, _) in enumerate(sheets, 1))
                           + "</sheets></workbook>",
        "xl/_rels/workbook.xml.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                                      + "".join(f'<Relationship Id="rId{i + 6}" Type="{rels}/worksheet" Target="worksheets/{part}"/>'
                                                for i, (_, part, _) in enumerate(sheets, 1))
                                      + "</Relationships>",
    }
    for _, part, sheet_rows in sheets:
        parts[f"xl/worksheets/{part}"] = sheet(sheet_rows)
    parts["xl/sharedStrings.xml"] = f'<sst xmlns="{main}">' + "".join(f"<si><t>{escape(t)}</t></si>" for t in strings) + "</sst>"
    with zipfile.ZipFile(path, "w") as archive:
        for name, content in parts.items():
            archive.writestr(name, content)


class ReadDuplicateReportTests(TempDirTestCase):
    ROUND_TRIP = [
        DuplicateSet("a & b.txt", datetime(2024, 1, 2, 3, 4, 5, 678_000), 1234, "AAAA", 3,
                     ["C:\\one", "C:\\two", "D:\\three"], utc_offset=timedelta(hours=10)),
        DuplicateSet("z.txt", datetime(2023, 6, 7, 8, 9, 10), 5, "BBBB", 2, ["C:\\x", "C:\\y"],
                     utc_offset=-timedelta(hours=4, minutes=30)),
    ]

    def test_validates_a_report_made_before_the_utc_offset_column_and_adds_the_column_when_it_rewrites_it(self):
        paths = [add_file(self.root, f"{folder}/x.txt") for folder in ("a", "b", "c")]
        path = os.path.join(self.root, "old.xlsx")
        info = os.stat(paths[0])
        serial = excel_serial(local_time(info.st_mtime_ns / 1e9))
        write_excel_saved_workbook(path, [
            ["File Name", "Last Modified", "Size (bytes)", "MD5", "Copies", "Location 1", "Location 2", "Location 3"],
            ["x.txt", serial, float(info.st_size), "ABC", 3.0] + [os.path.dirname(p) for p in paths],
        ])
        os.remove(paths[2])

        result = validate_report(path)

        self.assertEqual(result.copies_removed, 1, "the other two copies are found by local time")
        self.assertEqual(read_worksheet(path)[0][2], "UTC Offset")
        self.assertEqual(read_duplicate_report(path)[0].utc_offset, utc_offset(info.st_mtime_ns // 1_000_000_000))

    def test_writes_the_utc_offset_and_reads_it_back(self):
        for seconds, text in ((19800, "+05:30"), (-16200, "-04:30"), (0, "+00:00"), (50400, "+14:00"),
                              (1172, "+00:19:32")):  # the last: local mean time, as some zones used before 1900
            with self.subTest(text=text):
                self.assertEqual(utc_offset_text(timedelta(seconds=seconds)), text)
                self.assertEqual(parse_utc_offset(text, "report.xlsx"), timedelta(seconds=seconds))

    def test_rejects_a_utc_offset_it_cannot_read(self):
        with self.assertRaisesRegex(ValueError, "'10:00' is not a UTC offset"):
            parse_utc_offset("10:00", "report.xlsx")

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

    def test_reads_the_scan_settings_from_a_report_saved_by_excel(self):
        path = os.path.join(self.root, "settings.xlsx")
        write_excel_saved_workbook(
            path,
            [["File Name", "Last Modified", "Size (bytes)", "MD5", "Copies", "Location 1"]],
            [["Matching rules"], [EXCLUDE_NAMES_LABEL, None, "*.tmp", "Thumbs.db"], [MINIMUM_SIZE_LABEL, 2048]],
        )
        self.assertEqual(read_duplicate_workbook(path).settings, ScanSettings(["*.tmp", "Thumbs.db"], 2048))

    def test_ignores_an_exclusion_pattern_in_the_report_that_holds_a_slash_or_backslash(self):
        path = os.path.join(self.root, "patterns.xlsx")
        write_excel_saved_workbook(
            path, [["File Name", "Last Modified", "Size (bytes)", "MD5", "Copies", "Location 1"]],
            [[EXCLUDE_NAMES_LABEL, "photos/raw", "*.tmp"]],
        )
        with self.assertLogs("find_duplicates", "WARNING") as logs:
            self.assertEqual(read_duplicate_workbook(path).settings.exclude_names, ["*.tmp"])
        self.assertIn("'photos/raw'", logs.output[0])
        self.assertIn("not paths", logs.output[0])

    def test_ignores_a_smallest_file_size_in_the_report_that_is_not_a_whole_number_of_bytes(self):
        for bad in ("1MB", "1.5", "-1"):
            with self.subTest(bad=bad):
                path = os.path.join(self.root, "bad.xlsx")
                write_excel_saved_workbook(
                    path, [["File Name", "Last Modified", "Size (bytes)", "MD5", "Copies", "Location 1"]],
                    [[MINIMUM_SIZE_LABEL, bad]],
                )
                with self.assertLogs("find_duplicates", "WARNING") as logs:
                    self.assertEqual(read_duplicate_workbook(path).settings.minimum_size, 0)
                self.assertIn(f"'{bad}' is not a whole number of bytes", logs.output[0])
                os.remove(path)

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


class PreviousMd5Tests(TempDirTestCase):
    def scanned_tree(self):
        """Two duplicated files, each in two folders, and a report of them."""
        root = self.new_dir("tree")
        for path in ("a/x.txt", "b/x.txt"):
            add_file(root, path, "x")
        for path in ("c/y.txt", "d/y.txt"):
            add_file(root, path, "y")
        report = root + ".xlsx"
        export_duplicate_report(find_duplicate_files(list(iter_files(root))), report)
        return root, report

    def test_takes_the_md5_hashes_of_unchanged_files_from_the_previous_report(self):
        root, report = self.scanned_tree()
        with open(os.path.join(root, "c", "y.txt"), "w") as stream:
            stream.write("z")  # same size, new contents, saved now

        previous = validate.previous_md5(report, list(iter_files(root)))

        self.assertEqual(sorted(previous), [os.path.join(root, *p.split("/")) for p in ("a/x.txt", "b/x.txt", "d/y.txt")])
        self.assertEqual(previous[os.path.join(root, "a", "x.txt")], md5_file(os.path.join(root, "a", "x.txt")))

    def test_does_not_read_files_again_whose_md5_the_previous_report_holds(self):
        root, report = self.scanned_tree()
        files = list(iter_files(root))
        cache = validate.previous_md5(report, files)

        with mock.patch.object(matcher, "md5_file", side_effect=AssertionError("read again")), \
                mock.patch.object(matcher.log, "warning") as warning:
            result = find_duplicate_files(files, md5_cache=cache)

        warning.assert_not_called()
        self.assertEqual(len(result), 2)

    def test_takes_the_md5_hashes_from_a_report_whose_folders_are_in_another_case_on_windows(self):
        paths = [add_file(self.root, "My Photos/x.txt", "x"), add_file(self.root, "Backup/x.txt", "x")]
        files = self.records(*paths)
        report = os.path.join(self.root, "lower.xlsx")
        (found,) = find_duplicate_files(files)
        export_duplicate_report([dataclasses.replace(found, folders=[f.lower() for f in found.folders])], report)

        previous = validate.previous_md5(report, files)

        if sys.platform == "win32":
            self.assertEqual(sorted(previous), sorted(paths))
        else:
            self.assertEqual(previous, {}, "folder names are case-sensitive here")

    def test_ignores_an_md5_in_the_previous_report_that_is_not_an_md5(self):
        paths = [add_file(self.root, "a/x.txt"), add_file(self.root, "b/x.txt")]
        files = self.records(*paths)
        report = os.path.join(self.root, "edited.xlsx")
        first = files[0]
        export_duplicate_report([DuplicateSet(
            "x.txt", local_time(first.mtime_ns / 1e9), first.size, "not an md5", 2, [f.folder for f in files],
            utc_offset=utc_offset(first.mtime_ns // 1_000_000_000),
        )], report)

        self.assertEqual(validate.previous_md5(report, files), {})


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

    @unittest.skipIf(sys.platform == "win32", "the TZ variable sets the time zone only on Linux and macOS")
    def test_keeps_every_copy_when_validating_in_another_time_zone_than_the_scan(self):
        root = self.tree("a/x.txt", "b/x.txt")
        with time_zone("Asia/Kolkata"):
            report = self.scanned_report(root)
        with time_zone("America/New_York"):
            self.assertEqual(validate_report(report, dry_run=True).copies_removed, 0)

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
        with mock.patch.object(validate, "root_reachable", return_value=False), \
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
        with time_zone("America/New_York"):
            # 06:30 UTC on 3 Nov 2024 is 01:30 local, in the hour that happens twice.
            saved = datetime(2024, 11, 3, 6, 30, tzinfo=timezone.utc)
            path = add_file(self.root, "dst/x.txt", saved=saved)
            recorded = datetime.fromtimestamp(saved.timestamp())
            state = validate.check_copy(os.path.dirname(path), "x.txt", os.path.getsize(path), recorded)
            self.assertEqual(state, validate.PRESENT)

    def test_reports_a_copy_deleted_while_being_checked_as_missing(self):
        path = add_file(self.root, "gone.txt")
        os.remove(path)
        self.assertEqual(validate._copy_state(lambda: os.stat(path), 1, datetime.now(), None), validate.MISSING)

    def test_keeps_a_copy_whose_size_and_saved_date_cannot_be_read(self):
        # As for a file named NUL on Windows, which the system treats as a device.
        device = os.stat_result((stat.S_IFCHR | 0o666, 0, 0, 1, 0, 0, 0, 0, 0, 0))
        with mock.patch.object(validate.os.path, "isfile", return_value=True), \
                mock.patch.object(validate.os, "stat", return_value=device):
            state = validate.check_copy(self.root, "NUL", 0, datetime.now())
        self.assertEqual(state, validate.UNAVAILABLE)

    def test_checks_several_copies_in_one_folder_from_one_listing_by_exact_name_or_ignoring_case(self):
        present = add_file(self.root, "a/x.txt")
        add_file(self.root, "a/Photo.JPG")
        changed = add_file(self.root, "a/changed.txt")
        with open(changed, "w") as stream:
            stream.write("different now")
        info = os.stat(present)
        saved = local_time(info.st_mtime_ns / 1e9)
        checks = [validate._CopyCheck(name, name, info.st_size, saved, None)
                  for name in ("x.txt", "photo.jpg", "changed.txt", "gone.txt")]

        states = dict(validate._check_copies_in_folder(os.path.dirname(present), checks))

        self.assertEqual(states["x.txt"], validate.PRESENT)
        self.assertEqual(states["photo.jpg"], validate.PRESENT, "Photo.JPG is the same name, ignoring case")
        self.assertEqual(states["changed.txt"], validate.CHANGED)
        self.assertEqual(states["gone.txt"], validate.MISSING)

    def test_treats_every_copy_in_a_folder_that_has_gone_as_missing(self):
        checks = [validate._CopyCheck(name, name, 1, SAVED.replace(tzinfo=None), None) for name in ("x.txt", "y.txt")]
        states = validate._check_copies_in_folder(os.path.join(self.root, "gone"), checks)
        self.assertEqual([state for _, state in states], [validate.MISSING, validate.MISSING])

    def test_validates_the_same_way_checking_several_folders_at_a_time(self):
        root = self.new_dir("parallel")
        for folder in ("a", "b", "c"):
            for name in ("x.txt", "y.txt"):
                add_file(root, f"{folder}/{name}", name)
            add_file(root, f"{folder}/Photos/p.jpg", "photo")
        records = []
        files = list(iter_files(root, folders=records))
        report = root + ".xlsx"
        export_duplicate_report(find_duplicate_files(files), report, find_duplicate_folders(files, records))
        os.remove(os.path.join(root, "a", "x.txt"))
        shutil.rmtree(os.path.join(root, "b", "Photos"))

        one = validate_report(report, dry_run=True)
        with as_if_on_a_network_drive():
            many = validate_report(report, dry_run=True, throttle_limit=4)

        for field in ("copies_checked", "copies_removed", "rows_remaining",
                      "folder_copies_checked", "folder_copies_removed", "folder_rows_remaining"):
            self.assertEqual(getattr(many, field), getattr(one, field), field)
        self.assertEqual(one.copies_removed, 2, "a/x.txt and b/Photos/p.jpg are gone")
        self.assertEqual(one.folder_copies_removed, 1)

    def test_keeps_logging_after_checking_folder_copies_several_at_a_time(self):
        # Python-specific: each thread silences only its own messages while listing a folder copy.
        root = self.new_dir("logging")
        for folder in ("a", "b", "c", "d"):
            add_file(root, f"{folder}/Photos/p.jpg", "photo")
        records = []
        files = list(iter_files(root, folders=records))
        report = root + ".xlsx"
        export_duplicate_report(find_duplicate_files(files), report, find_duplicate_folders(files, records))

        with as_if_on_a_network_drive():
            validate_report(report, dry_run=True, throttle_limit=4)

        with self.assertLogs("find_duplicates", "WARNING"):
            logging.getLogger("find_duplicates").warning("still logging")

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

    def test_keeps_the_scan_settings_when_it_rewrites_the_report(self):
        root = self.tree("a/x.txt", "b/x.txt", "c/x.txt")
        report = root + ".xlsx"
        export_duplicate_report(find_duplicate_files(list(iter_files(root))), report, settings=ScanSettings(["Thumbs.db"], 5))
        os.remove(os.path.join(root, "c", "x.txt"))

        self.assertTrue(validate_report(report).saved)
        self.assertEqual(read_duplicate_workbook(report).settings, ScanSettings(["Thumbs.db"], 5))


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
        self.assertEqual(rows[1][5], "3")
        self.assertEqual(rows[1][6:9], [os.path.join(self.data, p) for p in ("2023", "backup", "old/copy".replace("/", os.sep))])
        self.assertIn("Found 1 duplicated files (3 copies in total).", out)

    def test_shows_the_full_help_instead_of_scanning_when_started_without_parameters(self):
        work = self.new_dir("work")
        add_file(work, "a/x.txt")
        add_file(work, "b/x.txt")
        out = io.StringIO()
        # As "python -m find_duplicates" calls it: main() reads sys.argv.
        with _chdir(work), mock.patch.object(sys, "argv", ["find-duplicates"]), contextlib.redirect_stdout(out):
            code = cli.main()

        self.assertEqual(code, 0)
        help_text = out.getvalue()
        self.assertIn("Find duplicate files in a folder", help_text)
        for option in ("path", "output", "--output-file", "--throttle-limit", "--folders", "--skip-folders",
                       "--ignore-empty-files",
                       "--exclude", "--minimum-size", "--skip-cloud-only", "--rehash", "--validate", "--dry-run",
                       "--verbose"):
            self.assertIn(option, help_text)
        self.assertIn("examples:", help_text)
        self.assertNotIn("Scanning '", help_text)
        self.assertFalse(os.path.exists(os.path.join(work, "duplicates.xlsx")))
        self.assertEqual(self.run_cli(), (0, help_text))

    def test_uses_the_given_output_name_and_adds_xlsx(self):
        work = self.new_dir("work")
        self.run_cli(self.data, os.path.join(work, "my-report"))
        self.assertTrue(os.path.exists(os.path.join(work, "my-report.xlsx")))
        self.run_cli(self.data, "-o", os.path.join(work, "named"))
        self.assertTrue(os.path.exists(os.path.join(work, "named.xlsx")))

    @unittest.skipUnless(sys.platform == "win32", "names ignore case only on Windows")
    def test_reports_the_folders_as_spelled_on_disk_when_given_the_path_in_another_case(self):
        root = self.new_dir("case")
        add_file(root, "My Photos/Trip A/x.txt")
        add_file(root, "My Photos/Trip B/x.txt")
        typed = os.path.join(root, "My Photos").lower()
        report = os.path.join(self.new_dir("work"), "case.xlsx")

        code, out = self.run_cli(typed, report)

        self.assertEqual(code, 0)
        (row,) = read_duplicate_report(report)
        self.assertEqual(row.folders, [os.path.join(root, "My Photos", "Trip A"), os.path.join(root, "My Photos", "Trip B")])
        self.assertIn(f"Scanning '{os.path.join(root, 'My Photos')}' ...", out)
        self.assertEqual(root[0], root[0].upper(), "drive letters are upper case")

    @unittest.skipUnless(sys.platform == "win32", "names ignore case only on Windows")
    def test_saves_the_report_under_its_folder_as_spelled_on_disk_keeping_the_new_file_name_as_given(self):
        reports = self.new_dir("Reports")
        typed = os.path.join(reports.lower(), "New Report")

        code, out = self.run_cli(self.data, typed)

        self.assertEqual(code, 0)
        self.assertIn(f"Report saved to '{os.path.join(reports, 'New Report.xlsx')}'.", out)

    def run_cli_with_errors(self, *args):
        out, err = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            code = cli.main(list(args))
        return code, out.getvalue(), err.getvalue()

    @unittest.skipUnless(sys.platform == "win32", "other programs lock files against writing only on Windows")
    def test_says_the_report_is_locked_before_scanning_when_another_program_holds_it_open(self):
        report = os.path.join(self.new_dir("work"), "locked.xlsx")
        self.run_cli(self.data, report)

        with locked_like_excel(report):
            code, out, err = self.run_cli_with_errors(self.data, report)

        self.assertEqual(code, 1)
        self.assertEqual(err.strip(), f"error: {locked_message(report)}")
        self.assertNotIn("Scanning", out, "nothing is scanned for a report that cannot be saved")

    @unittest.skipUnless(sys.platform == "win32", "other programs lock files against writing only on Windows")
    def test_says_the_report_is_locked_before_validating_when_another_program_holds_it_open(self):
        report = os.path.join(self.new_dir("work"), "locked.xlsx")
        self.run_cli(self.data, report)

        with locked_like_excel(report):
            code, out, err = self.run_cli_with_errors("--validate", report)

        self.assertEqual(code, 1)
        self.assertEqual(err.strip(), f"error: {locked_message(report)}")
        self.assertNotIn("Validating", out)

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

    def test_leaves_names_out_with_exclude_and_small_files_with_minimum_size(self):
        root = self.new_dir("excluded")
        for folder in ("a", "b"):
            add_file(root, f"{folder}/Thumbs.db")
            add_file(root, f"{folder}/small.txt", "small")
            add_file(root, f"{folder}/large.txt", "x" * 2048)
        report = os.path.join(self.new_dir("work"), "excluded.xlsx")

        self.run_cli(root, report, "--exclude", "thumbs.db", "--minimum-size", "2KB")

        self.assertEqual([d.file_name for d in read_duplicate_report(report)], ["large.txt"])
        rules = [row[0] for row in read_worksheet(report, "xl/worksheets/sheet3.xml", all_rows=True)]
        self.assertIn("Scan settings", rules)

    def test_reads_a_minimum_size_in_bytes_or_with_a_unit_and_rejects_anything_else(self):
        root = self.new_dir("sizes")
        for folder in ("a", "b"):
            add_file(root, f"{folder}/x.txt")
        for text, expected in (("1500", "1500"), ("2KB", "2048"), ("1.5MB", "1572864")):
            with self.subTest(text=text):
                report = os.path.join(self.root, "size.xlsx")
                self.run_cli(root, report, "--minimum-size", text)
                rows = read_worksheet(report, "xl/worksheets/sheet3.xml", all_rows=True)
                (size,) = [row for row in rows if row[0] == MINIMUM_SIZE_LABEL]
                self.assertEqual(size[1], expected)
        for text in ("-1", "10 bytes", "KB", "\u0661\u0662"):  # the last: Arabic-Indic digits
            with self.subTest(text=text), self.assertRaises(SystemExit):
                self.run_cli(root, os.path.join(self.root, "bad.xlsx"), "--minimum-size", text)

    def test_rejects_an_exclude_pattern_holding_a_slash_or_backslash_before_scanning(self):
        report = os.path.join(self.root, "never.xlsx")
        with self.assertRaises(SystemExit):
            self.run_cli(self.new_dir("any"), report, "--exclude", "photos/raw")
        self.assertFalse(os.path.exists(report))

    def test_records_ignore_empty_files_as_a_smallest_file_size_of_1_byte(self):
        root = self.new_dir("empty")
        for folder in ("a", "b"):
            add_file(root, f"{folder}/x.txt")
        report = os.path.join(self.root, "empty.xlsx")

        self.run_cli(root, report, "--ignore-empty-files")

        rows = read_worksheet(report, "xl/worksheets/sheet3.xml", all_rows=True)
        (size,) = [row for row in rows if row[0] == MINIMUM_SIZE_LABEL]
        self.assertEqual(size[1], "1")

    @unittest.skipUnless(sys.platform == "win32", "UNC paths are Windows-only")
    def test_scans_and_validates_a_network_share_given_as_a_unc_path(self):
        root = self.new_dir("unc")
        add_file(root, "a/x.txt")
        add_file(root, "b/x.txt")
        # Reach the local test folder through the administrative share, e.g. \\localhost\C$\...
        unc = "\\\\localhost\\" + root[0] + "$" + root[2:]
        if not os.path.isdir(unc):
            self.skipTest("the administrative share is not available")
        report = os.path.join(self.new_dir("work"), "unc.xlsx")

        code, out = self.run_cli(unc, report)
        self.assertEqual(code, 0)
        self.assertIn("On a network drive: working on 4 folders and files at a time", out)
        (row,) = read_duplicate_report(report)
        self.assertEqual(row.folders, [os.path.join(unc, "a"), os.path.join(unc, "b")])

        os.remove(os.path.join(root, "a", "x.txt"))
        self.assertEqual(self.run_cli("--validate", report)[0], 0)
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
            cli.main(["--validate", "--folders"])

    def test_reuses_the_previous_reports_md5_hashes_and_reads_every_file_again_with_rehash(self):
        report = os.path.join(self.new_dir("work"), "again.xlsx")
        self.run_cli(self.data, report)

        _, second = self.run_cli(self.data, report)
        _, third = self.run_cli(self.data, report, "--rehash")

        self.assertIn("Reusing 3 MD5 hashes from the previous report", second)
        self.assertNotIn("Reusing", third)
        self.assertEqual(len(read_duplicate_report(report)), 1)

    def test_scans_normally_when_the_previous_report_cannot_be_read(self):
        report = os.path.join(self.new_dir("work"), "damaged.xlsx")
        with open(report, "w") as stream:
            stream.write("not a workbook")

        with self.assertLogs("find_duplicates", "WARNING") as logs:
            code, _ = self.run_cli(self.data, report)

        self.assertEqual(code, 0)
        self.assertIn(f"Not reusing MD5 hashes from '{report}'", "\n".join(logs.output))
        self.assertEqual(len(read_duplicate_report(report)), 1)

    def test_checks_several_folders_at_a_time_with_validate_and_throttle_limit(self):
        root = self.new_dir("parallel")
        for folder in ("a", "b", "c"):
            add_file(root, f"{folder}/x.txt")
        report = os.path.join(self.new_dir("work"), "parallel.xlsx")
        self.run_cli(root, report)
        os.remove(os.path.join(root, "a", "x.txt"))

        with as_if_on_a_network_drive():
            code, _ = self.run_cli("--validate", report, "-j", "4")

        self.assertEqual(code, 0)
        self.assertEqual(read_duplicate_report(report)[0].count, 2)

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
