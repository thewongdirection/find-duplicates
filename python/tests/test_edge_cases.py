"""Edge cases: very long paths, huge files and trees, unusual dates and names, file
system quirks and Excel's limits. Mirrors the 'Edge cases' Describe block in
tests/DuplicateFinder.Tests.ps1, plus checks specific to Python."""

from __future__ import annotations

import contextlib
import hashlib
import io
import os
import shutil
import stat
import subprocess
import sys
import tempfile
import time
import unittest
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Iterator, List, Optional
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from find_duplicates import cli, scanner, xlsx  # noqa: E402
from find_duplicates.folders import find_duplicate_folders  # noqa: E402
from find_duplicates.matcher import DuplicateSet, find_duplicate_files  # noqa: E402
from find_duplicates.scanner import FolderRecord, iter_files, local_time  # noqa: E402
from find_duplicates.validate import validate_report  # noqa: E402
from find_duplicates.xlsx import (  # noqa: E402
    export_duplicate_report, read_duplicate_folder_report, read_duplicate_report,
)
from tests.helpers import SAVED, add_file  # noqa: E402

ON_WINDOWS = sys.platform == "win32"
LIBREOFFICE_REQUIRED = bool(os.environ.get("FIND_DUPLICATES_REQUIRE_LIBREOFFICE"))


def raw_path(path: str) -> str:
    """On Windows, the \\\\?\\ form of a full path: only that form can create names Windows
    otherwise reserves or trims, and paths longer than 260 characters."""
    return "\\\\?\\" + path if ON_WINDOWS else path


def add_raw_file(folder: str, name: str, content: str = "same content") -> str:
    """Create a file through its raw path; returns its plain full path."""
    os.makedirs(raw_path(folder), exist_ok=True)
    path = os.path.join(folder, name)
    with open(raw_path(path), "w", encoding="utf-8", newline="") as stream:
        stream.write(content)
    seconds = SAVED.timestamp()
    os.utime(raw_path(path), (seconds, seconds))
    return path


def running_as_root() -> bool:
    return not ON_WINDOWS and os.geteuid() == 0


def long_paths_enabled() -> bool:
    """False on Windows when the system setting that lifts the 260-character limit is off."""
    if not ON_WINDOWS:
        return True
    import winreg

    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\FileSystem") as key:
            return winreg.QueryValueEx(key, "LongPathsEnabled")[0] == 1
    except OSError:
        return False


@contextlib.contextmanager
def time_zone(name: str) -> Iterator[None]:
    """Run with the TZ variable set (Linux and macOS only)."""
    previous = os.environ.get("TZ")
    os.environ["TZ"] = name
    time.tzset()
    try:
        yield
    finally:
        if previous is None:
            del os.environ["TZ"]
        else:
            os.environ["TZ"] = previous
        time.tzset()


@contextlib.contextmanager
def unreadable(path: str) -> Iterator[None]:
    """Make a file unreadable to this process: on Windows, held open by a handle that shares
    nothing; elsewhere, no read permission."""
    if ON_WINDOWS:
        import ctypes
        from ctypes import wintypes

        create_file = ctypes.windll.kernel32.CreateFileW
        create_file.restype = wintypes.HANDLE
        generic_read, open_existing, share_none = 0x80000000, 3, 0
        handle = create_file(path, generic_read, share_none, None, open_existing, 0, None)
        try:
            yield
        finally:
            ctypes.windll.kernel32.CloseHandle(handle)
    else:
        os.chmod(path, 0)
        try:
            yield
        finally:
            os.chmod(path, stat.S_IRUSR | stat.S_IWUSR)


def save_with_libreoffice(path: str, out_folder: str) -> Optional[str]:
    """Open a report in LibreOffice Calc and save it again as .xlsx; returns the new file's
    path, or None when LibreOffice is not installed or cannot convert."""
    soffice = shutil.which("soffice")
    if soffice is None:
        return None
    profile = Path(out_folder, "profile").resolve().as_uri()
    subprocess.run(
        [soffice, f"-env:UserInstallation={profile}", "--headless", "--norestore",
         "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", out_folder, path],
        capture_output=True, timeout=300,
    )
    saved = os.path.join(out_folder, os.path.basename(path))
    return saved if os.path.exists(saved) else None


class EdgeCaseTests(unittest.TestCase):
    def setUp(self) -> None:
        self._temp = tempfile.TemporaryDirectory()
        self.root = scanner.full_path(self._temp.name)

    def tearDown(self) -> None:
        self._temp.cleanup()

    def new_root(self) -> str:
        return tempfile.mkdtemp(dir=self.root)

    def scan(self, root: str, throttle_limit: int = 1):
        """Files, duplicate files and duplicate folders below ``root``, sharing one MD5 cache."""
        folders: List[FolderRecord] = []
        files = list(iter_files(root, folders=folders))
        cache = {}
        found = find_duplicate_files(files, throttle_limit=throttle_limit, md5_cache=cache)
        return files, found, find_duplicate_folders(files, folders, md5_cache=cache)

    def test_finds_duplicates_in_folders_whose_full_path_is_longer_than_260_characters(self):
        root = self.new_root()
        deep = os.sep.join(["d" * 50] * 5)
        for copy in ("a", "b"):
            add_raw_file(os.path.join(root, copy, deep), "x.txt")

        try:
            with self.assertLogs("find_duplicates", "DEBUG") as logs:
                result = find_duplicate_files(list(iter_files(root)))
            if not long_paths_enabled():
                # Windows without the long path setting: it must say what it could not reach.
                self.assertTrue(any("WARNING" in line for line in logs.output))
                return
            self.assertEqual(len(result), 1)
            self.assertEqual(result[0].count, 2)
            for folder in result[0].folders:
                self.assertGreater(len(folder), 260)

            report = os.path.join(root, "long.xlsx")
            export_duplicate_report(result, report)
            self.assertEqual(validate_report(report).copies_removed, 0)
        finally:
            # Removed here: the temporary folder cleanup may not reach long paths on Windows.
            for copy in ("a", "b"):
                shutil.rmtree(raw_path(os.path.join(root, copy)))

    @unittest.skipIf(ON_WINDOWS, "sparse test files need extra set-up on Windows")
    def test_handles_files_larger_than_4_gb(self):
        root = self.new_root()
        size = 4 * 1024 ** 3 + 1
        for copy in ("one", "two"):
            folder = os.path.join(root, copy, "big")
            os.makedirs(folder)
            path = os.path.join(folder, "big.bin")
            # Extending a file without writing to it makes a sparse file: no disk space is used.
            with open(path, "wb") as stream:
                stream.truncate(size)
            os.utime(path, (SAVED.timestamp(), SAVED.timestamp()))

        _, result, folder_sets = self.scan(root, throttle_limit=2)
        report = os.path.join(self.new_root(), "big.xlsx")
        export_duplicate_report(result, report, folder_sets)

        self.assertEqual(len(result), 1)
        self.assertEqual(result[0].md5, "F18C798FF5D450DFE4D3ACDC12B621FF")
        self.assertEqual(read_duplicate_report(report)[0].size_bytes, size)
        self.assertEqual(read_duplicate_folder_report(report)[0].size_bytes, size)
        summary = validate_report(report)
        self.assertEqual(summary.copies_removed + summary.folder_copies_removed, 0)

    def test_finds_every_duplicate_in_a_tree_of_10000_files(self):
        root = self.new_root()
        # 50 folders in each of two trees, 100 files in each; every file has one copy in the
        # other tree, and many files share a name, date and size but not their contents.
        for copy in ("a", "b"):
            for f in range(50):
                folder = os.path.join(root, copy, f"folder{f}")
                os.makedirs(folder)
                for i in range(100):
                    path = os.path.join(folder, f"file{i}.txt")
                    with open(path, "w") as stream:
                        stream.write(f"{f}-{i}")
                    os.utime(path, (SAVED.timestamp(), SAVED.timestamp()))

        files, result, folder_sets = self.scan(root, throttle_limit=4)
        report = os.path.join(self.new_root(), "many.xlsx")
        export_duplicate_report(result, report, folder_sets)

        self.assertEqual(len(files), 10000)
        self.assertEqual(len(result), 5000)
        self.assertEqual(len(folder_sets), 50)
        self.assertEqual(len(read_duplicate_report(report)), 5000)
        self.assertEqual(validate_report(report).copies_removed, 0)

    def test_does_not_match_a_copy_whose_saved_time_a_fat_drive_rounded_to_2_seconds(self):
        # FAT and exFAT (USB sticks, memory cards) store saved times to 2 seconds, so a copy of
        # a file saved at 10:30:01 reads 10:30:02. The saved date must match to the second.
        root = self.new_root()
        add_file(root, "disk/x.txt", saved=SAVED + timedelta(seconds=1))
        add_file(root, "usb/x.txt", saved=SAVED + timedelta(seconds=2))
        self.assertEqual(find_duplicate_files(list(iter_files(root))), [])

    def test_finds_saves_and_validates_duplicates_saved_before_1970_and_after_2038(self):
        for year in (1960, 2040):  # a negative Unix time; beyond 32-bit Unix time
            with self.subTest(year=year):
                saved = datetime(year, 1, 15, 8, 0, 0, tzinfo=timezone.utc)
                root = self.new_root()
                add_file(root, "a/dated.txt", saved=saved)
                add_file(root, "b/dated.txt", saved=saved)
                report = os.path.join(root, "dates.xlsx")
                export_duplicate_report(find_duplicate_files(list(iter_files(root))), report)

                self.assertEqual(read_duplicate_report(report)[0].last_write_time, local_time(saved.timestamp()))
                self.assertEqual(validate_report(report).copies_removed, 0)

    def test_converts_saved_dates_before_1970_where_the_platform_cannot(self):
        # Python-specific: datetime.fromtimestamp rejects negative times on Windows.
        real = datetime

        class NoNegativeTimes(datetime):
            @classmethod
            def fromtimestamp(cls, t, tz=None):
                if t < 0:
                    raise OSError(22, "Invalid argument")
                return real.fromtimestamp(t, tz)

        for zone in ("UTC", "America/New_York") if not ON_WINDOWS else ("",):
            for month in (1, 7):
                with self.subTest(zone=zone, month=month):
                    seconds = real(1960, month, 15, 12, tzinfo=timezone.utc).timestamp()
                    with time_zone(zone) if zone else contextlib.nullcontext():
                        expected = real.fromtimestamp(seconds) if not ON_WINDOWS else None
                        with mock.patch.object(scanner, "datetime", NoNegativeTimes):
                            converted = local_time(seconds)
                    if expected is not None:
                        self.assertEqual(converted, expected)
                    self.assertEqual(converted.year, 1960)

    @unittest.skipIf(ON_WINDOWS, "the TZ variable sets the time zone only on Linux and macOS")
    def test_keeps_copies_saved_in_winter_and_in_summer_when_validating_in_a_time_zone_with_daylight_saving(self):
        with time_zone("America/New_York"):
            root = self.new_root()
            for name, month in (("summer", 7), ("winter", 1)):
                saved = datetime(2024, month, 15, 12, 0, 0, tzinfo=timezone.utc)
                add_file(root, f"a/{name}.txt", saved=saved)
                add_file(root, f"b/{name}.txt", saved=saved)
            report = os.path.join(root, "seasons.xlsx")
            export_duplicate_report(find_duplicate_files(list(iter_files(root))), report)

            # Noon UTC is 8:00 in summer (UTC-4) and 7:00 in winter (UTC-5).
            self.assertEqual([s.last_write_time.hour for s in read_duplicate_report(report)], [8, 7])
            self.assertEqual(validate_report(report).copies_removed, 0)

    def test_scans_a_folder_given_as_a_symbolic_link(self):
        root = self.new_root()
        add_file(root, "real/a/x.txt")
        add_file(root, "real/b/x.txt")
        link = os.path.join(root, "link")
        try:
            os.symlink(os.path.join(root, "real"), link, target_is_directory=True)
        except (OSError, NotImplementedError) as exc:
            self.skipTest(f"symbolic links cannot be created here: {exc}")

        result = find_duplicate_files(list(iter_files(link)))
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0].folders, [os.path.join(link, "a"), os.path.join(link, "b")])

    def test_skips_a_folder_deleted_during_the_scan_and_carries_on(self):
        root = self.new_root()
        for folder in ("first", "second", "third"):
            add_file(root, f"{folder}/x.txt")

        def on_folder(path: str, folders: int, files: int) -> None:
            # Delete "second" while "first" is being scanned, after the root's listing included it.
            if os.path.basename(path) == "first":
                shutil.rmtree(os.path.join(os.path.dirname(path), "second"))

        with self.assertLogs("find_duplicates", "WARNING") as logs:
            files = list(iter_files(root, on_folder=on_folder))

        self.assertEqual([f.folder for f in files], [os.path.join(root, "first"), os.path.join(root, "third")])
        self.assertIn(os.path.join(root, "second"), "\n".join(logs.output))
        self.assertEqual(len(find_duplicate_files(files)), 1)

    @unittest.skipIf(running_as_root(), "root can read every file")
    def test_skips_a_file_the_operating_system_will_not_let_it_read(self):
        root = self.new_root()
        locked = add_file(root, "locked/x.txt")
        add_file(root, "b/x.txt")
        add_file(root, "c/x.txt")

        with unreadable(locked), self.assertLogs("find_duplicates", "WARNING") as logs:
            result = find_duplicate_files(list(iter_files(root)))

        self.assertIn("locked", "\n".join(logs.output))
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0].count, 2)

    def test_reports_two_files_in_one_folder_whose_names_differ_only_in_case(self):
        root = self.new_root()
        folder = os.path.join(root, "photos")
        os.makedirs(folder)
        if ON_WINDOWS:
            # Windows folders can be made case-sensitive (as WSL does); that needs no admin rights.
            subprocess.run(["fsutil.exe", "file", "setCaseSensitiveInfo", folder, "enable"], capture_output=True)
        add_file(root, "photos/IMG.JPG")
        if os.path.exists(os.path.join(folder, "img.jpg")):
            self.skipTest("this folder ignores case")
        add_file(root, "photos/img.jpg")

        result = find_duplicate_files(list(iter_files(root)))
        report = os.path.join(self.new_root(), "case.xlsx")
        export_duplicate_report(result, report)

        self.assertEqual(len(result), 1)
        self.assertEqual(result[0].file_name, "IMG.JPG")
        self.assertEqual(result[0].folders, [folder, folder])
        self.assertEqual(validate_report(report).copies_removed, 0)

    def wide_set(self, copies: int, first_folder: str = "/copy1") -> DuplicateSet:
        folders = [first_folder] + [f"/copy{n}" for n in range(2, copies + 1)]
        return DuplicateSet("x.txt", SAVED.replace(tzinfo=None), 1, "A", copies, folders)

    def test_writes_a_file_with_as_many_copies_as_excel_has_location_columns(self):
        report = os.path.join(self.new_root(), "wide.xlsx")
        export_duplicate_report([self.wide_set(16379)], report)
        self.assertEqual(read_duplicate_report(report)[0].count, 16379)

    def test_refuses_a_file_with_more_copies_than_excel_has_location_columns(self):
        report = os.path.join(self.new_root(), "too-wide.xlsx")
        with self.assertRaisesRegex(ValueError, "^A file has 16380 copies; Excel supports at most 16379 location columns.$"):
            export_duplicate_report([self.wide_set(16380)], report)
        self.assertFalse(os.path.exists(report))

    def test_refuses_more_duplicated_files_than_excel_has_rows(self):
        report = os.path.join(self.new_root(), "too-long.xlsx")
        sets = [DuplicateSet(f"{n}.txt", SAVED.replace(tzinfo=None), 1, "A", 2, ["/a", "/b"]) for n in (1, 2, 3)]
        with mock.patch.object(xlsx, "EXCEL_MAX_ROWS", 3):  # instead of writing a million rows
            with self.assertRaisesRegex(ValueError, "^Found 3 duplicated files; Excel supports at most 2 rows.$"):
                export_duplicate_report(sets, report)

    def test_refuses_text_longer_than_an_excel_cell_can_hold(self):
        report = os.path.join(self.new_root(), "long-cell.xlsx")
        with self.assertRaisesRegex(ValueError, "^Cell F2 would hold 32768 characters; Excel allows at most 32767.$"):
            export_duplicate_report([self.wide_set(2, "/" + "x" * 32767)], report)
        self.assertFalse(os.path.exists(report))

    def test_counts_cell_text_the_way_excel_does(self):
        # Python-specific: Excel counts UTF-16 code units, so an emoji counts twice.
        emoji = chr(0x1F600)
        fits = self.wide_set(2, "/" + emoji * 16383)          # 1 + 2 * 16383 = 32767
        too_long = self.wide_set(2, "/" + emoji * 16383 + "x")
        export_duplicate_report([fits], os.path.join(self.new_root(), "fits.xlsx"))
        with self.assertRaisesRegex(ValueError, "32768 characters"):
            export_duplicate_report([too_long], os.path.join(self.new_root(), "too-long.xlsx"))

    def test_reads_and_validates_a_report_after_libreoffice_has_saved_it(self):
        root = self.new_root()
        data = os.path.join(root, "data")
        for copy in ("a", "b"):
            add_file(data, f"{copy}/x & y.txt")
            add_file(data, f"{copy}/Holiday/p.jpg", "photo")
        report = os.path.join(root, "report.xlsx")
        with contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(cli.main([data, report, "--folders"]), 0)
        saved = save_with_libreoffice(report, os.path.join(root, "resaved"))
        if saved is None:
            if LIBREOFFICE_REQUIRED:
                self.fail("LibreOffice Calc is required for this test but could not convert the report.")
            self.skipTest("LibreOffice Calc is not installed")

        self.assertEqual(read_duplicate_report(saved), read_duplicate_report(report))
        self.assertEqual([f.folder_name for f in read_duplicate_folder_report(saved)], ["Holiday"])
        summary = validate_report(saved)
        self.assertEqual(summary.copies_removed + summary.folder_copies_removed, 0)

    def test_copes_with_names_windows_reserves_or_trims(self):
        names = (
            "NUL",               # reserved on Windows (a device)
            "PRN.txt",           # reserved on Windows, even with an extension
            "ends in a dot.",    # Windows trims a trailing dot from ordinary paths
            "ends in a space ",  # ... and a trailing space
        )
        for name in names:
            with self.subTest(name=name):
                root = self.new_root()
                for copy in ("a", "b"):
                    add_raw_file(os.path.join(root, copy), name)
                try:
                    # Windows: it must finish without failing (such copies may be reported or skipped
                    # with a warning). Elsewhere these are ordinary names and must be matched.
                    with self.assertLogs("find_duplicates", "DEBUG"):
                        result = find_duplicate_files(list(iter_files(root)))
                        report = os.path.join(self.new_root(), "names.xlsx")
                        export_duplicate_report(result, report)
                        summary = validate_report(report)
                    if not ON_WINDOWS:
                        self.assertEqual(len(result), 1)
                        self.assertEqual(result[0].file_name, name)
                        self.assertEqual(summary.copies_removed, 0)
                finally:
                    # Removed here: ordinary deletion cannot reach these names on Windows.
                    for copy in ("a", "b"):
                        shutil.rmtree(raw_path(os.path.join(root, copy)))

    def test_includes_hidden_and_system_files_and_folders(self):
        root = self.new_root()
        paths = [add_file(root, ".hidden-folder/.hidden"), add_file(root, "b/.hidden")]
        if ON_WINDOWS:
            import ctypes

            hidden, system, archive = 0x2, 0x4, 0x20
            for path in paths:
                ctypes.windll.kernel32.SetFileAttributesW(path, hidden | system | archive)
            ctypes.windll.kernel32.SetFileAttributesW(os.path.join(root, ".hidden-folder"), hidden)

        result = find_duplicate_files(list(iter_files(root)))
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0].file_name, ".hidden")

    def test_orders_rows_by_file_name_ignoring_case_then_saved_date_then_md5(self):
        root = self.new_root()
        later = SAVED + timedelta(hours=1)
        add_file(root, "one/B.txt", "b")
        add_file(root, "two/b.txt", "b")
        add_file(root, "one/a.txt", "x", later)
        add_file(root, "two/a.txt", "x", later)
        add_file(root, "three/a.txt", "x")
        add_file(root, "four/a.txt", "x")
        add_file(root, "five/A.txt", "y")
        add_file(root, "six/A.txt", "y")
        add_file(root, "one/a", "z")
        add_file(root, "two/a", "z")
        same_date = sorted(hashlib.md5(text).hexdigest().upper() for text in (b"x", b"y"))

        result = find_duplicate_files(list(iter_files(root)))

        # A name sorts before the longer names it starts.
        self.assertEqual([s.file_name.upper() for s in result], ["A", "A.TXT", "A.TXT", "A.TXT", "B.TXT"])
        self.assertEqual([s.md5 for s in result[1:3]], same_date)
        self.assertEqual(result[3].last_write_time, local_time(later.timestamp()))


if __name__ == "__main__":
    unittest.main()
