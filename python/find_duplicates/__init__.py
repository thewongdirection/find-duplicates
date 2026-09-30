"""Find duplicate files (same name, saved date and MD5) and save them to Excel.

Python equivalent of the PowerShell tool in the repository root; see
python/README.md for how the two map onto each other.
"""

from .folders import DuplicateFolderSet, find_duplicate_folders
from .matcher import DuplicateSet, find_duplicate_files, md5_file, md5_map
from .scanner import FileRecord, FolderRecord, is_cloud_only, is_folder_link, iter_files
from .validate import ValidationResult, check_copy, check_folder_copy, previous_md5, validate_report
from .xlsx import (
    DuplicateWorkbook,
    column_name,
    export_duplicate_report,
    read_duplicate_folder_report,
    read_duplicate_report,
    read_duplicate_workbook,
)

__all__ = [
    "DuplicateFolderSet",
    "DuplicateSet",
    "DuplicateWorkbook",
    "FileRecord",
    "FolderRecord",
    "ValidationResult",
    "check_copy",
    "check_folder_copy",
    "column_name",
    "export_duplicate_report",
    "find_duplicate_files",
    "find_duplicate_folders",
    "is_cloud_only",
    "is_folder_link",
    "iter_files",
    "md5_file",
    "md5_map",
    "previous_md5",
    "read_duplicate_folder_report",
    "read_duplicate_report",
    "read_duplicate_workbook",
    "validate_report",
]
