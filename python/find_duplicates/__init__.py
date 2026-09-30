"""Find duplicate files (same name, saved date and MD5) and save them to Excel.

Python equivalent of the PowerShell tool in the repository root; see
python/README.md for how the two map onto each other.
"""

from .matcher import DuplicateSet, find_duplicate_files, md5_file
from .scanner import FileRecord, is_cloud_only, is_folder_link, iter_files
from .xlsx import column_name, export_duplicate_report

__all__ = [
    "DuplicateSet",
    "FileRecord",
    "column_name",
    "export_duplicate_report",
    "find_duplicate_files",
    "is_cloud_only",
    "is_folder_link",
    "iter_files",
    "md5_file",
]
