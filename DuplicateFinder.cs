// Compiled helpers for the loops of DuplicateFinder.psm1 that run once per file, where
// PowerShell's own overhead would cost many times the work itself. Loaded by the module with
// Add-Type. Behaviour is exactly that of the PowerShell code it replaces (and so of the Python
// port); only the speed differs.
//
// Kept to C# 5 and to types that Windows PowerShell 5.1 (.NET Framework) and PowerShell 7
// (.NET) both reference by default, so that it compiles on both. ASCII only.

using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace FindDuplicates
{
    // One folder's files and sub folders, each in ordinal name order, or why it could not be read.
    public sealed class FolderListing
    {
        public string Path;
        public FileInfo[] Files;
        public DirectoryInfo[] Folders;
        public string Error;
    }

    public static class Native
    {
        const int HashChunkBytes = 1024 * 1024;

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        struct Win32FindData
        {
            public uint Attributes;
            // Three FILETIMEs: pairs of 32-bit words, so not 8-byte aligned like a long.
            public uint CreationLow, CreationHigh, AccessLow, AccessHigh, WriteLow, WriteHigh;
            public uint SizeHigh, SizeLow, Reserved0, Reserved1;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string FileName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 14)] public string AlternateFileName;
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        static extern IntPtr FindFirstFileW(string fileName, out Win32FindData data);

        [DllImport("kernel32.dll")]
        static extern bool FindClose(IntPtr handle);

        // An absolute path as Windows stores it: drive letter in upper case, each existing
        // folder or file name spelled exactly as on disk (short 8.3 names expanded), no
        // trailing separator. The part from the first name that cannot be looked up on is kept
        // as given, as are a share's server and share names and \\?\ paths. Links are not
        // resolved. Elsewhere, where names are case-sensitive, the path is only made absolute.
        // (Python: scanner.full_path.)
        public static string ExactPath(string path)
        {
            path = Path.GetFullPath(path);
            if (Environment.OSVersion.Platform != PlatformID.Win32NT) { return TrimSeparator(path); }
            if (path.StartsWith(@"\\?\", StringComparison.Ordinal) || path.StartsWith(@"\\.\", StringComparison.Ordinal))
            {
                return path;
            }
            string root = Path.GetPathRoot(path);
            string[] names = path.Substring(root.Length).Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
            if (root.Length == 3 && root[1] == ':') { root = root.Substring(0, 1).ToUpperInvariant() + @":\"; }
            else if (!root.EndsWith(@"\", StringComparison.Ordinal)) { root += @"\"; }

            StringBuilder exact = new StringBuilder(root);
            bool lookUp = true;
            for (int i = 0; i < names.Length; i++)
            {
                string name = names[i];
                if (i > 0) { exact.Append('\\'); }
                if (lookUp)
                {
                    string found = null;
                    if (name.IndexOf('*') < 0 && name.IndexOf('?') < 0)
                    {
                        Win32FindData data;
                        IntPtr handle = FindFirstFileW(exact.ToString() + name, out data);
                        if (handle != new IntPtr(-1))
                        {
                            FindClose(handle);
                            found = data.FileName;
                        }
                    }
                    if (String.IsNullOrEmpty(found)) { lookUp = false; }
                    else { name = found; }
                }
                exact.Append(name);
            }
            return exact.ToString();
        }

        // Whether an existing file is locked by another program (for example a report open in
        // Excel), so that it cannot be opened for writing. A missing file is not locked; one
        // the user may not write (permissions, read-only) is not locked either.
        // (Python: xlsx.report_locked.)
        public static bool IsLocked(string path)
        {
            try
            {
                using (new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { }
                return false;
            }
            catch (FileNotFoundException) { return false; }
            catch (DirectoryNotFoundException) { return false; }
            catch (IOException) { return true; }
            catch (UnauthorizedAccessException) { return false; }
        }

        static string TrimSeparator(string path)
        {
            string root = Path.GetPathRoot(path);
            string trimmed = path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            return trimmed.Length < root.Length ? root : trimmed;
        }

        // A name in Unicode normal form C (so "e + accent" as macOS often stores it equals
        // the single character Windows stores); a name that is not valid UTF-16 as it is.
        public static string Normalize(string name)
        {
            try { return name.IsNormalized() ? name : name.Normalize(); }
            catch (ArgumentException) { return name; }
        }

        // The form in which names are compared: normalised and upper case (see ConvertTo-NameKey).
        public static string NameKey(string name)
        {
            return Normalize(name).ToUpperInvariant();
        }

        // The normalised name of each file or folder.
        public static string[] NormalizedNames(FileSystemInfo[] items)
        {
            string[] names = new string[items.Length];
            for (int i = 0; i < items.Length; i++) { names[i] = Normalize(items[i].Name); }
            return names;
        }

        // Groups items by the matching key (ignoring case), keeping only the groups that hold
        // more than one item, in the order their keys first appear (see Group-ByKey).
        public static List<object[]> GroupByKey(object[] items, string[] keys)
        {
            Dictionary<string, List<object>> groups = new Dictionary<string, List<object>>(StringComparer.OrdinalIgnoreCase);
            List<List<object>> order = new List<List<object>>();
            for (int i = 0; i < items.Length; i++)
            {
                List<object> group;
                if (!groups.TryGetValue(keys[i], out group))
                {
                    group = new List<object>();
                    groups.Add(keys[i], group);
                    order.Add(group);
                }
                group.Add(items[i]);
            }
            List<object[]> result = new List<object[]>();
            foreach (List<object> group in order)
            {
                if (group.Count > 1) { result.Add(group.ToArray()); }
            }
            return result;
        }

        // Files or folders sorted in place into ordinal name order; returns them.
        public static FileSystemInfo[] SortByName(FileSystemInfo[] items)
        {
            if (items.Length > 1)
            {
                string[] names = new string[items.Length];
                for (int i = 0; i < items.Length; i++) { names[i] = items[i].Name; }
                Array.Sort(names, items, StringComparer.Ordinal);
            }
            return items;
        }

        // One folder's files and sub folders from a single listing of it (over a network every
        // listing is a round trip), each in ordinal name order (see Get-FolderListing).
        public static FolderListing ListFolder(string path)
        {
            FolderListing listing = new FolderListing();
            listing.Path = path;
            FileSystemInfo[] entries;
            try { entries = new DirectoryInfo(path).GetFileSystemInfos(); }
            catch (Exception e)
            {
                if (!(e is UnauthorizedAccessException || e is IOException || e is System.Security.SecurityException)) { throw; }
                listing.Files = new FileInfo[0];
                listing.Folders = new DirectoryInfo[0];
                listing.Error = e.Message;
                return listing;
            }
            SortByName(entries);
            List<FileInfo> files = new List<FileInfo>(entries.Length);
            List<DirectoryInfo> folders = new List<DirectoryInfo>();
            foreach (FileSystemInfo entry in entries)
            {
                FileInfo file = entry as FileInfo;
                if (file != null) { files.Add(file); }
                else
                {
                    DirectoryInfo folder = entry as DirectoryInfo;
                    if (folder != null) { folders.Add(folder); }
                }
            }
            listing.Files = files.ToArray();
            listing.Folders = folders.ToArray();
            return listing;
        }

        // The files or folders whose name keys do not match the filter (see ConvertTo-NameFilter),
        // in order.
        public static FileSystemInfo[] WithoutMatchingNames(FileSystemInfo[] items, Regex filter)
        {
            List<FileSystemInfo> kept = new List<FileSystemInfo>(items.Length);
            foreach (FileSystemInfo item in items) { if (!filter.IsMatch(NameKey(item.Name))) { kept.Add(item); } }
            return kept.ToArray();
        }

        // The innermost exception, whose message and type say what actually failed.
        public static Exception Innermost(Exception e)
        {
            while (e.InnerException != null) { e = e.InnerException; }
            return e;
        }

        // Whether a file's saved date (as local and UTC) is a report's, to the whole second, in
        // exact integer arithmetic as when scanning: compared as instants when the report has the
        // UTC offset, as local times otherwise (reports made before the UTC Offset column).
        public static bool SameSavedDate(DateTime local, DateTime utc, DateTime lastWriteTime, TimeSpan? utcOffset)
        {
            long ticks = local.Ticks;
            long saved = lastWriteTime.Ticks;
            if (utcOffset.HasValue) { ticks = utc.Ticks; saved -= utcOffset.Value.Ticks; }
            return ticks - (ticks % TimeSpan.TicksPerSecond) == saved - (saved % TimeSpan.TicksPerSecond);
        }

        // Present or Changed for a copy's file, by its size and saved date; Missing when it was
        // deleted while being checked; Unavailable when its details cannot be read.
        public static string CopyState(FileInfo file, long sizeBytes, DateTime lastWriteTime, TimeSpan? utcOffset)
        {
            long size;
            DateTime local, utc;
            try
            {
                size = file.Length;
                local = file.LastWriteTime;
                utc = file.LastWriteTimeUtc;
            }
            catch (Exception e)
            {
                return Innermost(e) is FileNotFoundException ? "Missing" : "Unavailable";
            }
            return size == sizeBytes && SameSavedDate(local, utc, lastWriteTime, utcOffset) ? "Present" : "Changed";
        }

        // The state of each recorded copy in one folder (see Test-CopyInFolder), from a single
        // listing of it: by exact name, else by name key (ignoring case and Unicode form).
        public static string[] CheckCopiesInFolder(string folder, string[] fileNames, long[] sizes,
            DateTime[] lastWriteTimes, TimeSpan?[] utcOffsets)
        {
            string[] states = new string[fileNames.Length];
            FileInfo[] files;
            try { files = new DirectoryInfo(folder).GetFiles(); }
            catch (Exception e)
            {
                string failure = Innermost(e) is DirectoryNotFoundException ? "Missing" : "Unavailable";
                for (int i = 0; i < states.Length; i++) { states[i] = failure; }
                return states;
            }
            string[] names = new string[files.Length];
            for (int i = 0; i < files.Length; i++) { names[i] = files[i].Name; }
            string[] keys = null;  // worked out only when a name is not found as it is
            for (int i = 0; i < fileNames.Length; i++)
            {
                int at = Array.IndexOf(names, fileNames[i]);
                if (at < 0)
                {
                    if (keys == null)
                    {
                        keys = new string[names.Length];
                        for (int k = 0; k < names.Length; k++) { keys[k] = NameKey(names[k]); }
                    }
                    at = Array.IndexOf(keys, NameKey(fileNames[i]));
                }
                states[i] = at < 0 ? "Missing" : CopyState(files[at], sizes[i], lastWriteTimes[i], utcOffsets[i]);
            }
            return states;
        }

        // Text of an inline (<is>) or shared (<si>) string, including rich-text runs (<r>);
        // phonetic runs (<rPh>) are left out.
        public static string CellText(XmlNode node, string ns)
        {
            if (node == null) { return ""; }
            StringBuilder text = new StringBuilder();
            foreach (XmlNode child in node.ChildNodes)
            {
                if (child.NamespaceURI != ns) { continue; }
                if (child.LocalName == "t") { text.Append(child.InnerText); }
                else if (child.LocalName == "r")
                {
                    XmlElement run = child["t", ns];
                    if (run != null) { text.Append(run.InnerText); }
                }
            }
            return text.ToString();
        }

        // Column number of a cell reference's letters: A -> 1, Z -> 26, AA -> 27 ...
        static int ColumnIndex(string letters)
        {
            int index = 0;
            foreach (char letter in letters.ToUpperInvariant()) { index = index * 26 + (letter - 64); }
            return index;
        }

        // One worksheet as rows of cell text, one entry per column ($null for a missing cell).
        // Handles workbooks written by this tool and the same workbook after Excel saved it,
        // which may leave out empty cells: each cell is placed by its reference when present.
        public static List<string[]> SheetRows(XmlDocument sheet, IList<string> shared, string ns)
        {
            List<string[]> rows = new List<string[]>();
            XmlElement root = sheet.DocumentElement;
            XmlElement data = root == null ? null : root["sheetData", ns];
            if (data == null || root.LocalName != "worksheet" || root.NamespaceURI != ns) { return rows; }
            Dictionary<string, int> columnOf = new Dictionary<string, int>();
            foreach (XmlNode row in data.ChildNodes)
            {
                if (row.LocalName != "row" || row.NamespaceURI != ns) { continue; }
                SortedDictionary<int, string> cells = new SortedDictionary<int, string>();
                int column = 0;
                foreach (XmlNode node in row.ChildNodes)
                {
                    XmlElement cell = node as XmlElement;
                    if (cell == null || cell.LocalName != "c" || cell.NamespaceURI != ns) { continue; }
                    string reference = cell.GetAttribute("r");
                    if (reference.Length > 0)
                    {
                        string letters = reference.TrimEnd('0', '1', '2', '3', '4', '5', '6', '7', '8', '9');
                        if (!columnOf.TryGetValue(letters, out column))
                        {
                            column = ColumnIndex(letters);
                            columnOf[letters] = column;
                        }
                    }
                    else { column++; }
                    XmlElement value = cell["v", ns];
                    string type = cell.GetAttribute("t");
                    if (type == "s") { cells[column] = shared[int.Parse(value.InnerText, System.Globalization.CultureInfo.InvariantCulture)]; }
                    else if (type == "inlineStr") { cells[column] = CellText(cell["is", ns], ns); }
                    else if (value != null) { cells[column] = value.InnerText; }
                    else { cells[column] = ""; }
                }
                int width = 0;
                foreach (int c in cells.Keys) { width = Math.Max(width, c); }
                string[] values = new string[width];
                foreach (KeyValuePair<int, string> entry in cells) { values[entry.Key - 1] = entry.Value; }
                rows.Add(values);
            }
            return rows;
        }

        // MD5 of a file's contents as upper-case hex (the Get-FileHash format); with limit above
        // 0, of its first limit bytes only. Reads sequentially without buffering, in chunks of
        // up to 1 MB into a buffer no larger than the file.
        public static string Md5(string path, long limit)
        {
            using (MD5 md5 = MD5.Create())
            using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete, 1, FileOptions.SequentialScan))
            {
                long remaining = limit > 0 ? limit : long.MaxValue;
                byte[] buffer = new byte[(int)Math.Max(1L, Math.Min(Math.Min((long)HashChunkBytes, stream.Length), remaining))];
                int read;
                while (remaining > 0 && (read = stream.Read(buffer, 0, (int)Math.Min((long)buffer.Length, remaining))) > 0)
                {
                    md5.TransformBlock(buffer, 0, read, null, 0);
                    remaining -= read;
                }
                md5.TransformFinalBlock(buffer, 0, 0);
                StringBuilder hex = new StringBuilder(32);
                foreach (byte b in md5.Hash) { hex.Append(b.ToString("X2")); }
                return hex.ToString();
            }
        }
    }
}
