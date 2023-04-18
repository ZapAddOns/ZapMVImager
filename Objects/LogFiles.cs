using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;

namespace ZapMVImager.Objects
{
    public static class LogFiles
    {

        #region Static Functions

        public static List<LogFile> CreateListOfFiles(string name)
        {
            var result = new List<LogFile>();

            if (Directory.Exists(name))
            {
                // Name is a directory
                result.AddRange(GetAllTreatmentViewFilesInDir(name));
            }
            else if (IsZipFile(name))
            {
                result.AddRange(GetAllTreatmentViewFilesInZip(name));
            }
            else if (File.Exists(name))
            {
                result.AddRange(new List<LogFile> { new LogFile(name, "TXT", "") });
            }

            return result;
        }

        public static ObservableCollection<LogFile> SortFiles(List<LogFile> files)
        {
            if (files.Count <= 1)
            {
                return new ObservableCollection<LogFile>(files);
            }

            files.Sort(new LogFileComparer(_regexFilename));

            return new ObservableCollection<LogFile>(files);
        }

        #endregion

        #region Private Static Functions

        private static Regex _regexFilename = new Regex(@"TreatmentView_(\d{4}-\d{2}-\d{2})\.log\.?(\d*)", RegexOptions.Compiled);

        private static bool IsZipFile(string filename)
        {
            if (!File.Exists(filename))
            {
                return false;
            }

            var buffer = new byte[4];
            var file = File.OpenRead(filename);

            file.Read(buffer, 0, 4);
            file.Close();

            return buffer[0] == 0x50 && buffer[1] == 0x4b && buffer[2] == 0x03 && buffer[3] == 0x04;
        }

        private static List<LogFile> GetAllTreatmentViewFilesInDir(string path)
        {
            var result = new List<LogFile>();

            foreach (var filename in Directory.EnumerateFiles(path))
            {
                if (IsZipFile(filename))
                {
                    result.AddRange(GetAllTreatmentViewFilesInZip(filename));
                }
                if (IsCorrectFilename(filename))
                {
                    result.Add(new LogFile(Path.Combine(path, filename), "TXT", ""));
                }
            }

            foreach (var directory in Directory.EnumerateDirectories(path))
            {
                result.AddRange(GetAllTreatmentViewFilesInDir(directory));
            }

            return result;
        }

        private static List<LogFile> GetAllTreatmentViewFilesInZip(string filename)
        {
            var result = new List<LogFile>();

            using (ZipArchive archive = ZipFile.OpenRead(filename))
            {
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    if (IsCorrectFilename(entry.Name))
                    {
                        result.Add(new LogFile(entry.FullName, "ZIP", filename));
                    }
                }
            }

            return result;
        }

        private static bool IsCorrectFilename(string filename)
        {
            return _regexFilename.IsMatch(filename);
        }

        #endregion
    }
}
