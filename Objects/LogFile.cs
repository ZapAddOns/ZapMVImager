using System.IO;
using System.IO.Compression;

namespace ZapMVImager.Objects
{
    public class LogFile
    {
        readonly string _file;
        readonly string _filetype;
        readonly string _zipPath;
        StreamReader _streamReader;
        ZipArchive _zipArchive;

        public LogFile(string filename, string filetype, string zipPath)
        {
            _file = filename;
            _filetype = filetype;
            _zipPath = zipPath;
        }

        public string Filename { get => Path.GetFileName(_file); }

        public bool Open()
        {
            try
            {
                switch (_filetype)
                {
                    case "TXT":
                        _streamReader = File.OpenText(_file);
                        break;
                    case "ZIP":
                        _zipArchive = ZipFile.OpenRead(_zipPath);
                        _streamReader = new StreamReader(_zipArchive.GetEntry(_file).Open());
                        break;
                }
            }
            catch
            {
                return false;
            }

            return true;
        }

        public bool Close()
        {
            try
            {
                _streamReader?.Close();
                _zipArchive?.Dispose();
            }
            catch
            {
                return false;
            }

            return true;
        }

        public string GetNextLine()
        {
            return _streamReader.ReadLine();
        }

        public override bool Equals(object obj)
        {
            var item = obj as LogFile;

            if (item == null)
            {
                return false;
            }

            return _file.Equals(item._file);
        }

        public override int GetHashCode()
        {
            return _file.GetHashCode();
        }
    }
}
