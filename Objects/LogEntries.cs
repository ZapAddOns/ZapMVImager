using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ZapMVImager.Objects
{
    public class LogEntries
    {
        int _fileIndex;
        List<LogFile> _files;
        LogFile _activeFile;
        Dictionary<string, List<LogEntry>> _logEntries;
        Action<int, int, string> _progressFiles;
        Action<int, int> _progressEntries;

        public LogEntries(List<LogFile> files, Action<int, int, string> progressFiles, Action<int, int> progressEntries)
        {
            _files = files;
            _fileIndex = -1;
            _progressFiles = progressFiles;
            _progressEntries = progressEntries;

            _logEntries = new Dictionary<string, List<LogEntry>>();
        }

        public List<string> PlanNames { get => _logEntries.Keys.ToList(); }

        public void CreateLogEntries()
        {
            // Clear all log entries in case this function is called again
            _logEntries = new Dictionary<string, List<LogEntry>>();

            var content = GetNextContent();
            var totalEntries = 0;

            while (content != null)
            {
                if (content.IsTreatment)
                {
                    if (!_logEntries.ContainsKey(content.PlanName))
                    {
                        _logEntries[content.PlanName] = new List<LogEntry>();
                    }

                    _logEntries[content.PlanName].Add(content);

                    _progressEntries(_logEntries.Keys.Count, ++totalEntries);
                }

                content = GetNextContent();
            }

            UpdateIsocenterValues();
        }

        public List<string> GetDatesForPlan(string planName)
        {
            var result = new List<string>();

            if (string.IsNullOrEmpty(planName))
            {
                return result;
            }

            if (!_logEntries.ContainsKey(planName))
            {
                return result;
            }

            foreach (var entry in _logEntries[planName])
            {
                var date = entry.Time.Date.ToShortDateString();

                if (!result.Contains(date))
                {
                    result.Add(date);
                }
            }

            return result;
        }

        public List<LogEntry> GetEntriesForPlanAndDate(string planName, DateTime date)
        {
            var startDate = date;
            var endDate = date.AddDays(1);
            var result = new List<LogEntry>();

            foreach (var entry in _logEntries[planName])
            {
                if (entry.Time > startDate && entry.Time < endDate)
                {
                    result.Add(entry);
                }
            }

            return result;
        }

        private void UpdateIsocenterValues()
        {
            foreach (var plan in _logEntries.Keys)
            {
                var isocenter = 1;
                var lastDate = DateTime.MinValue;
                var lastBeamNum = -1;

                foreach (var entry in _logEntries[plan])
                {
                    if (entry.Node < lastBeamNum)
                    {
                        isocenter++;
                    }

                    if (entry.Time.Date != lastDate)
                    {
                        isocenter = 1;
                        lastDate = entry.Time.Date;
                    }

                    entry.Isocenter = isocenter;

                    lastBeamNum = entry.Node;
                }
            }
        }

        private LogEntry GetNextContent()
        {
            LogEntry result;

            var line = GetNextLine();

            while (line != null && !LogEntry.RegexMeasured.IsMatch(line))
            {
                line = GetNextLine();
            }

            if (line == null)
            {
                return null;
            }

            // We now have the first line of a content

            result = new LogEntry();

            while (line != null)
            {
                if (LogEntry.RegexMeasured.IsMatch(line))
                {
                    var match = LogEntry.RegexMeasured.Match(line);

                    result.Time = DateTime.ParseExact(match.Groups[1].Value + " " + match.Groups[2].Value, "MM.dd.yyyy HH:mm:ss.FFF", CultureInfo.InvariantCulture);
                    result.Intensity = double.Parse(match.Groups[3].Value, CultureInfo.InvariantCulture);
                    result.FieldSizeInMm = double.Parse(match.Groups[5].Value, CultureInfo.InvariantCulture);
                    result.Node = int.Parse(match.Groups[6].Value);
                    result.IsValid = match.Groups[7].Value.ToUpper() == "TRUE" ? true : false;
                }

                if (LogEntry.RegexDoseChecker.IsMatch(line))
                {
                    var match = LogEntry.RegexDoseChecker.Match(line);

                    result.DeliveredMU = double.Parse(match.Groups[4].Value, CultureInfo.InvariantCulture);
                    result.ImagerMU = double.Parse(match.Groups[5].Value, CultureInfo.InvariantCulture);
                    result.IsFlagged = match.Groups[7].Value.ToUpper() == "YES" ? true : false;
                }

                if (LogEntry.RegexCumulative.IsMatch(line))
                {
                    var match = LogEntry.RegexCumulative.Match(line);

                    result.CumulativeDeliveredMU = double.Parse(match.Groups[3].Value, CultureInfo.InvariantCulture);
                    result.CumulativeImagerMU = double.Parse(match.Groups[4].Value, CultureInfo.InvariantCulture);
                    //result.CumulativeDifferenceMU = double.Parse(match.Groups[5].Value, CultureInfo.InvariantCulture);
                    //result.CumulativeDifferencePercent = double.Parse(match.Groups[6].Value, CultureInfo.InvariantCulture) * 100.0;
                }

                if (LogEntry.RegexSystemData.IsMatch(line))
                {
                    var match = LogEntry.RegexSystemData.Match(line);

                    result.TreatmentType = match.Groups[3].Value;
                    result.IsTreatment = match.Groups[4].Value.ToUpper() == "TRUE" ? true : false;
                    result.PlannedMU = double.Parse(match.Groups[5].Value, CultureInfo.InvariantCulture);
                    result.ColliSize = double.Parse(match.Groups[6].Value, CultureInfo.InvariantCulture);
                    result.Axial = double.Parse(match.Groups[7].Value, CultureInfo.InvariantCulture);
                    result.Oblique = double.Parse(match.Groups[8].Value, CultureInfo.InvariantCulture);
                    result.PlanName = match.Groups[9].Value;

                    // Last entry in the log file for this content

                    return result;
                }

                line = GetNextLine();
            }

            return null;
        }

        private string GetNextLine()
        {
            // Is this the first call?
            if (_fileIndex == -1)
            {
                _fileIndex = 0;
                _activeFile = _files[_fileIndex];
                _activeFile.Open();

                _progressFiles(_fileIndex, _files.Count, _activeFile.Filename);
            }

            var line = _activeFile.GetNextLine();

            // Are there no more lines to check in this file left?
            while (line == null)
            {
                _fileIndex++;

                if (_fileIndex >= _files.Count)
                {
                    // No more files left
                    _activeFile.Close();

                    _progressFiles(_fileIndex, _files.Count, string.Empty);

                    return null;
                }

                _activeFile.Close();

                _activeFile = _files[_fileIndex];

                _progressFiles(_fileIndex, _files.Count, _activeFile.Filename);

                _activeFile.Open();

                line = _activeFile.GetNextLine();
            }

            return line;
        }
    }
}
