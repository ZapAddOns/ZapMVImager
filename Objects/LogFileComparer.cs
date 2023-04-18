using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ZapMVImager.Objects
{
    internal class LogFileComparer : IComparer<LogFile>
    {
        Regex _regex;

        public LogFileComparer(Regex regex)
        {
            _regex = regex;
        }

        public int Compare(LogFile x, LogFile y)
        {
            var matchX = _regex.Match(x.Filename);
            var matchY = _regex.Match(y.Filename);

            if (matchX.Groups[1].Value == matchY.Groups[1].Value)
            {
                if (string.IsNullOrEmpty(matchX.Groups[2].Value))
                {
                    return 1;
                }
                if (string.IsNullOrEmpty(matchY.Groups[2].Value))
                {
                    return -1;
                }

                var valueX = int.Parse(matchX.Groups[2].Value);
                var valueY = int.Parse(matchY.Groups[2].Value);

                if (valueX < valueY)
                {
                    return 1;
                }
                else
                {
                    return -1;
                }
            }

            return string.Compare(matchX.Groups[1].Value, matchY.Groups[1].Value);
        }
    }
}
