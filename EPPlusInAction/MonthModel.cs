using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelReportGenerator
{
    [DebuggerDisplay("Name = {Name}")]
    public class MonthModel
    {
        private string _name;
        public List<RecordRaw1> Records { get; private set; }

        public string NameWithComma
        {
            get
            {
                if (paramsList != null)
                {
                    return String.Format("{0} {1}, {2}", paramsList[0], paramsList[1], paramsList[2]);
                }

                return String.Empty;
            }
        }

        public string Name
        {
            get { return _name; }
            set
            {
                _name = value;
                paramsList = reverseStringFormat(NameTemplate, value);
            }
        }

        public string MonthNameAndYear
        {
            get
            {
                if (paramsList != null)
                {
                    return String.Format("{0} {1}", paramsList[0] ?? String.Empty,paramsList[2] ?? String.Empty);
                }

                return String.Empty;
            }
        }

        public string MonthName
        {
            get
            {
                if (paramsList != null)
                {
                    return paramsList[0];
                }

                return String.Empty;
            }
        }

        public int Day
        {
            get
            {
                if (paramsList != null)
                {
                    int result;
                    
                    if (Int32.TryParse(paramsList[1], out result))
                    {
                        return result;
                    }
                }

                return 0;
            }
        }

        public int? Year
        {
            get
            {
                if (paramsList != null)
                {
                    int result;

                    if (Int32.TryParse(paramsList[2], out result))
                    {
                        return result;
                    }
                }

                return 0;
            }
        }

        private const string NameTemplate = "{0} {1} {2}";

        private List<string> paramsList;

        // from StackOverflow
        private List<string> reverseStringFormat(string template, string str)
        {
            string pattern = "^" + Regex.Replace(template, @"\{[0-9]+\}", "(.*?)") + "$";

            var r = new Regex(pattern);
            Match m = r.Match(str);

            var ret = new List<string>();

            for (int i = 1; i < m.Groups.Count; i++)
            {
                ret.Add(m.Groups[i].Value);
            }

            return ret;
        }

        public MonthModel()
        {
            Records = new List<RecordRaw1>();
        }

        public MonthModel(IEnumerable<RecordRaw1> records)
        {
            Records = records.ToList();
        }
    }
}
