using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExcelReportGenerator
{
    public class MonthModel
    {
        public List<RecordRaw1> Records { get; private set; }
        public string Name { get; set; }

        public string MonthName
        {
            get
            {
                if (!string.IsNullOrEmpty(Name))
                {
                    return Name.Substring(0, Name.IndexOf(' '));
                }

                return String.Empty;
            }
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
