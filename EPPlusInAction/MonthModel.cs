using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExcelReportGenerator
{
    public class MonthModel
    {
        public List<Record> Records { get; private set; }

        public MonthModel()
        {
            Records = new List<Record>();
        }

        public MonthModel(IEnumerable<Record> records)
        {
            Records = records.ToList();
        }
    }
}
