using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExcelReportGenerator
{
    public class MonthModel
    {
        public List<RecordRaw1> Records { get; private set; }

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
