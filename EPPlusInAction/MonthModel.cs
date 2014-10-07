using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExcelReportGenerator
{
    public class MonthModel
    {
        public List<RecordStrict> Records { get; private set; }

        public MonthModel()
        {
            Records = new List<RecordStrict>();
        }

        public MonthModel(IEnumerable<RecordStrict> records)
        {
            Records = records.ToList();
        }
    }
}
