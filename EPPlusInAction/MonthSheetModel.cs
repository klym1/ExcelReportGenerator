using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace ExcelReportGenerator
{
    public class MonthSheetModel
    {
        public string Name { get; set; }
        public Collection<DaySheetModel> DaySheetModels { get; set; }

        public HashSet<string> Columns { get; set; }

        public Collection<long> ColumnsTotals { get; set; } 

        public MonthSheetModel()
        {
            DaySheetModels = new Collection<DaySheetModel>();
            Columns = new HashSet<string>();
            ColumnsTotals = new Collection<long>();
        }
    }
}