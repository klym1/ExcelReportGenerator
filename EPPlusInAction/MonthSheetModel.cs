using System.Collections.Generic;
using System.Collections.ObjectModel;
using DocumentFormat.OpenXml.Office.CoverPageProps;
using ExcelReportGenerator.Models;

namespace ExcelReportGenerator
{
    public class MonthSheetModel
    {
        private string _name;

        public string Name
        {
            get { return _name; }
            set
            {
                _name = value;
                Order = MonthesAndIds.GetMonthIdByName(value);
            }
        }

        public int Order { get; set; }

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