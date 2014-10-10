using System.Collections.Generic;
using System.Diagnostics;

namespace ExcelReportGenerator.Models
{
    [DebuggerDisplay("Day = {DayNumber}")]
    public class DaySheetModel
    {
        public string DayNumber { get; set; }

        public Dictionary<string, int> Row { get; set; }

        public DaySheetModel()
        {
            Row = new Dictionary<string, int>();
        }
    }
}