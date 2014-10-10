using System.Collections.Generic;
using System.Linq;

namespace ExcelReportGenerator
{
    internal static class MonthesAndIds
    {
        private static readonly Dictionary<string, int> Dictionary = new Dictionary<string, int>
        {
            {"Jan", 1},
            {"Feb", 2},
            {"Mar", 3},
            {"Apr", 4},
            {"May", 5},
            {"Jun", 6},
            {"Jul", 7},
            {"Aug", 8},
            {"Sep", 9},
            {"Oct", 10},
            {"Nov", 11},
            {"Dec", 12}
        };
        
        public static int GetMonthIdByName(string monthName)
        {
            var find = Dictionary.FirstOrDefault(it => monthName.Contains(it.Key));

            return find.Value;
        }
    }
}