using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;

namespace ExcelReportGenerator
{
    public class DataProcessor
    {
        private readonly Collection<MonthModel> _monthModels;
        
        public DataProcessor(Collection<MonthModel> monthModels)
        {
            _monthModels = monthModels;
        }

        public void Process()
        {
            var monthSheets = new Collection<MonthSheetModel>();

            var t = from model in _monthModels
                orderby model.Day
                group model by model.MonthName
                into monthWithDays
                select new
                {
                    Month = monthWithDays.Key,
                    Days = monthWithDays
                };
                
            foreach (var variable in t)
            {
                var model = new MonthSheetModel
                {
                    Name = variable.Month,
                };

                foreach (var monthModel in variable.Days)
                {
                    var day = new DaySheetModel
                    {
                        DayNumber = monthModel.Name,
                    };

                    var rows = from record in monthModel.Records
                        group record by record.sys_acs_cd
                        into groupped
                        select new
                        {
                            Key = groupped.Key.ToString(),
                            Count = (int?)groupped.Count()
                        };

                    day.Row = rows.ToDictionary(it => it.Key, it => it.Count);

                    foreach (var i in day.Row)
                    {
                        model.Columns.Add(i.Key);
                    }

                   model.DaySheetModels.Add(day); 
                }

                monthSheets.Add(model);
            }
        }
    }

    public class MonthSheetModel
    {
        public string Name { get; set; }
        public Collection<DaySheetModel> DaySheetModels { get; set; }

        public HashSet<string> Columns { get; set; } 

        public MonthSheetModel()
        {
            DaySheetModels = new Collection<DaySheetModel>();
            Columns = new HashSet<string>();
        }
    }

    [DebuggerDisplay("Day = {DayNumber}")]
    public class DaySheetModel
    {
        public string DayNumber { get; set; }

        public Dictionary<string, int?> Row { get; set; }

        public DaySheetModel()
        {
            Row = new Dictionary<string, int?>();
        }
    }

//    public class RecordRaw
//    {
//        public string sys_acs_cd { get; set; }
//        public double? cst_nm { get; set; }
//        public string cst_nam { get; set; }
//        public double? bsns_srvc_nm { get; set; }
//        public string bsns_trn_long_nam { get; set; }
//        public string bsns_srvc_prcs_dt { get; set; }
//        public double? plcy_nm { get; set; }
//        public string plt_cd { get; set; }
//        public string iss_stck_ref_nm { get; set; }
//        public object bsns_srvc_sum_amt { get; set; }
//    }
//
//    public class RecordStrict
//    {
//        public string sys_acs_cd { get; set; }
//        public int? cst_nm { get; set; }
//        public string cst_nam { get; set; }
//        public double? bsns_srvc_nm { get; set; }
//        public string bsns_trn_long_nam { get; set; }
//        public string bsns_srvc_prcs_dt { get; set; }
//        public double? plcy_nm { get; set; }
//        public string plt_cd { get; set; }
//        public string iss_stck_ref_nm { get; set; }
//        public object bsns_srvc_sum_amt { get; set; }
//    }

    public class RecordRaw1
    {
        public object sys_acs_cd { get; set; }
        public object cst_nm { get; set; }
        public object cst_nam { get; set; }
        public object bsns_srvc_nm { get; set; }
        public object bsns_trn_long_nam { get; set; }
        public object bsns_srvc_prcs_dt { get; set; }
        public object plcy_nm { get; set; }
        public object plt_cd { get; set; }
        public object iss_stck_ref_nm { get; set; }
        public object bsns_srvc_sum_amt { get; set; }
    }
}