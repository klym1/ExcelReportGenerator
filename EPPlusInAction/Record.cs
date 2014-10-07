using System;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExcelReportGenerator
{
    public class DataProcessor
    {
        private readonly Collection<MonthModel> _monthModels;


        public DataProcessor(Collection<MonthModel> monthModels)
        {
            this._monthModels = monthModels;
        }

        public void Process()
        {
            var t = from model in _monthModels
                group model by model.MonthName
                into monthWithDays
                select monthWithDays;
        }

    }

    public class MonthSheetModel
    {
        public string Name { get; set; }

    }

    public class RecordRaw
    {
        public string sys_acs_cd { get; set; }
        public double? cst_nm { get; set; }
        public string cst_nam { get; set; }
        public double? bsns_srvc_nm { get; set; }
        public string bsns_trn_long_nam { get; set; }
        public string bsns_srvc_prcs_dt { get; set; }
        public double? plcy_nm { get; set; }
        public string plt_cd { get; set; }
        public string iss_stck_ref_nm { get; set; }
        public object bsns_srvc_sum_amt { get; set; }
    }

    public class RecordStrict
    {
        public string sys_acs_cd { get; set; }
        public int? cst_nm { get; set; }
        public string cst_nam { get; set; }
        public double? bsns_srvc_nm { get; set; }
        public string bsns_trn_long_nam { get; set; }
        public string bsns_srvc_prcs_dt { get; set; }
        public double? plcy_nm { get; set; }
        public string plt_cd { get; set; }
        public string iss_stck_ref_nm { get; set; }
        public object bsns_srvc_sum_amt { get; set; }
    }

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