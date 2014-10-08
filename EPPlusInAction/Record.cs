using System;
using System.Collections.Generic;
using System.Diagnostics;
using ExcelFile.net;
using NPOI.HSSF.UserModel;

namespace ExcelReportGenerator
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