using System;

namespace ExcelReportGenerator
{
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

        public bool IsEmpty
        {
            get
            {
                return sys_acs_cd == null
                       && cst_nm == null
                       && cst_nam == null
                       && bsns_srvc_nm == null
                       && bsns_trn_long_nam == null
                       && bsns_srvc_prcs_dt == null
                       && plcy_nm == null
                       && plt_cd == null
                       && iss_stck_ref_nm == null
                       && bsns_srvc_sum_amt == null;
            }
        }
    }
}