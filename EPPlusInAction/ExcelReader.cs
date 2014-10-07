﻿using System;
using System.Data;
using System.IO;
using System.Linq;
using Excel;

namespace ExcelReportGenerator
{
    public class ExcelReader
    {
        private readonly string _filePath;
        
        public ExcelReader(string filePath)
        {
            _filePath = filePath;

            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException(filePath);
            }
        }

        private int? GetNullableIntFromDouble(double? d)
        {
            if (d.HasValue)
            {
                if (Math.Round((decimal) d.Value) == Convert.ToDecimal(d.Value))
                {
                    return Convert.ToInt32(d.Value);
                }

                throw new Exception("True doulbe");
            }

            return null;
        }

        public MonthModel GetMonthModel()
        {
            FileStream stream = File.Open(_filePath, FileMode.Open,
                              FileAccess.Read, FileShare.ReadWrite);
            
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream))
            {
                excelReader.IsFirstRowAsColumnNames = true;
                DataSet result = excelReader.AsDataSet();

                var table = result.Tables[0].AsEnumerable();

                var g =
                    table.AsEnumerable()
                        .Select(dataRow => new RecordRaw
                        {
                            sys_acs_cd = dataRow.Field<string>("sys_acs_cd"),
                            cst_nm = dataRow.Field<double?>("cst_nm"),
                            cst_nam = dataRow.Field<string>("cst_nam"),
                            bsns_srvc_nm = dataRow.Field<double?>("bsns_srvc_nm"),
                            bsns_trn_long_nam = dataRow.Field<string>("bsns_trn_long_nam"),
                            bsns_srvc_prcs_dt = dataRow.Field<string>("bsns_srvc_prcs_dt"),
                            plcy_nm = dataRow.Field<double?>("plcy_nm"),
                            plt_cd = dataRow.Field<string>("plt_cd"),
                            iss_stck_ref_nm = dataRow.Field<string>("iss_stck_ref_nm"),
                            bsns_srvc_sum_amt = dataRow.Field<double?>("bsns_srvc_sum_amt"),
                        })
                        .Select(it => new RecordStrict
                        {
                            sys_acs_cd = it.sys_acs_cd,
                            cst_nm = GetNullableIntFromDouble(it.cst_nm),

                            cst_nam = it.cst_nam,
                            bsns_srvc_nm = it.bsns_srvc_nm,
                            bsns_trn_long_nam = it.bsns_trn_long_nam,
                            bsns_srvc_prcs_dt = it.bsns_srvc_prcs_dt,
                            plcy_nm = it.plcy_nm,
                            plt_cd = it.plt_cd,
                            iss_stck_ref_nm = it.iss_stck_ref_nm,
                            bsns_srvc_sum_amt = it.bsns_srvc_sum_amt,
                        });
                        
                

                return new MonthModel(g);
            }
        }
    }
}