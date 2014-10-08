using System;
using System.Collections.Generic;
using System.IO;

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
               
        public MonthModel GetMonthModel()
        {
            var t = new MyExcelReader();

            try
            {
                var g = t.Page_Load(_filePath);

                var list = new List<RecordRaw1>();

                //excluding headers, starting from first row

                var rowsNumber = g.GetLength(0);
                
                for (int i = 1; i < rowsNumber; i++)
                {
                    var row = new RecordRaw1
                    {
                        sys_acs_cd = g[i, 0],
                        cst_nm = g[i, 1],
                        cst_nam = g[i, 2],
                        bsns_srvc_nm = g[i, 3],
                        bsns_trn_long_nam = g[i, 4],
                        bsns_srvc_prcs_dt = g[i, 5],
                        plcy_nm = g[i, 6],
                        plt_cd = g[i, 7],
                        iss_stck_ref_nm = g[i, 8],
                        bsns_srvc_sum_amt = g[i, 9]
                    };

                    list.Add(row);
                }

                return new MonthModel(list)
                {
                    Name = Path.GetFileNameWithoutExtension(_filePath)
                };
            }
            catch (Exception)
            {
                throw new Exception("Input file was in incorrect format!");
            }
            
        }
    }
}