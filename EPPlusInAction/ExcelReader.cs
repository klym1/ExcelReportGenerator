using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Falcon.Excel;
using log4net.ObjectRenderer;

namespace ExcelReportGenerator
{
    public static class Single
    {
        private static ExcelController _excelController;

        public static ExcelController Instanse
        {
            get
            {
                if (_excelController == null)
                {
                    _excelController = new ExcelController();
                }

                return _excelController;
            }
        }
    }

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

        private const int MaxRows = 8096;

        private void DebugOpcode(UInt16 opcode)
        {
            switch (opcode)
            {
                case 0x00:
                    Debug.WriteLine("Dimensions");
                    break;
            }
        }

        private void ReadBody(BinaryReader reader, int bytes)
        {
            switch (bytes)
            {
                case 0: break;
                case 1:
                    reader.ReadByte(); 
                    break;
                case 2:
                    reader.ReadUInt16();
                    break;
                case 4:
                    reader.ReadUInt32();
                    break;
                case 8:
                    reader.ReadUInt64();
                    break;
                case 12:
                    reader.ReadUInt64();
                    reader.ReadUInt32();
                    break;
            }
        }

        public MonthModel GetMonthModel()
        {
            var sheetName =  Path.GetFileNameWithoutExtension(_filePath);


            var url1 = @"C:\Users\mykola.klymyuk\Desktop\3rd Quater 2014\Aug " + 19 + @" 2014.xls";
            var url2 = @"C:\Users\mykola.klymyuk\Desktop\3rd Quater 2014\Aug " + 7 + @" 2014.xls";

            var arr1 = new Collection<Int64>();

            using (Stream stream = new FileStream(url1, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (var memoryStream = new BinaryReader(stream))
                {
                    for (int i = 0; i < 50; i++)
                    {
                        arr1.Add(memoryStream.ReadInt64());
                    }
                    
                }
            }

            var arr2 = new Collection<Int64>();

            using (Stream stream = new FileStream(url2, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (var memoryStream = new BinaryReader(stream))
                {
                    for (int i = 0; i < 50; i++)
                    {
                        arr2.Add(memoryStream.ReadInt64());
                    }

                }
            }

           
           
           


            var t = new MyExcelReader();

         var g =   t.Page_Load(_filePath);

            //var t = new Net.SourceForge.Koogra.Excel.Workbook(_filePath);

           // var sheets = t.Sheets;
            
          //  Single.Instanse.OpenFile(_filePath);
            
            var currentRowIndex = 1; // exclude headers

            var list = new List<RecordRaw1>();

            while (currentRowIndex < MaxRows)
            {
                var row = new RecordRaw1
               {
                   sys_acs_cd = ((object)Single.Instanse.GetData(sheetName, currentRowIndex, 0)),
//                   cst_nm = ((object)g.GetData(sheetName, currentRowIndex, 1)),
//                   cst_nam = ((object)g.GetData(sheetName, currentRowIndex, 2)),
//                   bsns_srvc_nm = ((object)g.GetData(sheetName, currentRowIndex, 3)),
//                   bsns_trn_long_nam = ((object)g.GetData(sheetName, currentRowIndex, 4)),
//                   bsns_srvc_prcs_dt = ((object)g.GetData(sheetName, currentRowIndex, 5)),
//                   plcy_nm = ((object)g.GetData(sheetName, currentRowIndex, 6)),
//                   plt_cd = ((object)g.GetData(sheetName, currentRowIndex, 7)),
//                   iss_stck_ref_nm = ((object)g.GetData(sheetName, currentRowIndex, 8)),
//                   bsns_srvc_sum_amt = ((object)g.GetData(sheetName, currentRowIndex, 9))
                };

                if(row.IsEmpty) break;
                
                list.Add(row);
                currentRowIndex++;
            }
            
            return new MonthModel(list);
        }
    }
}