using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace ExcelReportGenerator
{
    public class DataProcessor
    {
        private readonly Collection<MonthModel> _monthModels;
        
        public DataProcessor(Collection<MonthModel> monthModels)
        {
            _monthModels = monthModels;
        }

        public void GenerateExcelResult(Collection<MonthSheetModel> monthSheets)
        {
          //  Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            
            string file = Path.GetTempFileName() + ".xlsx";
            
            var workbook = new XLWorkbook();
            
            foreach (var monthSheetModel in monthSheets)
            {
                var worksheet = workbook.Worksheets.Add(monthSheetModel.Name);
                
                var row1 = worksheet.Row(1);
                row1.Style.Font.Bold = true;
                row1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                var col1 = worksheet.Column("A");
                col1.Width = 20;
                col1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                
                for (int i = 0; i < monthSheetModel.Columns.ToArray().Length; i++)
                {
                    var column = monthSheetModel.Columns.ToArray()[i];

                //    row.Cell(column, excel.NewStyle().Align(HorizontalAlignment.Center).Bold());
                    worksheet.Cell(1, i + 2).Value = column;
                }

                int j;

                for (j = 0; j < monthSheetModel.DaySheetModels.Count; j++)
                {
                    //row = excel.Row();
                    if (j > 0)
                    {
                        worksheet.Column(j + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    }

                    var dict = monthSheetModel.DaySheetModels[j].Row;

                  //  row.Cell(monthSheetModel.DaySheetModels[j].DayNumber);

                    worksheet.Cell(j + 2, 1).Value = "'" + monthSheetModel.DaySheetModels[j].DayNumber;

                    for (int i = 0; i < monthSheetModel.Columns.ToArray().Length; i++)
                    {
                        var column = monthSheetModel.Columns.ToArray()[i];

                        if (dict.ContainsKey(column))
                        {
                            worksheet.Cell(j + 2, i + 2).Value = dict[column];
                        }

                    }
                }

                worksheet.Row(j + 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Row(j + 3).Style.Font.FontSize = 20;
                worksheet.Row(j + 3).Style.Font.Bold = true;

                for (int i = 0; i < monthSheetModel.ColumnsTotals.ToArray().Length; i++)
                {
                    var column = monthSheetModel.ColumnsTotals.ToArray()[i];

                    worksheet.Cell(j + 3, i + 2).Value = column;

                }
                
                worksheet.Cell(j + 3, monthSheetModel.ColumnsTotals.ToArray().Length + 3).Value = monthSheetModel.ColumnsTotals.Sum();
                worksheet.Column(monthSheetModel.ColumnsTotals.ToArray().Length + 3).Width = 20;
            }
            

            workbook.SaveAs(file);

            System.Diagnostics.Process.Start(file);
        }

        public void Process()
        {
            var monthSheets = new Collection<MonthSheetModel>();

            var t = from model in _monthModels
                orderby model.Day
                group model by model.MonthNameAndYear
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
                        DayNumber = monthModel.NameWithComma
                    };

                    var recordsWithDistintClient = from record in monthModel.Records
                        group record by record.cst_nm
                        into groupped
                        select groupped.First();

                    var rows = from record in FilteredRecords(recordsWithDistintClient.ToList())
                        group record by record.sys_acs_cd
                        into groupped
                        select new
                        {
                            Key = groupped.Key.ToString(),
                            Count = groupped.Count()
                        };

                    day.Row = rows.ToDictionary(it => it.Key, it => it.Count);

                    foreach (var i in day.Row)
                    {
                        model.Columns.Add(i.Key);
                    }
                    
                    model.DaySheetModels.Add(day); 
                }

                foreach (var column in model.Columns)
                {
                    model.ColumnsTotals.Add(model.DaySheetModels.Where(it => it.Row.ContainsKey(column)).Sum(it => it.Row[column]));
                }

                monthSheets.Add(model);
            }

            GenerateExcelResult(monthSheets);
        }

        private IEnumerable<RecordRaw1> FilteredRecords(List<RecordRaw1> input)
        {
            var filteredCollection = new Collection<RecordRaw1>();

            var records = input.ToArray();

            for (int i = 0; i < records.Length; i++)
            {
                string nextClientAction = String.Empty;
                string nextCustomerName = String.Empty;

                if (i + 1 < records.Length)
                {
                    nextClientAction = records[i + 1].bsns_trn_long_nam != null ? records[i + 1].bsns_trn_long_nam.ToString() : null;
                    nextCustomerName = records[i + 1].cst_nam != null ? records[i + 1].cst_nam.ToString() : null;
                }

                var clientAction = records[i].bsns_trn_long_nam != null ? records[i].bsns_trn_long_nam.ToString() : null;
                var customerName = records[i].cst_nam != null ? records[i].cst_nam.ToString() : null;

                if(clientAction == null) continue;
                
                if (clientAction.Contains("Cust.Cntct") || clientAction.Contains("Notified"))
                {
                    continue;
                }
                
                if (clientAction.Contains("Customer Comments"))
                {
                    if (nextCustomerName == null || nextClientAction == null) continue;
                    
                    if (nextCustomerName.Equals(customerName) && nextClientAction.Contains("Cust.Cntct"))
                    {
                        continue;
                    }
                }

                filteredCollection.Add(records[i]);
            }

            return filteredCollection;
        }
    }
}