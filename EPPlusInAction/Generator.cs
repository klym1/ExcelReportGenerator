using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace ExcelReportGenerator
{
    public class Generator
    {
        private readonly DataProcessor _dataProcessor;

        public Generator(DataProcessor dataProcessor)
        {
            _dataProcessor = dataProcessor;
        }

        private Collection<MonthSheetModel> _privateMonthModels;
        private Collection<SummarySheetRowModel> _summarySheetRowModels;

        private Collection<MonthSheetModel> MonthSheetModels
        {
            get { return _privateMonthModels ?? (_privateMonthModels = _dataProcessor.GetMonthSheetModels()); }
        }

        private Collection<SummarySheetRowModel> SummarySheetRowModels
        {
            get { return _summarySheetRowModels ?? (_summarySheetRowModels = _dataProcessor.GetSummarySheetRowModels()); }
        }

        public void GenerateExcelResult()
        {
            string file = Path.GetTempFileName() + ".xlsx";

            var workbook = new XLWorkbook();

            GenerateMonthTabs(MonthSheetModels, workbook);
            GenerateSummaryTab(SummarySheetRowModels, workbook);

            workbook.SaveAs(file);

            Process.Start(file);
        }

        private void GenerateSummaryTab(Collection<SummarySheetRowModel> monthSheets, XLWorkbook workbook)
        {
            var worksheet = workbook.Worksheets.Add("Summary");

            worksheet.Column(1).Width = 25;
            worksheet.Column(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

            worksheet.Column(2).Width = 20;
            worksheet.Column(2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            worksheet.Column(3).Width = 15;
            worksheet.Column(3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Column(3).Style.Font.Bold = true;

            int startingRow = 1;
            string previousAgentCode = null;

            int i;

            for (i = 0; i < monthSheets.Count; i++)
            {
                var agentCode = monthSheets[i].AgentCode;

                if (previousAgentCode != null)
                {
                    if (!previousAgentCode.Equals(agentCode))
                    {
                        worksheet
                            .Range(startingRow, 1, i, 1)
                            .Merge();
                        
                        startingRow = i + 1;
                    }
                   
                }

                worksheet.Cell(i + 1, 1).Value = "'" + agentCode;
                worksheet.Cell(i + 1, 2).Value = "'" + monthSheets[i].Month;
                worksheet.Cell(i + 1, 3).Value = monthSheets[i].NumberOfClients;

                previousAgentCode = agentCode;
            }

            worksheet.Range(startingRow, 1, i, 1).Merge();
        }

        private static void GenerateMonthTabs(Collection<MonthSheetModel> monthSheets, XLWorkbook workbook)
        {
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

                    worksheet.Cell(1, i + 2).Value = column;
                }

                int j;

                for (j = 0; j < monthSheetModel.DaySheetModels.Count; j++)
                {
                    if (j > 0)
                    {
                        worksheet.Column(j + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    }

                    var dict = monthSheetModel.DaySheetModels[j].Row;

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
        }
    }
}