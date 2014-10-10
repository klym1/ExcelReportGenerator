using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using ExcelReportGenerator.Models;

namespace ExcelReportGenerator
{
    public class DataProcessor
    {
        private readonly Collection<MonthModel> _monthModels;
        
        public DataProcessor(Collection<MonthModel> monthModels)
        {
            _monthModels = monthModels;
        }

        private Collection<MonthSheetModel> _privateMonthSheetModels;
        private Collection<SummarySheetRowModel> _privateSummarySheetRowModels;

        public Collection<SummarySheetRowModel> GetSummarySheetRowModels()
        {
            if (_privateSummarySheetRowModels == null)
            {
                _privateSummarySheetRowModels = CalculateSummarySheetRowModels(GetMonthSheetModels());
            }

            return _privateSummarySheetRowModels;
        }

        public Collection<MonthSheetModel> GetMonthSheetModels()
        {
            return _privateMonthSheetModels ?? (_privateMonthSheetModels = CalculateSheetModels());
        }

        private Collection<MonthSheetModel> CalculateSheetModels()
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

            return new Collection<MonthSheetModel>(monthSheets.OrderByDescending(it => it.Order).ToList());
        }

        private Collection<SummarySheetRowModel> CalculateSummarySheetRowModels(
            Collection<MonthSheetModel> inputMonthSheetModels)
        {
            var result = new Collection<SummarySheetRowModel>();

            foreach (var inputMonthSheetModel in inputMonthSheetModels)
            {
                var monthName = inputMonthSheetModel.Name;

                var dict = inputMonthSheetModel.Columns.Zip(inputMonthSheetModel.ColumnsTotals,
                    (key, value) => new SummarySheetRowModel
                    {
                        AgentCode = key,
                        NumberOfClients = value,
                        Month = monthName
                    });

                foreach (var it in dict)
                {
                    result.Add(it);
                }
            }

            return new Collection<SummarySheetRowModel>(result
                .OrderBy(it => it.AgentCode)
                .ThenBy(it => it.MonthId).ToList());
        } 

        private IEnumerable<Record> FilteredRecords(List<Record> input)
        {
            var filteredCollection = new Collection<Record>();

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