namespace ExcelReportGenerator
{
    public class SummarySheetRowModel
    {
        private string _month;
        public string AgentCode { get; set; }

        public string Month
        {
            get { return _month; }
            set
            {
                _month = value;
                MonthId = MonthesAndIds.GetMonthIdByName(value);
            }
        }

        public int MonthId { get; set; }

        public long NumberOfClients { get; set; }
    }
}