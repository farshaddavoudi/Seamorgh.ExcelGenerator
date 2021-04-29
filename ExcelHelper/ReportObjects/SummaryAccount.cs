using System.Collections.Generic;

namespace ExcelHelper.ReportObjects
{
    public class SummaryAccount : IPrintExcel
    {
        public SummaryAccount()
        {
            Multiplex = new();
        }
        public string AccountName { get; set; }

        [ExcelReport(Visible = false)]
        public List<Multiplex> Multiplex { get; set; }
    }
}
