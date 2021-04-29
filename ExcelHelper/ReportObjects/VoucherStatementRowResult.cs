using ExcelHelper.ReportObjects;

namespace ExcelHelper.VoucherStatementReport
{
    public class VoucherStatementRowResult
    {
        [ExcelReport(Name = "AccountCode")]

        public string AccountCode { get; set; }
        [ExcelReport(Name = "Debit")]
        public decimal Debit { get; set; }
        [ExcelReport(Name = "Credit")]
        public decimal Credit { get; set; }
    }
}
