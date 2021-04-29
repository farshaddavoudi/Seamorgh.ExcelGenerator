using ExcelHelper.ReportObjects;
using System.Collections.Generic;

namespace ExcelHelper.VoucherStatementReport
{
    public class VoucherStatementPageResult : IPrintExcel
    {
        public VoucherStatementPageResult()
        {
            RowResult = new List<VoucherStatementRowResult>();
            Accounts = new List<AccountDto>();
        }
        public string ReportName { get; set; }
        public decimal FinalRemain { get; set; }

        public List<VoucherStatementRowResult> RowResult { get; set; }

        [ExcelReport(Visible = false)]
        public List<AccountDto> Accounts { get; set; }

        [ExcelReport(UseInFormulas = false , Visible = false)]
        public List<SummaryAccount> SummaryAccounts { get; set; }

    }
}
