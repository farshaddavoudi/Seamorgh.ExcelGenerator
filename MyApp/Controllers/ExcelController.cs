using ExcelHelper.ReportObjects;
using ExcelHelper.VoucherStatementReport;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.Linq;

namespace MyApp.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        [HttpGet("ExportExcel")]
        public IActionResult ExportExcel()
        {
            // TODO: Dear Shahab, please replace my naive example with something real and more practical to build an Excel like you sent to me
            var arg = new VoucherStatementPageResult
            {
                ReportName = "TestReport",
                SummaryAccounts = new List<SummaryAccount>
                {
                    new SummaryAccount {AccountName = "MyAccount"},
                    new SummaryAccount {AccountName = "MyAccount2"}
                      
                },
                
                RowResult = new List<VoucherStatementRowResult>
                {
                    new VoucherStatementRowResult
                    {
                        AccountCode = "Code1",
                        Credit = 2342,
                        Debit = 232
                    },
                    new VoucherStatementRowResult
                    {
                        AccountCode = "Code2",
                        Credit = 222,
                        Debit = 23333
                    }
                }
            };
            Multiplex multiplex = new Multiplex
            {
                After = 100000,
                Befor=50000
            };
            arg.SummaryAccounts.Where(x=>x.AccountName== "MyAccount").FirstOrDefault().Multiplex.Add(multiplex);
            arg.SummaryAccounts.Where(x => x.AccountName == "MyAccount2").FirstOrDefault().Multiplex.Add(multiplex);


            var result = ExcelReportGenerator.VoucherStatementExcelReport(arg);

            return File(result.Content, result.ContentType, result.FileName);
        }
    }
}