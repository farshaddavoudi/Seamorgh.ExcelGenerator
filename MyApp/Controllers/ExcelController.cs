using ExcelHelper.ReportObjects;
using ExcelHelper.VoucherStatementReport;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;

namespace MyApp.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        [HttpGet("ExportExcel")]
        public IActionResult ExportExcel()
        {
            var arg = new VoucherStatementPageResult
            {
                ReportName = "TestReport",
                SummaryAccounts = new List<SummaryAccount>
                {
                    new SummaryAccount
                    {
                        AccountName = "کارخانه دان-51011" ,
                        Multiplex =new List<Multiplex>
                        {
                            new Multiplex{After = 5000000,Befor = 4000 },
                        },

                    },
                    new SummaryAccount
                    {
                        AccountName = "پرورش پولت-51018" ,
                        Multiplex =new List<Multiplex>
                        {
                            new Multiplex{After = 5000000,Befor = 4000 },
                        },

                    },
                    new SummaryAccount
                    {
                        AccountName = "تخم گزار تجاری-51035" ,
                        Multiplex =new List<Multiplex>
                        {
                            new Multiplex{After = 5000000,Befor = 4000 },
                        },
                    }

                },

                Accounts = new List<AccountDto>
                {
                   new AccountDto
                   {
                        Name="حقوق پایه",
                        Code="81010"
                   },
                   new AccountDto
                   {
                        Name="اضافه کار",
                        Code="81011"
                   },

                },
                RowResult = new List<VoucherStatementRowResult>
                {
                    new VoucherStatementRowResult
                    {
                        AccountCode = "13351",
                        Credit = 50000,
                        Debit = 0
                    },
                    new VoucherStatementRowResult
                    {
                        AccountCode = "21253",
                        Credit = 0,
                        Debit = 50000
                    },
                    new VoucherStatementRowResult
                    {
                        AccountCode = "13556",
                        Credit = 1000000,
                        Debit = 0
                    },
                    new VoucherStatementRowResult
                    {
                        AccountCode = "13500",
                        Credit = 1000000,
                        Debit = 0
                    },
                    new VoucherStatementRowResult
                    {
                        AccountCode = "13499",
                        Credit = 2000000,
                        Debit = 0
                    },
                    new VoucherStatementRowResult
                    {
                        AccountCode = "22500",
                        Credit = 0,
                        Debit = 4000000
                    }

                }
            };



            var result2 = ExcelReportGenerator.VoucherStatementExcelReport(arg);
            var result22Url = ExcelReportGenerator.VoucherStatementExcelReportUrl(arg, @"C:\LocalNugets", "MyCustomName");


            var result = ExcelReportGenerator.TestReport();

            return Ok(result22Url);
        }
    }
}