using ExcelGenerator;
using ExcelHelper.Reports.ExcelReports;
using ExcelHelper.Reports.ExcelReports.Template;
using ExcelHelper.VoucherStatementReport;

namespace MyApp
{
    public static class ExcelReportGenerator
    {
        public static ExcelGeneratedFileResult VoucherStatementExcelReport(VoucherStatementPageResult result)
        {
            var workBook = new WorkBook("FileName");
            var sheet1 = SheetTemplates.VoucherStatementTemplate(result);
            workBook.Sheets.Add(sheet1);

            // Generate Excel from "WorkBook" instance
            return ExcelService.GenerateExcel(workBook);
        }
    }
}
