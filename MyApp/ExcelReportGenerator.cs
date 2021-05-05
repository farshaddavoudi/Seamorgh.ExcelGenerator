using ExcelGenerator;
using ExcelHelper.Reports.ExcelReports;
using ExcelHelper.Reports.ExcelReports.Template;
using ExcelHelper.VoucherStatementReport;
using System.Collections.Generic;

namespace MyApp
{
    public static class ExcelReportGenerator
    {
        public static ExcelGeneratedFileResult VoucherStatementExcelReport(VoucherStatementPageResult result)
        {
            var workBook = new WorkBook { FileName = "FileName" };
            var sheet1 = SheetTemplates.VoucherStatementTemplate(result);
            workBook.Sheets.Add(sheet1);

            // Generate Excel from "WorkBook" instance
            return ExcelService.GenerateExcel(workBook);
        }

        public static ExcelGeneratedFileResult TestReport()
        {
            var workbook = new WorkBook
            {
                FileName = "TestName",
                WBProps = new WBProps { DefaultColumnWidth = 40 },
                Sheets = new List<Sheet> { new()
                    {
                        Name = "MySheet",
                        Cells = new List<Cell>
                        {
                            new Cell(new Location(1,1)){Value = 11, Category = Category.Percentage},
                            new Cell(new Location(2,1)){Value = 112343, Category = Category.Currency},
                            new Cell(new Location(3,1)){Value = 112},
                            new Cell(new Location(1,2)){Value = 211},
                            new Cell(new Location(2,2)){Value = 212},
                        },
                        ColumnPropsList = new List<ColumnProps>
                        {
                            new (){ColumnNo = 3,Width=new ColumnWidth(10)},
                            new(){ColumnNo = 1,Width = new ColumnWidth{CalculateType = ColumnWidthCalculateType.AdjustToContents}}
                        }
                    }
                }
            };

            return ExcelService.GenerateExcel(workbook);
        }
    }
}
