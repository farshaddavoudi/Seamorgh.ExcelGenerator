using ExcelGenerator;
using ExcelHelper.Reports.ExcelReports;
using ExcelHelper.Reports.ExcelReports.Template;
using ExcelHelper.VoucherStatementReport;
using System.Collections.Generic;
using System.Drawing;

namespace MyApp
{
    public static class ExcelReportGenerator
    {
        public static string VoucherStatementExcelReportUrl(VoucherStatementPageResult result, string basePath, string excelName)
        {
            var workBook = new ExcelFileModel { FileName = "FileName" };
            var sheet1 = SheetTemplates.VoucherStatementTemplate(result);
            workBook.Sheets.Add(sheet1);

            // Generate Excel from "WorkBook" instance
            return ExcelService.GenerateExcel(workBook, basePath, excelName);
        }

        public static GeneratedExcelFile VoucherStatementExcelReport(VoucherStatementPageResult result)
        {
            var workBook = new ExcelFileModel { FileName = "FileName" };
            var sheet1 = SheetTemplates.VoucherStatementTemplate(result);
            workBook.Sheets.Add(sheet1);

            // Generate Excel from "WorkBook" instance
            return ExcelService.GenerateExcel(workBook);
        }

        public static GeneratedExcelFile TestReport()
        {
            var workbook = new ExcelFileModel
            {
                FileName = "TestName",
                SheetsDefaultStyles = new SheetsDefaultStyle { ColumnWidth = 40 },
                Sheets = new List<Sheet> { new("")
                    {
                        Name = "MySheet",
                        SheetTables = new List<Table>
                        {
                            new()
                            {
                                TableRows = new List<Row>
                                {
                                    new()
                                    {
                                        Cells = new List<Cell>
                                        {
                                            new(new CellLocation(3,5)){Value = "احمد", Category =Category.Text, TextAlign = TextAlign.Center}
                                        },
                                        MergedCellsList = new(){"C5:D5"},
                                        //StartLocation = new Location(3,5),
                                        //EndLocation = new Location(4,5),
                                        ForeColor = Color.DarkGreen,
                                        BackColor = Color.Aqua,
                                        OutsideBorder = new Border(LineStyle.DashDotDot, Color.Brown)
                                    },
                                    new()
                                    {
                                        Cells = new List<Cell>
                                        {
                                            new(new CellLocation(3,6)){Value = "کامبیز دیرباز", Category =Category.Text, TextAlign = TextAlign.Center}
                                        },
                                        MergedCellsList = new(){"C6:D6"},
                                        //StartLocation = new Location(3,6),
                                        //EndLocation = new Location(4,6),
                                        ForeColor = Color.DarkGreen,
                                        BackColor = Color.Aqua,
                                        OutsideBorder = new Border(LineStyle.DashDotDot, Color.Brown)
                                    },
                                    new()
                                    {
                                        Cells = new List<Cell>
                                        {
                                            new(new CellLocation(3,7)){Value = "اصغر فرهادی", Category =Category.Text, TextAlign = TextAlign.Center}
                                        },
                                        MergedCellsList = new(){"C7:D7"},
                                        //StartLocation = new Location(3,7),
                                        //EndLocation = new Location(4,7),
                                        ForeColor = Color.DarkGreen,
                                        BackColor = Color.Aqua,
                                        OutsideBorder = new Border(LineStyle.DashDotDot, Color.Brown)
                                    }
                                },
                                //StartLocation = new Location(3,5), //TODO: Can't be inferred from First Row StartLocation???
                                //EndLocation = new Location(4,7), //TODO: Can't be inferred from EndLocation of last Row???
                                OutsideBorder = new Border(LineStyle.Thick, Color.GreenYellow),
                                MergedCells = new List<string>{ "C5:D6" }
                            },

                        },
                        ColumnsStyle = new List<ColumnStyle>
                        {
                            new (){ColumnNo = 3,Width=new ColumnWidth(10)},
                            new(){ColumnNo = 1, IsLocked = true,Width = new ColumnWidth{CalculationType = ColumnWidthCalculationType.AdjustToContents}}
                        },
                        SheetRows = new List<Row>
                        {
                            new()
                            {
                                Cells = new List<Cell>
                                {
                                    new(new CellLocation(3,2)){Value = "فرشاد", Category =Category.Text, TextAlign = TextAlign.Right}
                                },
                                MergedCellsList = new(){"C2:D2"},
                                //StartLocation = new Location(2,2),
                                //EndLocation = new Location(4,2),
                                ForeColor = Color.BlueViolet,
                                BackColor = Color.AliceBlue,
                                OutsideBorder = new Border(LineStyle.DashDotDot, Color.Red)
                            }
    },
                        SheetCells = new List<Cell>
                        {
                            new Cell(new CellLocation("A",1)){Value = 11, Category = Category.Percentage, TextAlign = TextAlign.Left
},
                            new Cell(new CellLocation(2, 1)) { Value = 112343, Category = Category.Currency },
                            new Cell(new CellLocation("D", 1)) { Value = 112 },
                            new Cell(new CellLocation(1, 2)) { Value = 211, TextAlign = TextAlign.Center },
                            new Cell(new CellLocation(2, 2)) { Value = 212 },
                        },
                    }
                }
            };

            return ExcelService.GenerateExcel(workbook);
        }
    }
}
