using ExcelHelper.Reports.ExcelReports.PropertyOptions;
using ExcelHelper.VoucherStatementReport;
using System.Collections.Generic;
using System.Drawing;

namespace ExcelHelper.Reports.ExcelReports.Template
{
    public static class SheetTemplates
    {
        public static Sheet VoucherStatementTemplate(VoucherStatementPageResult result)
        {
            ExcelReportBuilder builder = new();
            Sheet sheet = new("RemainReport");
            sheet.IsSheetLocked = false;
            var row = builder.AddRow(new List<string> { "کد حساب", "بدهکار", "بستانکار" }, new RowPropertyOptions(new CellLocation("A", 3)));
            var cell = builder.AddCell(result.ReportName, "ReportName",
                new CellsPropertyOptions(new CellLocation("H", 1)));
            var table = builder.AddTable(result.RowResult, new TablePropertyOptions(new CellLocation("A", 4)));
            var currentLocation = table.NextVerticalCellLocation;
            var row2 = builder.AddRow(new List<string> { "کد حساب", "بدهکار", "بستانکار" }, new RowPropertyOptions(currentLocation));
            currentLocation = row2.NextVerticalCellLocation;
            var table2 = builder.AddTable(result.RowResult, new TablePropertyOptions(currentLocation));
            currentLocation = table2.NextVerticalCellLocation;
            var accountheader = TableTemplates.AccountHeader(currentLocation);
            currentLocation = accountheader.NextVerticalCellLocation;
            var accounts = TableTemplates.Accounts(result.Accounts, currentLocation);
            currentLocation = accounts.NextHorizontalCellLocation;
            currentLocation = new CellLocation(currentLocation.X, currentLocation.Y - 3);

            var multiplexHeader = TableTemplates.Multiplex(result.SummaryAccounts, currentLocation);


            Border border = new(LineStyle.Thick, Color.Black);
            row.BackColor = Color.Gray;
            row2.BackColor = Color.Gray;
            table.InlineBorder = border;
            table.OutsideBorder = border;
            table2.InlineBorder = border;
            table2.OutsideBorder = border;
            sheet.SheetTables.Add(table);
            sheet.SheetTables.Add(table2);
            sheet.SheetTables.Add(accountheader);
            sheet.SheetTables.Add(accounts);
            sheet.SheetTables.Add(multiplexHeader);
            sheet.SheetRows.Add(row);
            sheet.SheetRows.Add(row2);
            sheet.SheetCells.Add(cell);
            ColumnStyle column = new ColumnStyle();
            column.IsLocked = true;
            column.ColumnNo = 1;
            sheet.ColumnsStyle = new List<ColumnStyle> { column };
            sheet.MergedCells.Add("A1:H2");
            sheet.MergedCells.Add("L17:L18");
            return sheet;
        }
    }
}
