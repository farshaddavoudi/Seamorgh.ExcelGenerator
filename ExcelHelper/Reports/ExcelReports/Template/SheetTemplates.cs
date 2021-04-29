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
            Sheet sheet = new("RemainReport", "RemainReport");
            var row = builder.AddRow(new List<string> { "کد حساب", "بدهکار", "بستانکار" }, new RowPropertyOptions(new Location("A", 3)));
            var column = builder.AddColumn(result.ReportName,"ReportName", new ColumnPropertyOptions(new Location("H", 1)));
            var table = builder.AddTable(result.RowResult, new TablePropertyOptions(new Location("A", 4)));
            var currentLocation = table.NextVerticalLocation;
            var row2 = builder.AddRow(new List<string> { "کد حساب", "بدهکار", "بستانکار" }, new RowPropertyOptions(currentLocation));
            currentLocation = new Location(currentLocation.X, currentLocation.Y + 1);
            var table2 = builder.AddTable(result.RowResult, new TablePropertyOptions(currentLocation));
            currentLocation = table2.NextHorizontalLocation;
            var accountheader = TableTemplates.AccountHeader(currentLocation);
            currentLocation = accountheader.NextVerticalLocation;
            var multiplexHeader = TableTemplates.Multiplex(result.SummaryAccounts, currentLocation);


            //var accounts=TableTemplates.Accounts(result.Accounts);

            Border border = new(LineStyle.Continuous, Color.Black);
            row.BackColor = Color.Gray;
            row2.BackColor = Color.Gray;
            table.InLineBorder = border;
            table.OutLineBorder = border;
            table2.InLineBorder = border;
            table2.OutLineBorder = border;
            sheet.Tables.Add(table);
            sheet.Tables.Add(table2);
            sheet.Tables.Add(accountheader);
            sheet.Tables.Add(multiplexHeader);
            sheet.Rows.Add(row);
            sheet.Rows.Add(row2);
            sheet.Columns.Add(column);

            sheet.MergedCells.Add("A1:C2");

            return sheet;
        }
    }
}
