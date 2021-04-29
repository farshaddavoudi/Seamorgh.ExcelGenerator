using ExcelHelper.ReportObjects;
using ExcelHelper.Reports.ExcelReports.PropertyOptions;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace ExcelHelper.Reports.ExcelReports.Template
{
    public static class TableTemplates
    {
        public static Table AccountHeader(Location startLocation)
        {
            Table table = new();
            ExcelReportBuilder builder = new();
            Border border = new(LineStyle.Continuous, Color.Black);
            var row = builder.AddRow(new List<string> { "نام حساب", "کد حساب" }, new RowPropertyOptions(startLocation));
            var location = row.NextVerticalLocation;
            var emtyrow = builder.EmptyRows(new List<string> { "", "" }, new RowPropertyOptions(location));
           // location = emtyrow.LastOrDefault().NextHorizontalLocation;

            row.BackColor = Color.DarkBlue;
            row.ForeColor = Color.White;
            row.MergedCells.Add("A17:A18");
            row.MergedCells.Add("B17:B18");
            row.InLineBorder = border;
            row.OutLineBorder = border;
            table.Rows.Add(row);
            table.Rows.AddRange(emtyrow);
            return table;
        }

        public static Table Multiplex(List<SummaryAccount> summary,Location currentLocation)
        {
            Table table = new();
            ExcelReportBuilder builder = new();
            Border border = new(LineStyle.Continuous, Color.Black);
            foreach (var item in summary)
            {
                var row=builder.AddRow(summary, new RowPropertyOptions(currentLocation));
                currentLocation = row.NextHorizontalLocation;
                row.BackColor = Color.DarkBlue;
                row.ForeColor = Color.White;
                row.InLineBorder = border;
                row.OutLineBorder = border;
                table.Rows.Add(row);
                foreach (var result in item.Multiplex)
                {
                    var header = builder.AddRow(new List<string> { "قبل از تسهیم", "بعد از تسهیم", }, new RowPropertyOptions(currentLocation));
                    table.Rows.Add(header);
                    currentLocation = header.NextVerticalLocation;
                    var childrow = builder.AddRow(item.Multiplex, new RowPropertyOptions(currentLocation));
                    row.BackColor = Color.DarkBlue;
                    row.ForeColor = Color.White;
                    ///
                    ///Adding Column For Formulas
                    ///
                    childrow.Formulas = $"{childrow.GetColumn(childrow.StartLocation.X).Location.GetName()}:{childrow.GetColumn(childrow.EndLocation.X).Location.GetName()}";
                    var sumcolum = childrow.AddColumn();
                    sumcolum.Value = childrow.Formulas;
                    ////////
                    ///

                    table.Rows.Add(childrow);
                }
                currentLocation = row.NextHorizontalLocation;
            }

            return table;
        }

        public static Table Accounts(List<AccountDto> accounts)
        {
            throw new NotImplementedException();
        }
    }
}
