using ExcelHelper.ReportObjects;
using ExcelHelper.Reports.ExcelReports.PropertyOptions;
using System.Collections.Generic;
using System.Drawing;

namespace ExcelHelper.Reports.ExcelReports.Template
{
    public static class TableTemplates
    {
        public static Table AccountHeader(Location startLocation)
        {
            Table table = new();
            ExcelReportBuilder builder = new();
            Border border = new(LineStyle.Thick, Color.Black);
            var row = builder.AddRow(new List<string> { "نام حساب", "کد حساب" }, new RowPropertyOptions(startLocation));
            var location = row.NextVerticalLocation;
            var emtyrow = builder.EmptyRows(new List<string> { "", "" }, new RowPropertyOptions(location));
            // location = emtyrow.LastOrDefault().NextHorizontalLocation;

            row.BackColor = Color.DarkBlue;
            row.ForeColor = Color.White;
            //row.MergedCellsList.Add("A17:A18");
            //row.MergedCellsList.Add("B17:B18");
            table.MergedCells.Add("A17:A18");
            table.MergedCells.Add("B17:B18");
            row.AllBorder = border;
            row.OutsideBorder = border;
            table.Rows.Add(row);
            table.Rows.AddRange(emtyrow);
            return table;
        }

        public static Table Multiplex(List<SummaryAccount> summary, Location currentLocation)
        {
            Table table = new();
            List<string> Sumcells = new();
            ExcelReportBuilder builder = new();
            Border border = new(LineStyle.Thick, Color.Black);
            Cell sumColumn;
            var row = builder.AddRow(summary, new RowPropertyOptions(currentLocation), 2);
            currentLocation = row.NextVerticalLocation;
            row.BackColor = Color.DarkBlue;
            Row childrow=new();
            row.ForeColor = Color.White;
            row.AllBorder = border;
            row.OutsideBorder = border;
            table.Rows.Add(row);
            foreach (var item in summary)
            {

                foreach (var result in item.Multiplex)
                {
                    var header = builder.AddRow(new List<string> { "قبل از تسهیم", "بعد از تسهیم", "جمع" }, new RowPropertyOptions(currentLocation));
                    table.Rows.Add(header);
                    currentLocation = header.NextVerticalLocation;
                    childrow = builder.AddRow(item.Multiplex, new RowPropertyOptions(currentLocation));
                    row.BackColor = Color.DarkBlue;
                    row.ForeColor = Color.White;
                    header.BackColor = Color.DarkBlue;
                    header.ForeColor = Color.White;

                    // Add Cell For Formulas
                    childrow.Formulas = $"=sum({childrow.GetCell(childrow.StartLocation.X).Location.GetName()}:{childrow.GetCell(childrow.EndLocation.X).Location.GetName()})";
                    sumColumn = childrow.AddCell();
                    sumColumn.Category = Category.Formula;
                    sumColumn.Value = childrow.Formulas;
                    Sumcells.Add(sumColumn.Location.GetName());
                    ////////

                    table.Rows.Add(childrow);
                    currentLocation = new Location(childrow.NextHorizontalLocation.X, header.EndLocation.Y);
                    var avgtitle = header.AddCell();
                    avgtitle.Value = "میانگین";
                }
                var avg = childrow.AddCell();
                string avgstr = "=(";
                for (int i = 0; i < Sumcells.Count; i++)
                {
                    avgstr += Sumcells[i];
                    if (i < Sumcells.Count-1)
                        avgstr += "+";
                }
                avgstr += ")/"+ Sumcells.Count + "";
                avg.Value = avgstr;
                avg.Category = Category.Formula;
                
            }
            return table;
        }

        public static Table Accounts(List<AccountDto> accounts, Location currentLocation)
        {
            Table table = new();
            ExcelReportBuilder builder = new();
            Border border = new(LineStyle.Thick, Color.Black);

            var childrow = builder.AddTable(accounts, new TablePropertyOptions(currentLocation));
            table = childrow;
            table.InlineBorder = border;
            table.OutsideBorder = border;
            return table;
        }
    }
}
