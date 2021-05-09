﻿using ExcelHelper.ReportObjects;
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
            Border border = new(LineStyle.Thick, Color.Black);
            var row = builder.AddRow(new List<string> { "نام حساب", "کد حساب" }, new RowPropertyOptions(startLocation));
            var location = row.NextVerticalLocation;
            var emtyrow = builder.EmptyRows(new List<string> { "", "" }, new RowPropertyOptions(location));
           // location = emtyrow.LastOrDefault().NextHorizontalLocation;

            row.BackColor = Color.DarkBlue;
            row.ForeColor = Color.White;
            row.MergedCellsList.Add("A17:A18");
            row.MergedCellsList.Add("B17:B18");
            row.AllBorder = border;
            row.OutsideBorder = border;
            table.Rows.Add(row);
            table.Rows.AddRange(emtyrow);
            return table;
        }

        public static Table Multiplex(List<SummaryAccount> summary,Location currentLocation)
        {
            Table table = new();
            ExcelReportBuilder builder = new();
            Border border = new(LineStyle.Thick, Color.Black);
            foreach (var item in summary)
            {
                var row=builder.AddRow(summary, new RowPropertyOptions(currentLocation));
                currentLocation = row.NextHorizontalLocation;
                row.BackColor = Color.DarkBlue;
                row.ForeColor = Color.White;
                row.AllBorder = border;
                row.OutsideBorder = border;
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
                    ///Adding Cell For Formulas
                    ///
                    childrow.Formulas = $"{childrow.GetCell(childrow.StartLocation.X).Location.GetName()}:{childrow.GetCell(childrow.EndLocation.X).Location.GetName()}";
                    var sumcolum = childrow.AddCell();
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
