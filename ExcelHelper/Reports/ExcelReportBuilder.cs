using ExcelHelper.ReportObjects;
using ExcelHelper.Reports.ExcelReports;
using ExcelHelper.Reports.ExcelReports.PropertyOptions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;

namespace ExcelHelper.Reports
{
    public interface IExcelReportBuilder
    {
        WorkBook AddFile(string Path, string filename);
        Sheet AddSheet(string title);
        Row AddRow(object list, RowPropertyOptions options);
        List<Row> EmptyRows(object list, RowPropertyOptions options, int count = 1);
        Table AddTable(object rows, TablePropertyOptions options);
        Column AddColumn(object cell, string cellName, ColumnPropertyOptions options);
        List<Column> EmptyColumns(ColumnPropertyOptions options, int count = 1);
    }


    public class ExcelReportBuilder : IExcelReportBuilder
    {

        public WorkBook AddFile(string Path, string filename)
        {
            return null;
        }

        public Sheet AddSheet(string title)
        {
            return new(title, title);
        }

        /// <summary>
        /// Cells is "List" of any Objects
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public Row AddRow(object list, RowPropertyOptions options)
        {
            if (list is IEnumerable columns)
            {

                Row row = new();
                row.StartLocation = new Location(options.StartLocation.X, options.StartLocation.Y);
                row.EndLocation = new Location(options.StartLocation.X, options.StartLocation.Y);
                var location =new Location(options.StartLocation.X,options.StartLocation.Y);
                foreach (var column in columns)
                {
                    if (column is string)
                    {
                        row.Columns.Add(AddColumn(column, "", new ColumnPropertyOptions(location)));
                        location.X++;
                    }
                    else
                    {
                        PropertyInfo[] props = column.GetType().GetProperties();
                        foreach (PropertyInfo prop in props)
                        {
                            if (column != null)
                            {
                                row.Columns.Add(AddColumn(column, prop.Name, new ColumnPropertyOptions(new Location(location.X,location.Y))));
                                location.X++;
                            }
                        }
                    }
                }
                row.EndLocation = new Location(location.X-1, location.Y) ;
                return row;
            }
            return null;
        }

        private Row EmptyRow(object list, RowPropertyOptions options)
        {
            if (list is IEnumerable columns)
            {
                var location = options.StartLocation;
                Row row = new();
                foreach (var column in columns)
                {
                    row.Columns.Add(AddColumn(string.Empty, string.Empty, new ColumnPropertyOptions(location)));
                    location.X++;
                }
                row.EndLocation = location;
                return row;
            }
            return null;
        }

        public List<Row> EmptyRows(object list, RowPropertyOptions options, int count = 1)
        {
            List<Row> rows = new();
            for (int i = 0; i < count; i++)
                EmptyRow(list, options);

            return rows;
        }

        public Table AddTable(object list, TablePropertyOptions options)
        {
            if (list is IEnumerable rows)
            {
                Table table = new();
                var location = options.StartLocation;
                table.StartLocation = new Location(location.X, location.Y);
                table.EndLocation = new Location(location.X, location.Y);
                foreach (var item in rows)
                {
                    table.Rows.Add(AddRow(new List<object> { item }, new RowPropertyOptions(location)));
                    location.Y++;
                }
                table.EndLocation = location;
                return table;
            }
            return null;
        }

        public Column AddColumn(object cell, string cellName, ColumnPropertyOptions options)
        {
            if (cell is IEnumerable && !(cell is string)) return null;
            var col = ConfigColumn(cell, cellName, options);
            return col;
        }

        private Column EmptyColumn(ColumnPropertyOptions options) => ConfigColumn(string.Empty, string.Empty, options);


        public List<Column> EmptyColumns(ColumnPropertyOptions options, int count = 1)
        {
            List<Column> columns = new();
            for (int i = 0; i < count; i++)
                columns.Add(EmptyColumn(options));

            return columns;
        }

        private static Column ConfigColumn(object cell, string cellName, ColumnPropertyOptions options)
        {
            Column column = new(options.StartLocation);
            column.Location = options.StartLocation;
            if (cell is string)
            {
                column.Value = cell;
                ConfigByType(cell, column);
                return column;
            }
            else
            {
                column.Value = GetPropValue(cell, cellName);
                column.Type = cell.GetType();
                column.Name = cellName;

                ConfigByType(cell, column);
                ConfigByName(cell, cellName, column);
                return column;
            }
        }

        private static void ConfigByName(object cell, string cellName, Column column)
        {

            PropertyInfo[] props = cell.GetType().GetProperties();
            foreach (PropertyInfo prop in props)
            {
                object[] attrs = prop.GetCustomAttributes(true);
                foreach (var item in attrs)
                {
                    var excelAttr = item as ExcelReportAttribute;
                    if (prop.Name == column.Name)
                    {
                        column.Visible = excelAttr.Visible;
                    }
                }
            }
            switch (cellName)
            {
                case "Debit":
                    column.Align = TextAlign.rtl;
                    column.AutoFill = true;
                    column.Category = Category.Currency;
                    break;
                case "Credit":
                    column.Align = TextAlign.rtl;
                    column.AutoFill = false;
                    column.Category = Category.Currency;
                    break;

                default:
                    break;
            }
        }

        private static void ConfigByType(object cell, Column column)
        {
            switch (Type.GetTypeCode(cell.GetType()))
            {
                case TypeCode.Decimal:
                    column.Align = TextAlign.rtl;
                    column.Width = 20;
                    column.Category = Category.Currency;
                    break;
                case TypeCode.Int32:
                    column.Align = TextAlign.rtl;
                    column.Width = 10;
                    column.Category = Category.Number;
                    break;
                case TypeCode.String:
                    column.Width = 40;
                    column.Wordwrap = true;
                    column.Category = Category.Text;
                    break;
                default:
                    column.Align = TextAlign.rtl;
                    column.Category = Category.General;
                    break;
            }
        }

        private static object GetPropValue(object src, string propName)
        {
            return src.GetType().GetProperty(propName).GetValue(src, null);
        }
    }
}
