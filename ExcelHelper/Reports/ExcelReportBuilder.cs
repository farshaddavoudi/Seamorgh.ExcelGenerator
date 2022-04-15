using ExcelHelper.ReportObjects;
using ExcelHelper.Reports.ExcelReports;
using ExcelHelper.Reports.ExcelReports.PropertyOptions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelHelper.Reports
{
    public interface IExcelReportBuilder
    {
        EasyExcelModel AddFile(string path, string filename);
        Sheet AddSheet(string title);
        Row AddRow(object list, RowPropertyOptions options, int emptyCells = 0);
        List<Row> EmptyRows(object list, RowPropertyOptions options, int count = 1);
        Table AddTable(object rows, TablePropertyOptions options);
        Cell AddCell(object cell, string cellName, CellsPropertyOptions options);
        List<Cell> EmptyCells(CellsPropertyOptions options, int count = 1);
    }

    public class ExcelReportBuilder : IExcelReportBuilder
    {

        public EasyExcelModel AddFile(string path, string filename)
        {
            return null;
        }

        public Sheet AddSheet(string title)
        {
            return new(title);
        }

        /// <summary>
        /// Cells is "List" of any Objects
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public Row AddRow(object list, RowPropertyOptions options, int emptyCells = 0)
        {
            if (list is IEnumerable cells)
            {

                Row row = new();
                var location = new CellLocation(options.StartCellLocation.X, options.StartCellLocation.Y);
                foreach (var cell in cells)
                {
                    if (cell is string)
                    {
                        row.Cells.Add(AddCell(cell, "", new CellsPropertyOptions(new CellLocation(location.X, location.Y))));
                        location.X++;
                    }
                    else
                    {
                        PropertyInfo[] props = cell.GetType().GetProperties();
                        foreach (PropertyInfo prop in props)
                        {
                            if (cell != null)
                            {
                                var att = prop.GetCustomAttributes(true).Where(x => x is ExcelReportAttribute).FirstOrDefault();
                                if (att is ExcelReportAttribute attr)
                                {
                                    if (attr.Visible != false)
                                    {
                                        row.Cells.Add(AddCell(cell, prop.Name, new CellsPropertyOptions(new CellLocation(location.X, location.Y))));
                                        location.X++;
                                    }
                                }
                                else
                                {
                                    row.Cells.Add(AddCell(cell, prop.Name, new CellsPropertyOptions(new CellLocation(location.X, location.Y))));
                                    location.X++;
                                }
                            }
                        }
                    }
                    for (int i = 0; i < emptyCells; i++)
                    {
                        row.Cells.Add(AddCell("", "", new CellsPropertyOptions(new CellLocation(location.X, location.Y))));
                        location.X++;

                    }
                }
                return row;
            }
            return null;
        }

        private Row EmptyRow(object list, RowPropertyOptions options)
        {
            if (list is IEnumerable cells)
            {
                var location = new CellLocation(options.StartCellLocation.X, options.StartCellLocation.Y);
                Row row = new();
                foreach (var cell in cells)
                {
                    row.Cells.Add(AddCell(string.Empty, string.Empty, new CellsPropertyOptions(new CellLocation(location.X, location.Y))));
                    location.X++;
                }
                return row;
            }
            return null;
        }

        public List<Row> EmptyRows(object list, RowPropertyOptions options, int count = 1)
        {
            List<Row> rows = new();
            for (int i = 0; i < count; i++)
                rows.Add(EmptyRow(list, options));

            return rows;
        }

        public Table AddTable(object list, TablePropertyOptions options)
        {
            if (list is IEnumerable rows)
            {
                Table table = new();
                var location = options.StartCellLocation;
                //table.StartLocation = new Location(location.X, location.Y);
                //table.EndLocation = new Location(location.X, location.Y);
                foreach (var item in rows)
                {
                    table.TableRows.Add(AddRow(new List<object> { item }, new RowPropertyOptions(location)));
                    location.Y++;
                }
                //table.EndLocation = location;
                return table;
            }
            return null;
        }

        public Cell AddCell(object cell, string cellName, CellsPropertyOptions options)
        {
            if (cell is IEnumerable && !(cell is string)) return null; //TODO: Is it OK to say cellObj is string? Why not change arg param to string? Why Value in "Cell" model is object then?
            var col = ConfigCell(cell, cellName, options);
            return col;
        }

        private Cell EmptyCell(CellsPropertyOptions options) => ConfigCell(string.Empty, string.Empty, options);


        public List<Cell> EmptyCells(CellsPropertyOptions options, int count = 1)
        {
            List<Cell> cells = new();
            for (int i = 0; i < count; i++)
                cells.Add(EmptyCell(options));

            return cells;
        }

        private static Cell ConfigCell(object cellObj, string cellName, CellsPropertyOptions options)
        {
            Cell cell = new(options.StartCellLocation) { CellLocation = options.StartCellLocation };
            if (cellObj is string)
            {
                cell.Value = cellObj;
                ConfigByType(cellObj, cell);
                return cell;
            }

            cell.Value = GetPropValue(cellObj, cellName);
            cell.Type = cellObj.GetType();
            cell.Name = cellName;

            ConfigByType(cellObj, cell);
            ConfigByName(cellObj, cellName, cell);
            return cell;
        }

        private static void ConfigByName(object cellObj, string cellName, Cell cell)
        {

            PropertyInfo[] props = cellObj.GetType().GetProperties();
            foreach (PropertyInfo prop in props)
            {
                object[] attrs = prop.GetCustomAttributes(true);
                foreach (var item in attrs)
                {
                    var excelAttr = item as ExcelReportAttribute;
                    if (prop.Name == cell.Name)
                    {
                        cell.Visible = excelAttr.Visible;
                    }
                }
            }
            switch (cellName)
            {
                case "Debit":
                    cell.TextAlign = TextAlign.Right;
                    cell.Category = Category.Currency;
                    break;
                case "Credit":
                    cell.TextAlign = TextAlign.Right;
                    cell.Category = Category.Currency;
                    break;

                default:
                    break;
            }
        }

        private static void ConfigByType(object cellObj, Cell cell)
        {
            switch (Type.GetTypeCode(cellObj.GetType()))
            {
                case TypeCode.Decimal:
                    cell.TextAlign = TextAlign.Right;
                    cell.Category = Category.Currency;
                    break;
                case TypeCode.Int32:
                    cell.TextAlign = TextAlign.Right;
                    cell.Category = Category.Number;
                    break;
                case TypeCode.String:
                    cell.Wordwrap = true;
                    cell.Category = Category.Text;
                    break;
                default:
                    cell.TextAlign = TextAlign.Right;
                    cell.Category = Category.General;
                    break;
            }
        }

        private static object GetPropValue(object src, string propName)
        {
            return src.GetType().GetProperty(propName)?.GetValue(src, null);
        }
    }
}
