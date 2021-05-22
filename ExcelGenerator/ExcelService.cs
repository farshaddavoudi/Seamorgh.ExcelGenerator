using ClosedXML.Excel;
using DNTPersianUtils.Core;
using ExcelHelper.Reports.ExcelReports;
using System;
using System.IO;
using System.Linq;

namespace ExcelGenerator
{
    public static class ExcelService
    {
        public static ExcelGeneratedFileResult GenerateExcel(WorkBook workBook)
        {
            try
            {
                //-------------------------------------------
                //  Create Workbook (integrated with using statement)
                //-------------------------------------------
                using var xlWorkbook = new XLWorkbook
                {
                    RightToLeft = workBook.WBProps.IsRightToLeft,
                    ColumnWidth = workBook.WBProps.DefaultColumnWidth,
                    RowHeight = workBook.WBProps.DefaultRowHeight
                };

                // Check sheet names are unique
                var sheetNames = workBook.Sheets.Select(s => s.Name).ToList();

                var uniqueSheetNames = sheetNames.Distinct().ToList();

                if (sheetNames.Count != uniqueSheetNames.Count)
                    throw new Exception("Sheet names should be unique");

                // Check any sheet available
                if (workBook.Sheets.Count == 0)
                    throw new Exception("No sheet is available to create Excel workbook");

                //-------------------------------------------
                //  Add Sheets one by one to ClosedXML Workbook instance
                //-------------------------------------------
                foreach (var sheet in workBook.Sheets)
                {
                    // Set name
                    var xlSheet = xlWorkbook.Worksheets.Add(sheet.Name);

                    // Set protection level
                    xlSheet.Protect().AllowedElements = XLSheetProtectionElements.SelectUnlockedCells;

                    // Set direction
                    if (sheet.WSProps.IsRightToLeft is not null)
                        xlSheet.RightToLeft = (bool)sheet.WSProps.IsRightToLeft;

                    // Set default column width
                    if (sheet.WSProps.DefaultColumnWidth is not null)
                        xlSheet.ColumnWidth = (double)sheet.WSProps.DefaultColumnWidth;

                    // Set default row height
                    if (sheet.WSProps.DefaultRowHeight is not null)
                        xlSheet.RowHeight = (double)sheet.WSProps.DefaultRowHeight;

                    // Set visibility
                    xlSheet.Visibility = sheet.WSProps.Visibility switch
                    {
                        SheetVisibility.Hidden => XLWorksheetVisibility.Hidden,
                        SheetVisibility.VeryHidden => XLWorksheetVisibility.VeryHidden,
                        _ => XLWorksheetVisibility.Visible
                    };

                    xlSheet.Columns().Style.Alignment.Horizontal = sheet.WSProps.DefaultTextAlign switch
                    {
                        TextAlign.Center => XLAlignmentHorizontalValues.Center,
                        TextAlign.Right => XLAlignmentHorizontalValues.Right,
                        TextAlign.Left => XLAlignmentHorizontalValues.Left,
                        TextAlign.Justify => XLAlignmentHorizontalValues.Justify,
                        _ => throw new ArgumentOutOfRangeException()
                    };

                    // Apply sheet merges here
                    foreach (var mergedCells in sheet.MergedCells)
                    {
                        xlSheet.Range(mergedCells).Merge();
                    }

                    //-------------------------------------------
                    //  Columns properties
                    //-------------------------------------------
                    foreach (var colProps in sheet.Columns)
                    {
                        // Infer XLAlignment from "ColumnProp"
                        var columnAlignmentHorizontalValue = colProps.TextAlign switch
                        {
                            TextAlign.Center => XLAlignmentHorizontalValues.Center,
                            TextAlign.Justify => XLAlignmentHorizontalValues.Justify,
                            TextAlign.Left => XLAlignmentHorizontalValues.Left,
                            TextAlign.Right => XLAlignmentHorizontalValues.Right,
                            _ => throw new ArgumentOutOfRangeException()
                        };

                        if (colProps.Width is not null)
                        {
                            if (colProps.Width.CalculateType == ColumnWidthCalculateType.AdjustToContents)
                                xlSheet.Column(colProps.ColumnNo).AdjustToContents();

                            else
                                xlSheet.Column(colProps.ColumnNo).Width = (double)colProps.Width.Value!;
                        }

                        if (colProps.IsHidden)
                            xlSheet.Column(colProps.ColumnNo).Hide();

                        xlSheet.Column(colProps.ColumnNo).Style.Alignment
                            .SetHorizontal(columnAlignmentHorizontalValue);
                    }

                    //-------------------------------------------
                    //  Map Tables
                    //-------------------------------------------
                    foreach (var table in sheet.Tables)
                    {
                        foreach (var tableRow in table.Rows)
                        {
                            xlSheet.ConfigureRow(tableRow, sheet.IsLocked);
                        }

                        var tableRange = xlSheet.Range(table.StartLocation.Y, table.StartLocation.X,
                            table.EndLocation.Y, table.EndLocation.X);

                        XLBorderStyleValues? outsideBorder = GetXlBorderLineStyle(table.OutsideBorder.LineStyle);

                        if (outsideBorder is not null)
                        {
                            tableRange.Style.Border.SetOutsideBorder((XLBorderStyleValues)outsideBorder);
                            tableRange.Style.Border.SetOutsideBorderColor(XLColor.FromColor(table.OutsideBorder.Color));
                        }

                        // Apply table merges here
                        foreach (var mergedCells in table.MergedCells)
                        {
                            xlSheet.Range(mergedCells).Merge();
                        }

                    }

                    //-------------------------------------------
                    //  Map Rows 
                    //-------------------------------------------
                    foreach (var row in sheet.Rows)
                    {
                        xlSheet.ConfigureRow(row, sheet.IsLocked);
                    }

                    //-------------------------------------------
                    //  Map Cells
                    //-------------------------------------------
                    foreach (var cell in sheet.Cells)
                    {
                        if (cell.Visible is false)
                            continue;

                        xlSheet.ConfigureCell(cell, sheet.IsLocked);
                    }
                }

                // Save
                using var stream = new MemoryStream();
                xlWorkbook.SaveAs(stream);
                var content = stream.ToArray();
                return new ExcelGeneratedFileResult { Content = content, FileName = workBook.FileName };
            }
            catch (Exception e)
            {
                // ignored
                throw;
            }
        }

        private static void ConfigureCell(this IXLWorksheet xlSheet, Cell cell, bool isSheetLocked)
        {
            // Infer XLDataType and value from "cell" category
            XLDataType? xlDataType;
            object cellValue = cell.Value;

            switch (cell.Category)
            {
                case Category.Number:
                    xlDataType = XLDataType.Number;
                    break;

                case Category.Percentage:
                    xlDataType = XLDataType.Text;
                    cellValue = $"{cellValue}%";
                    break;

                case Category.Currency:
                    xlDataType = XLDataType.Text;
                    if (cellValue.IsNumber() is false)
                        throw new Exception("Cell with Currency category should be Number type");
                    cellValue = Convert.ToDecimal(cellValue).ToString("##,###");
                    break;

                case Category.MiladiDate:
                    xlDataType = XLDataType.DateTime;
                    if (cellValue is not DateTime)
                        throw new Exception("Cell with MiladiDate category should be DateTime type");
                    break;

                case Category.SolarHijriDate:
                    if (cellValue is not DateTime)
                        throw new Exception("Cell with SolarHijriDate category should be DateTime type");
                    cellValue = Convert.ToDateTime(cellValue).ToShortPersianDateString();
                    xlDataType = XLDataType.Text;
                    break;

                case Category.Text:
                case Category.Formula:
                    xlDataType = XLDataType.Text;
                    break;

                default: // = Category.General
                    xlDataType = null;
                    break;
            }

            // Infer XLAlignment from "cell"
            XLAlignmentHorizontalValues? cellAlignmentHorizontalValue = cell.TextAlign switch
            {
                TextAlign.Center => XLAlignmentHorizontalValues.Center,
                TextAlign.Left => XLAlignmentHorizontalValues.Left,
                TextAlign.Right => XLAlignmentHorizontalValues.Right,
                TextAlign.Justify => XLAlignmentHorizontalValues.Justify,
                _ => null
            };

            // Get IsLocked property based on Sheet and Cell "IsLocked" prop
            bool isLocked = cell.IsLocked ?? isSheetLocked;

            //-------------------------------------------
            //  Map column per Cells loop cycle
            //-------------------------------------------
            var locationCell = xlSheet.Cell(cell.Location.Y, cell.Location.X);

            if (cell.Category == Category.Formula)
                locationCell.SetFormulaA1(cellValue.ToString());
            else
                locationCell.SetValue(cellValue);

            locationCell.Style
                .Alignment.SetWrapText(cell.Wordwrap)
                .Protection.SetLocked(isLocked);

            if (cellAlignmentHorizontalValue is not null)
                locationCell.Style.Alignment.SetHorizontal((XLAlignmentHorizontalValues)cellAlignmentHorizontalValue!);

            if (xlDataType is not null)
                locationCell.SetDataType((XLDataType)xlDataType);
        }

        private static void ConfigureRow(this IXLWorksheet xlSheet, Row row, bool isSheetLocked)
        {
            foreach (var rowCell in row.Cells)
            {
                if (rowCell.Visible is false)
                    continue;

                xlSheet.ConfigureCell(rowCell, isSheetLocked);
            }

            // Configure merged cells in the row
            foreach (var cellsToMerge in row.MergedCellsList)
            {
                // CellsToMerge example is "B2:D2"
                xlSheet.Range(cellsToMerge).Row(1).Merge();
            }

            if (row.Cells.Count != 0)
            {
                if (row.StartLocation is not null && row.EndLocation is not null)
                {
                    var xlRow = xlSheet.Row(row.Cells.First().Location.Y);
                    if (row.Height is not null)
                        xlRow.Height = (double)row.Height;

                    var xlRowRange = xlSheet.Range(row.StartLocation.Y, row.StartLocation.X, row.EndLocation.Y,
                        row.EndLocation.X);
                    xlRowRange.Style.Font.SetFontColor(XLColor.FromColor(row.ForeColor));
                    xlRowRange.Style.Fill.SetBackgroundColor(XLColor.FromColor(row.BackColor));

                    XLBorderStyleValues? outsideBorder = GetXlBorderLineStyle(row.OutsideBorder.LineStyle);

                    if (outsideBorder is not null)
                    {
                        xlRowRange.Style.Border.SetOutsideBorder((XLBorderStyleValues)outsideBorder);
                        xlRowRange.Style.Border.SetOutsideBorderColor(
                            XLColor.FromColor(row.OutsideBorder.Color));
                    }
                    //xlRowRange.Style.Border.SetInsideBorder(XLBorderStyleValues.Thick);
                    //xlRowRange.Style.Border.SetTopBorder(XLBorderStyleValues.Thick);
                    //xlRowRange.Style.Border.SetRightBorder(XLBorderStyleValues.DashDotDot);
                }
                else
                {
                    var xlRow = xlSheet.Row(row.Cells.First().Location.Y);
                    if (row.Height is not null)
                        xlRow.Height = (double)row.Height;
                    xlRow.Style.Font.SetFontColor(XLColor.FromColor(row.ForeColor));
                    xlRow.Style.Fill.SetBackgroundColor(XLColor.FromColor(row.BackColor));
                    xlRow.Style.Border.SetOutsideBorder(XLBorderStyleValues.Dotted);
                    xlRow.Style.Border.SetInsideBorder(XLBorderStyleValues.Thick);
                    xlRow.Style.Border.SetTopBorder(XLBorderStyleValues.Thick);
                    xlRow.Style.Border.SetRightBorder(XLBorderStyleValues.DashDotDot);
                }
            }
        }

        private static XLBorderStyleValues? GetXlBorderLineStyle(LineStyle borderLineStyle)
        {
            return borderLineStyle switch
            {
                LineStyle.DashDotDot => XLBorderStyleValues.DashDotDot,
                LineStyle.Thick => XLBorderStyleValues.Thick,
                LineStyle.Thin => XLBorderStyleValues.Thin,
                LineStyle.Dotted => XLBorderStyleValues.Dotted,
                LineStyle.Double => XLBorderStyleValues.Double,
                LineStyle.DashDot => XLBorderStyleValues.DashDot,
                LineStyle.Dashed => XLBorderStyleValues.Dashed,
                LineStyle.SlantDashDot => XLBorderStyleValues.SlantDashDot,
                LineStyle.None => XLBorderStyleValues.None,
                _ => null
            };
        }
    }
}