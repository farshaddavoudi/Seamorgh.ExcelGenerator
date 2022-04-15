using ClosedXML.Excel;
using DNTPersianUtils.Core;
using ExcelHelper.Reports.ExcelReports;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelGenerator
{
    public static class ExcelService
    {
        /// <summary>
        /// Generate Excel file into file result
        /// </summary>
        /// <param name="excelFileModel"></param>
        /// <returns></returns>
        public static GeneratedExcelFile GenerateExcel(ExcelFileModel excelFileModel)
        {
            using var xlWorkbook = GenerateExcelFromModel(excelFileModel);

            // Save
            using var stream = new MemoryStream();
            xlWorkbook.SaveAs(stream);
            var content = stream.ToArray();
            return new GeneratedExcelFile { Content = content, FileName = excelFileModel.FileName };
        }

        /// <summary>
        /// Generate Excel file and save it in path and return the saved url
        /// </summary>
        /// <param name="excelFileModel"></param>
        /// <param name="basePath"></param>
        /// <param name="excelFileNameWithoutXlsxExtension"></param>
        /// <returns></returns>
        public static string GenerateExcel(ExcelFileModel excelFileModel, string basePath, string excelFileNameWithoutXlsxExtension)
        {
            using var xlWorkbook = GenerateExcelFromModel(excelFileModel);

            var saveUrl = $"{basePath}\\{excelFileNameWithoutXlsxExtension}.xlsx";

            // Save
            xlWorkbook.SaveAs(saveUrl);

            return saveUrl;
        }

        private static XLWorkbook GenerateExcelFromModel(ExcelFileModel excelFileModel)
        {
            //-------------------------------------------
            //  Create Workbook (integrated with using statement)
            //-------------------------------------------
            var xlWorkbook = new XLWorkbook
            {
                RightToLeft = excelFileModel.SheetsDefaultStyles.IsRightToLeft,
                ColumnWidth = excelFileModel.SheetsDefaultStyles.ColumnWidth,
                RowHeight = excelFileModel.SheetsDefaultStyles.RowHeight
            };

            // Check sheet names are unique
            var sheetNames = excelFileModel.Sheets.Select(s => s.Name).ToList();

            var uniqueSheetNames = sheetNames.Distinct().ToList();

            if (sheetNames.Count != uniqueSheetNames.Count)
                throw new Exception("Sheet names should be unique");

            // Check any sheet available
            if (excelFileModel.Sheets.Count == 0)
                throw new Exception("No sheet is available to create Excel workbook");

            //-------------------------------------------
            //  Add Sheets one by one to ClosedXML Workbook instance
            //-------------------------------------------
            foreach (var sheet in excelFileModel.Sheets)
            {
                // Set name
                var xlSheet = xlWorkbook.Worksheets.Add(sheet.Name);

                // Set protection level
                var protection = xlSheet.Protect(sheet.SheetProtectionLevels.Password);
                if (sheet.SheetProtectionLevels.Deletecolumns)
                    protection.Protect().AllowedElements = XLSheetProtectionElements.DeleteColumns;
                if (sheet.SheetProtectionLevels.Editobjects)
                    protection.Protect().AllowedElements = XLSheetProtectionElements.EditObjects;
                if (sheet.SheetProtectionLevels.Formatcells)
                    protection.Protect().AllowedElements = XLSheetProtectionElements.FormatCells;
                if (sheet.SheetProtectionLevels.Formatcolumns)
                    protection.Protect().AllowedElements = XLSheetProtectionElements.FormatColumns;
                if (sheet.SheetProtectionLevels.Formatrows)
                    protection.Protect().AllowedElements = XLSheetProtectionElements.FormatRows;
                if (sheet.SheetProtectionLevels.Insertcolumns)
                    protection.Protect().AllowedElements = XLSheetProtectionElements.InsertColumns;
                if (sheet.SheetProtectionLevels.Inserthyperlinks)
                    protection.Protect().AllowedElements = XLSheetProtectionElements.InsertHyperlinks;
                if (sheet.SheetProtectionLevels.Insertrows)
                    protection.Protect().AllowedElements = XLSheetProtectionElements.InsertRows;
                if (sheet.SheetProtectionLevels.Selectlockedcells)
                    protection.Protect().AllowedElements = XLSheetProtectionElements.SelectLockedCells;
                if (sheet.SheetProtectionLevels.Deleterows)
                    protection.Protect().AllowedElements = XLSheetProtectionElements.DeleteRows;
                if (sheet.SheetProtectionLevels.Editscenarios)
                    protection.Protect().AllowedElements = XLSheetProtectionElements.EditScenarios;
                if (sheet.SheetProtectionLevels.Selectunlockedcells)
                    protection.Protect().AllowedElements = XLSheetProtectionElements.SelectUnlockedCells;
                if (sheet.SheetProtectionLevels.Sort)
                    protection.Protect().AllowedElements = XLSheetProtectionElements.Sort;
                if (sheet.SheetProtectionLevels.UseAutoFilter)
                    protection.Protect().AllowedElements = XLSheetProtectionElements.AutoFilter;
                if (sheet.SheetProtectionLevels.UsePivotTablereports)
                    protection.Protect().AllowedElements = XLSheetProtectionElements.PivotTables;

                // Set direction
                if (sheet.SheetStyle.IsRightToLeft is not null)
                    xlSheet.RightToLeft = (bool)sheet.SheetStyle.IsRightToLeft;

                // Set default column width
                if (sheet.SheetStyle.DefaultColumnWidth is not null)
                    xlSheet.ColumnWidth = (double)sheet.SheetStyle.DefaultColumnWidth;

                // Set default row height
                if (sheet.SheetStyle.DefaultRowHeight is not null)
                    xlSheet.RowHeight = (double)sheet.SheetStyle.DefaultRowHeight;

                // Set visibility
                xlSheet.Visibility = sheet.SheetStyle.Visibility switch
                {
                    SheetVisibility.Hidden => XLWorksheetVisibility.Hidden,
                    SheetVisibility.VeryHidden => XLWorksheetVisibility.VeryHidden,
                    _ => XLWorksheetVisibility.Visible
                };

                // Set TextAlign
                var textAlign = sheet.SheetStyle.SheetDefaultTextAlign ?? excelFileModel.SheetsDefaultStyles.TextAlign;

                xlSheet.Columns().Style.Alignment.Horizontal = textAlign switch
                {
                    TextAlign.Center => XLAlignmentHorizontalValues.Center,
                    TextAlign.Right => XLAlignmentHorizontalValues.Right,
                    TextAlign.Left => XLAlignmentHorizontalValues.Left,
                    TextAlign.Justify => XLAlignmentHorizontalValues.Justify,
                    _ => throw new ArgumentOutOfRangeException()
                };

                //-------------------------------------------
                //  Columns properties
                //-------------------------------------------
                foreach (var colProps in sheet.ColumnsStyle)
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
                        if (colProps.Width.CalculationType == ColumnWidthCalculationType.AdjustToContents)
                            xlSheet.Column(colProps.ColumnNo).AdjustToContents();

                        else
                            xlSheet.Column(colProps.ColumnNo).Width = (double)colProps.Width.Value!;
                    }

                    if (colProps.AutoFit)
                        xlSheet.Column(colProps.ColumnNo).AdjustToContents();

                    if (colProps.IsHidden)
                        xlSheet.Column(colProps.ColumnNo).Hide();

                    xlSheet.Column(colProps.ColumnNo).Style.Alignment
                        .SetHorizontal(columnAlignmentHorizontalValue);
                }

                //-------------------------------------------
                //  Map Tables
                //-------------------------------------------
                foreach (var table in sheet.SheetTables)
                {
                    foreach (var tableRow in table.TableRows)
                    {
                        xlSheet.ConfigureRow(tableRow, sheet.ColumnsStyle, sheet.IsSheetLocked ?? excelFileModel.SheetsDefaultIsLocked);
                    }

                    var tableRange = xlSheet.Range(table.StartCellLocation.Y, table.StartCellLocation.X,
                        table.EndCellLocation.Y, table.EndCellLocation.X);

                    // Config Outside-Border
                    XLBorderStyleValues? outsideBorder = GetXlBorderLineStyle(table.OutsideBorder.LineStyle);

                    if (outsideBorder is not null)
                    {
                        tableRange.Style.Border.SetOutsideBorder((XLBorderStyleValues)outsideBorder);
                        tableRange.Style.Border.SetOutsideBorderColor(XLColor.FromColor(table.OutsideBorder.Color));
                    }

                    // Config Inside-Border
                    XLBorderStyleValues? insideBorder = GetXlBorderLineStyle(table.InlineBorder.LineStyle);

                    if (insideBorder is not null)
                    {
                        tableRange.Style.Border.SetInsideBorder((XLBorderStyleValues)insideBorder);
                        tableRange.Style.Border.SetInsideBorderColor(XLColor.FromColor(table.InlineBorder.Color));
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
                foreach (var row in sheet.SheetRows)
                {
                    xlSheet.ConfigureRow(row, sheet.ColumnsStyle, sheet.IsSheetLocked ?? excelFileModel.SheetsDefaultIsLocked);
                }

                //-------------------------------------------
                //  Map Cells
                //-------------------------------------------
                foreach (var cell in sheet.SheetCells)
                {
                    if (cell.Visible is false)
                        continue;

                    xlSheet.ConfigureCell(cell, sheet.ColumnsStyle, sheet.IsSheetLocked ?? excelFileModel.SheetsDefaultIsLocked);
                }

                // Apply sheet merges here
                foreach (var mergedCells in sheet.MergedCells)
                {
                    var rangeToMerge = xlSheet.Range(mergedCells).Cells();

                    var value = rangeToMerge.FirstOrDefault(r => r.IsEmpty() is false)?.Value;

                    rangeToMerge.First().SetValue(value);

                    xlSheet.Range(mergedCells).Merge();
                }
            }

            return xlWorkbook;
        }

        private static void ConfigureCell(this IXLWorksheet xlSheet, Cell cell, List<ColumnStyle> columnProps, bool isSheetLocked)
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
                    xlDataType = XLDataType.Number;
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
            bool? isLocked = cell.IsLocked;

            if (isLocked is null)
            { // Get from ColumnProps level
                var x = cell.CellLocation.X;

                var relatedColumnProp = columnProps.SingleOrDefault(c => c.ColumnNo == x);

                isLocked = relatedColumnProp?.IsLocked;

                if (isLocked is null)
                { // Get from sheet level
                    isLocked = isSheetLocked;
                }
            }

            //-------------------------------------------
            //  Map column per Cells loop cycle
            //-------------------------------------------
            var locationCell = xlSheet.Cell(cell.CellLocation.Y, cell.CellLocation.X);

            if (xlDataType is not null)
                locationCell.SetDataType((XLDataType)xlDataType);

            if (cell.Category == Category.Formula)
                locationCell.SetFormulaA1(cellValue.ToString());
            else
                locationCell.SetValue(cellValue);

            locationCell.Style
                .Alignment.SetWrapText(cell.Wordwrap);

            locationCell.Style.Protection.SetLocked((bool)isLocked!);

            if (cellAlignmentHorizontalValue is not null)
                locationCell.Style.Alignment.SetHorizontal((XLAlignmentHorizontalValues)cellAlignmentHorizontalValue!);
        }

        private static void ConfigureRow(this IXLWorksheet xlSheet, Row row, List<ColumnStyle> columnProps, bool isSheetLocked)
        {
            foreach (var rowCell in row.Cells)
            {
                if (rowCell.Visible is false)
                    continue;

                xlSheet.ConfigureCell(rowCell, columnProps, isSheetLocked);
            }

            // Configure merged cells in the row
            foreach (var cellsToMerge in row.MergedCellsList)
            {
                // CellsToMerge example is "B2:D2"
                xlSheet.Range(cellsToMerge).Row(1).Merge();
            }

            if (row.Cells.Count != 0)
            {
                if (row.StartCellLocation is not null && row.EndCellLocation is not null)
                {
                    var xlRow = xlSheet.Row(row.Cells.First().CellLocation.Y);
                    if (row.Height is not null)
                        xlRow.Height = (double)row.Height;

                    var xlRowRange = xlSheet.Range(row.StartCellLocation.Y, row.StartCellLocation.X, row.EndCellLocation.Y,
                        row.EndCellLocation.X);
                    xlRowRange.Style.Font.SetFontColor(XLColor.FromColor(row.ForeColor));
                    xlRowRange.Style.Fill.SetBackgroundColor(XLColor.FromColor(row.BackColor));

                    XLBorderStyleValues? outsideBorder = GetXlBorderLineStyle(row.OutsideBorder.LineStyle);

                    if (outsideBorder is not null)
                    {
                        xlRowRange.Style.Border.SetOutsideBorder((XLBorderStyleValues)outsideBorder);
                        xlRowRange.Style.Border.SetOutsideBorderColor(
                            XLColor.FromColor(row.OutsideBorder.Color));
                    }

                    // TODO: For Inside border, the row should be considered as Ranged (like Table). I persume it is not important for this phase
                }
                else
                {
                    var xlRow = xlSheet.Row(row.Cells.First().CellLocation.Y);
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