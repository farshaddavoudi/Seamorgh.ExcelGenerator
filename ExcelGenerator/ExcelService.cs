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

                    // TODO: Make it dynamic by transfer it to Sheet model
                    xlSheet.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right; //Default align right for RTL direction

                    //-------------------------------------------
                    //  Set Columns custom Width
                    //-------------------------------------------
                    // Infer XLAlignment from "ColumnProp"
                    foreach (var colProps in sheet.ColumnPropsList)
                    {
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
                    //  Map Cells
                    //-------------------------------------------
                    foreach (var cell in sheet.Cells)
                    {
                        if (cell.Visible is false)
                            continue;

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

                        //-------------------------------------------
                        //  Map column per Cells loop cycle
                        //-------------------------------------------
                        var locationCell = xlSheet.Cell(cell.Location.Y, cell.Location.X);

                        locationCell.SetValue(cellValue)
                            .Style.Alignment.SetWrapText(cell.Wordwrap);

                        if (cellAlignmentHorizontalValue is not null)
                            locationCell.Style.Alignment.SetHorizontal((XLAlignmentHorizontalValues)cellAlignmentHorizontalValue!);

                        if (xlDataType is not null)
                            locationCell.SetDataType((XLDataType)xlDataType);
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
    }
}