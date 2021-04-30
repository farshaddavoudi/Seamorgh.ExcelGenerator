using ClosedXML.Excel;
using ExcelHelper.Reports.ExcelReports;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelGenerator
{
    public static class ExcelService
    {
        public static ExcelGeneratedFileResult GenerateExcel(WorkBook workBook)
        {
            var fakeReport = new List<string> { "Ahmad", "Zabih", "Ramin", "Marzieh", "Reza", "Ahmad" };

            try
            {
                //-------------------------------------------
                //  Create Workbook (integrated with using statement)
                //-------------------------------------------
                using var xlWorkbook = new XLWorkbook { RightToLeft = true };

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
                    var xlSheet = xlWorkbook.Worksheets.Add(sheet.Name);

                    //-------------------------------------------
                    //  Set Columns Width
                    //-------------------------------------------
                    xlSheet.Column(1).AdjustToContents();

                    //-------------------------------------------
                    //  Map Columns
                    //-------------------------------------------
                    foreach (var column in sheet.Columns)
                    {
                        // Infer XLDataType from "column"
                        XLDataType xlDataType;
                        switch (column.Category)
                        {
                            case Category.Number:
                                xlDataType = XLDataType.Number;
                                break;
                            case Category.Date:
                                xlDataType = XLDataType.DateTime;
                                break;

                            // TODO: Complete the rest after finding out what is Category about?

                            default:
                                xlDataType = XLDataType.Text;
                                break;
                        }

                        // Infer XLAlignment from "column"
                        XLAlignmentHorizontalValues xlAlignmentHorizontalValues;

                        switch (column.Align)
                        {
                            case TextAlign.Center:
                                xlAlignmentHorizontalValues = XLAlignmentHorizontalValues.Center;
                                break;

                            case TextAlign.Ltr:
                                xlAlignmentHorizontalValues = XLAlignmentHorizontalValues.Left;
                                break;

                            case TextAlign.Rtl:
                                xlAlignmentHorizontalValues = XLAlignmentHorizontalValues.Right;
                                break;

                            default:
                                xlAlignmentHorizontalValues = XLAlignmentHorizontalValues.Justify;
                                break;
                        }

                        //-------------------------------------------
                        //  Map column per Columns loop cycle
                        //-------------------------------------------
                        xlSheet
                            .Cell(column.Location.Y, column.Location.X)
                            .SetDataType(xlDataType)
                            .SetValue(column.Value)
                            .Style.Alignment.SetWrapText(column.Wordwrap)
                            .Alignment.SetHorizontal(xlAlignmentHorizontalValues);
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