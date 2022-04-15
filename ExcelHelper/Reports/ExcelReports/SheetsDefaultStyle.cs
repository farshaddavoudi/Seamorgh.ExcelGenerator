namespace ExcelHelper.Reports.ExcelReports
{
    public class SheetsDefaultStyle
    {
        public bool IsRightToLeft { get; set; } = true;

        public TextAlign TextAlign { get; set; } = TextAlign.Right;

        /// <summary>
        /// Default column width for the workbook.
        /// <para>All new worksheets will use this column width.</para>
        /// </summary>
        public double ColumnWidth { get; set; } = SheetDefaults.ColumnWidth;

        /// <summary>
        /// Default row height for the workbook.
        /// <para>All new worksheets will use this row height.</para>
        /// </summary>
        public double RowHeight { get; set; } = SheetDefaults.RowHeight;
    }
}