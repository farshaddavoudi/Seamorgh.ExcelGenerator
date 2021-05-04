namespace ExcelHelper.Reports.ExcelReports
{
    public class WBProps
    {
        public bool IsRightToLeft { get; set; } = true;

        /// <summary>
        /// Default column width for the workbook.
        /// <para>All new worksheets will use this column width.</para>
        /// </summary>
        public double DefaultColumnWidth { get; set; } = 20;

        /// <summary>
        /// Default row height for the workbook.
        /// <para>All new worksheets will use this row height.</para>
        /// </summary>
        public double DefaultRowHeight { get; set; } = 15;
    }
}