namespace ExcelHelper.Reports.ExcelReports
{
    public class WSProps
    {
        public bool? IsRightToLeft { get; set; } = null;

        /// <summary>
        /// Default column width for this worksheet.
        /// </summary>
        public double? DefaultColumnWidth { get; set; } = null;

        /// <summary>
        /// Default row height for this worksheet.
        /// </summary>
        public double? DefaultRowHeight { get; set; } = null;

        public SheetVisibility Visibility { get; set; } = SheetVisibility.Visible;
    }

    public enum SheetVisibility
    {
        Visible,
        // Can UnHide
        Hidden,
        // Can not be UnHide
        VeryHidden
    }
}

