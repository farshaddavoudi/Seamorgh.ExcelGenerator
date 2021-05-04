using System.Collections.Generic;

namespace ExcelHelper.Reports.ExcelReports
{
    public class ColumnsStyle
    {
        public List<ColumnStyle> CustomColumnStyleList { get; set; } = new();

        public ColumnWidth DefaultColumnsWidth { get; set; } = new();
    }

    public enum SetColumnWidthType
    {
        ExplicitValue,
        AdjustToContents
    }

    public class ColumnStyle : ColumnWidth
    {
        public int ColumnNumber { get; set; }
        public bool IsHidden { get; set; }
    }

    public class ColumnWidth
    {
        public SetColumnWidthType SetColumnWidthType { get; set; } = SetColumnWidthType.ExplicitValue;
        public double? Width { get; set; }
    }
}