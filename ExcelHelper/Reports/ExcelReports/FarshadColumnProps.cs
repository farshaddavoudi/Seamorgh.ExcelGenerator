using System.Collections.Generic;

namespace ExcelHelper.Reports.ExcelReports
{
    public class FarshadColumnProps
    {
        // TODO: Also other properties can be set for entire column

        public SetWidthScope SetWidthScope { get; set; } = SetWidthScope.PerColumn;

        public List<ColumnProps> ColumnPropsList { get; set; } = new();

        public ColumnWidth DefaultColumnWidth { get; set; } = new();
    }

    public enum SetWidthScope
    {
        AllColumns,
        PerColumn
    }

    public enum SetColumnWidthType
    {
        ExplicitValue,
        AdjustToContents
    }

    public class ColumnWidth
    {
        public SetColumnWidthType SetColumnWidthType { get; set; } = SetColumnWidthType.AdjustToContents;
        public double? Width { get; set; }
    }

    public class ColumnProps : ColumnWidth
    {
        public int ColumnNumber { get; set; }
        public bool IsHidden { get; set; }
    }

}