using System.Drawing;

namespace ExcelHelper.Reports.ExcelReports
{
    public class Border
    {
        public Border(LineStyle lineStyle, Color color)
        {
            LineStyle = lineStyle;
            Color = color;
        }

        public LineStyle LineStyle { get; set; }
        public Color Color { get; set; }
    }
}
