using System.Drawing;

namespace ExcelHelper.Reports.ExcelReports
{
    public class Border
    {
        public Border(LineStyle borderLineStyle, Color borderColor)
        {
            BorderLineStyle = borderLineStyle;
            BorderColor = borderColor;
        }

        public LineStyle BorderLineStyle { get; set; }
        public Color BorderColor { get; set; }
    }
}
