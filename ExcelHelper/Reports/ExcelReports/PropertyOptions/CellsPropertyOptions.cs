using System.Drawing;

namespace ExcelHelper.Reports.ExcelReports.PropertyOptions
{
    public class CellsPropertyOptions : PropertyOption
    {
        public CellsPropertyOptions(Location location) : base(location) { }
        public Size Size { get; set; }
    }
}
