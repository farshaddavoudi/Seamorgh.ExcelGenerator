using System.Drawing;

namespace ExcelHelper.Reports.ExcelReports.PropertyOptions
{
    public class CellsPropertyOptions : PropertyOption
    {
        public CellsPropertyOptions(Location startLocation) : base(startLocation) { }
        public string EndLocation { get; set; }
        public Size Size { get; set; }
    }
}
