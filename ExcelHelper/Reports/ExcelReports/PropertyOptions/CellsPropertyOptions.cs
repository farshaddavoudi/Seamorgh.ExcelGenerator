using System.Drawing;

namespace ExcelHelper.Reports.ExcelReports.PropertyOptions
{
    public class CellsPropertyOptions : PropertyOption
    {
        public CellsPropertyOptions(CellLocation cellLocation) : base(cellLocation) { }
        public Size Size { get; set; }
    }
}
