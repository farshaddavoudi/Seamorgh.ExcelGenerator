using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;

namespace ExcelHelper.Reports.ExcelReports
{
    public class Row
    {
        public Row()
        {
            Cells = new List<Cell>();
            MergedCells = new List<string>();
        }

        public DataRowCollection Data { get; set; }
        public Location StartLocation { get; set; }
        public Location EndLocation { get; set; }
        public Color BackColor { get; set; } = Color.White;
        public Color ForeColor { get; set; } = Color.Black;
        public List<Cell> Cells { get; set; }
        public int Height { get; set; }
        public List<string> MergedCells { get; set; }
        public Border InLineBorder { get; set; }
        public Border OutLineBorder { get; set; }
        public bool IsBordered { get; set; }
        public string Formulas { get; set; }
        public int CellsCount => Cells.Count;

        public Location NextHorizontalLocation
        {
            get
            {
                var y = EndLocation.Y - (EndLocation.Y - StartLocation.Y);
                return new Location(EndLocation.X + 1, y);

            }
        }

        public Cell AddCell()
        {
            Cell cell = new(NextVerticalLocation);
            Cells.Add(cell);
            return cell;
        }

        public Location NextVerticalLocation
        {
            get
            {
                var x = EndLocation.X - (EndLocation.X - StartLocation.X);
                return new Location(x, EndLocation.Y + 1);

            }
        }

        public Cell GetCell(int X)
        {
            return Cells.FirstOrDefault(x => x.Location.X == X);
        }
    }
}
