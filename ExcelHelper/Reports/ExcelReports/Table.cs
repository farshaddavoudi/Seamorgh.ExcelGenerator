using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace ExcelHelper.Reports.ExcelReports
{
    public class Table : IValidatableObject
    {
        public List<Row> Rows { get; set; } = new();
        public Location StartLocation { get; set; } //TODO: Discuss with Shahab. The Rows has StartLocation itself, which one should be considered?
        //TODO: StartLocation and EndLocation for Table model are critical and should exist and be exact to create desired result
        public Location EndLocation { get; set; } //TODO: above question
        public Border InLineBorder { get; set; } //TODO: What it is? Inside border can be set on cells or columns or rows
        public Border OutsideBorder { get; set; }
        public bool IsBordered { get; set; } //TODO? What is this? isn't it the default one?
        public List<string> MergedCells { get; set; } = new();
        public int RowsCount => Rows.Count;

        public Location NextHorizontalLocation
        {
            get
            {
                var y = Rows.LastOrDefault().EndLocation.Y - (Rows.LastOrDefault().EndLocation.Y - Rows.LastOrDefault().StartLocation.Y);
                return new Location(Rows.LastOrDefault().EndLocation.X + 1, y);
            }
        }
        public Location NextVerticalLocation
        {
            get
            {
                var x = Rows.LastOrDefault().EndLocation.X - (Rows.LastOrDefault().EndLocation.X - Rows.LastOrDefault().StartLocation.X);
                return new Location(x, Rows.LastOrDefault().EndLocation.Y + 1);
            }
        }

        public Cell GetCell(Location location)
        {
            return Rows[location.X - 1].Cells[location.Y - 1];
        }

        public List<Cell> GetCells(Location startLocation, Location endLocation)
        {
            List<Cell> cells = new();
            for (int i = startLocation.Y; i < endLocation.Y; i++)
            {
                cells.Add(GetCell(new Location(startLocation.X, i)));
            }

            return cells;
        }

        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
        {
            if (false)
                yield return new ValidationResult("");
            // TODO: Discuess with Shahab. Shouldn't Rows in a Table have common features like Same StartLocation.X and things like
        }
    }
}
