using DNTPersianUtils.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Drawing;
using System.Linq;

namespace ExcelHelper.Reports.ExcelReports
{
    public class Row : IValidatableObject
    {
        public Row()
        {
            AllBorder = new Border(LineStyle.Non, Color.Black);
            OutsideBorder = new Border(LineStyle.Non, Color.Black);
            MergedCellsList = new();
        }
        public Location StartLocation 
        {
            get
            {
                return Cells.FirstOrDefault().Location ;
            }
        }
        public Location EndLocation 
        {
            get
            {
                return Cells.LastOrDefault().Location;
            }
        }
        public Color BackColor { get; set; } = Color.White;
        public Color ForeColor { get; set; } = Color.Black;
        // TODO: Add below props
        // Bold, FontName, FontSize, Italic, Shadow, StrikeThrough
        public List<Cell> Cells { get; set; } = new(); //TODO: Discuss with Shahab, can Cells.Count == 0 for a row? If not, add validation
        public double? Height { get; set; }
        public List<string> MergedCellsList { get; set; } 
        public Border AllBorder { get; set; } 
        public Border OutsideBorder { get; set; }
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
            Cell cell = new(NextHorizontalLocation);
            Cells.Add(cell);
            return cell;
        }

        public Location NextVerticalLocation
        {
            get
            {
                var x = EndLocation.X - (EndLocation.X - StartLocation.X); //TODO: ?? (x-(x-y) => answer always is y)
                return new Location(x, EndLocation.Y + 1);

            }
        }

        public Cell GetCell(int X)
        {
            return Cells.FirstOrDefault(x => x.Location.X == X);
        }

        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
        {
            // Check if StartLocation exist, EndLocation should exist too ////TODO: Discuss with Shahab is it true validation or not
            if (StartLocation is not null && EndLocation is null ||
                StartLocation is null && EndLocation is not null)
                yield return new ValidationResult(
                    "Both StartLocation and EndLocation of Row should be null or not null simultaneously");

            // TODO: Check Y of StartLocation and EndLocation should be the equal and same with other Cells location Y property (Check with Shahab)

            // Checks Row cells have all Y location //TODO: Discuss with Shahab is it true validation or not
            if (Cells.Count != 0)
            {
                var firstCellYLoc = Cells.First().Location.Y;

                foreach (var cell in Cells)
                {
                    if (cell.Location.Y != firstCellYLoc)
                        yield return new ValidationResult("All row cells should have same Y location");
                }
            }

            // Check MergedCells format
            foreach (var cellsToMerge in MergedCellsList)
            {
                if (string.IsNullOrWhiteSpace(cellsToMerge) || cellsToMerge.Contains(":") is false)
                    yield return
                        new ValidationResult("Something is not right about MergedCells format specified in Row model");

                // A2:B2 should be along with cells with locationY=2 //TODO: Confirm it with Shahab
                foreach (var c in cellsToMerge!.ToCharArray())
                {
                    if (c.ToString().IsNumber() && Cells.Count != 0 && Convert.ToInt32(c.ToString()) != Cells.First()?.Location.Y)
                    {
                        yield return new ValidationResult("In MergedCell Az:Bz the z should be the same with Row Y location");
                    }
                }
            }
        }
    }
}
