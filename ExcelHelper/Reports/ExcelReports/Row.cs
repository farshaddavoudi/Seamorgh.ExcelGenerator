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
            AllBorder = new Border(LineStyle.None, Color.Black);
            OutsideBorder = new Border(LineStyle.None, Color.Black);
            MergedCellsList = new();
        }
        public CellLocation StartCellLocation 
        {
            get
            {
                return Cells.FirstOrDefault().CellLocation ;
            }
        }
        public CellLocation EndCellLocation 
        {
            get
            {
                return Cells.LastOrDefault().CellLocation;
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

        public CellLocation NextHorizontalCellLocation
        {
            get
            {
                var y = EndCellLocation.Y - (EndCellLocation.Y - StartCellLocation.Y);
                return new CellLocation(EndCellLocation.X + 1, y);

            }
        }

        public Cell AddCell()
        {
            Cell cell = new(NextHorizontalCellLocation);
            Cells.Add(cell);
            return cell;
        }

        public CellLocation NextVerticalCellLocation
        {
            get
            {
                var x = EndCellLocation.X - (EndCellLocation.X - StartCellLocation.X); //TODO: ?? (x-(x-y) => answer always is y)
                return new CellLocation(x, EndCellLocation.Y + 1);

            }
        }

        public Cell GetCell(int X)
        {
            return Cells.FirstOrDefault(x => x.CellLocation.X == X);
        }

        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
        {
            // Check if StartLocation exist, EndLocation should exist too ////TODO: Discuss with Shahab is it true validation or not
            if (StartCellLocation is not null && EndCellLocation is null ||
                StartCellLocation is null && EndCellLocation is not null)
                yield return new ValidationResult(
                    "Both StartLocation and EndLocation of Row should be null or not null simultaneously");

            // TODO: Check Y of StartLocation and EndLocation should be the equal and same with other Cells location Y property (Check with Shahab)

            // Checks Row cells have all Y location //TODO: Discuss with Shahab is it true validation or not
            if (Cells.Count != 0)
            {
                var firstCellYLoc = Cells.First().CellLocation.Y;

                foreach (var cell in Cells)
                {
                    if (cell.CellLocation.Y != firstCellYLoc)
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
                    if (c.ToString().IsNumber() && Cells.Count != 0 && Convert.ToInt32(c.ToString()) != Cells.First()?.CellLocation.Y)
                    {
                        yield return new ValidationResult("In MergedCell Az:Bz the z should be the same with Row Y location");
                    }
                }
            }
        }
    }
}
