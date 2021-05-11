using System;

namespace ExcelHelper.Reports.ExcelReports
{
    public class Cell
    {
        public Cell(Location location)
        {
            Location = location;
        }
        public string Name { get; set; } //TODO: Add Name property somehow as column (cell) identifier
        internal Type Type { get; set; }
        public object Value { get; set; }
        public Location Location { get; set; }
        public bool Wordwrap { get; set; }
        public TextAlign? TextAlign { get; set; }
        public Category Category { get; set; } = Category.General;
        public bool Visible { get; set; } = true;
        // TODO: Add Comments to cells
        public bool IsLock { get; set; } = false;
    }

    public enum Category
    {
        General,
        Number,
        Currency,
        MiladiDate,
        // TODO: Discussion with Shahab about removing it because it should be set in business, not in nuget
        SolarHijriDate, //Will convert Miladi date to Solar Hijri e.g. 1400/02/02 
        Text,
        Percentage
    }
}
