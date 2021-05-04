using System;
using System.Data;

namespace ExcelHelper.Reports.ExcelReports
{
    public class Cell
    {
        public Cell(Location location)
        {
            Location = location;
        }
        internal DataColumn Data { get; set; }
        public string Name { get; set; } //TODO: Add Name property somehow as column (cell) identifier
        internal Type Type { get; set; }
        public object Value { get; set; }
        public Location Location { get; set; }
        public bool Wordwrap { get; set; }
        public TextAlign Align { get; set; } = TextAlign.Rtl;
        public Category Category { get; set; } = Category.General; //TODO: I didn't understand this property. What is going to do?
        public bool Visible { get; set; } = true; //Resolved TODO: Move this to "FarshadColumnWidth" class
        public bool AutoFill { get; set; } //TODO: What is AutoFill property?
    }

    public enum TextAlign
    {
        Rtl = 0,
        Ltr = 1,
        Center = 2
    }

    public enum Category
    {
        General,
        Boolean,
        Number,
        Currency,
        Date,
        Time,
        Percentage,
        Text,
        Custom
    }
}
