using System;
using System.Data;

namespace ExcelHelper.Reports.ExcelReports
{
    public class Column //TODO: Change this name to Cell
    {
        public Column(Location location)
        {
            Location = location;
        }
        public DataColumn Data { get; set; } //Resolved TODO: What is usages of this property when we have Value property?
        public string Name { get; set; } //TODO: What is Name property for a Cell in Excel?
        public Type Type { get; set; } //Resolved TODO: What is this property for when we have Category property?
        public object Value { get; set; }
        public Location Location { get; set; }
        public bool Wordwrap { get; set; }
        public TextAlign Align { get; set; } = TextAlign.Rtl;
        public int Width { get; set; } = 20; //TODO: Move this to "Sheet" level (with AdjustToContent feature)
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
