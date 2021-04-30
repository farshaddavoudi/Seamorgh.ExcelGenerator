using System.Collections.Generic;

namespace ExcelHelper.Reports.ExcelReports
{
    public class Sheet
    {
        public Sheet(string caption, string name)
        {
            Name = name;
            Caption = caption;
        }

        public string Name { get; set; }
        public string Caption { get; set; }
        public List<Row> Rows { get; set; } = new();
        public List<Column> Columns { get; set; } = new();
        public List<Table> Tables { get; set; } = new();
        public List<string> MergedCells { get; set; } = new();
    }
}
