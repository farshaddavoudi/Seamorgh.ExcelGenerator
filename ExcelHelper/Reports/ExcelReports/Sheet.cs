using System.Collections.Generic;

namespace ExcelHelper.Reports.ExcelReports
{
    public class Sheet
    {
        public Sheet(string name)
        {
            Name = name;
        }

        public string Name { get; set; }
        public List<Row> Rows { get; set; } = new();
        public ColumnsStyle ColumnsStyle { get; set; } = new();
        public List<Cell> Cells { get; set; } = new();
        public List<Table> Tables { get; set; } = new();
        public List<string> MergedCells { get; set; } = new();
    }
}
