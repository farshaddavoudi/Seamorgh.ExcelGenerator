using ExcelHelper.Reports.ExcelReports.PropertyOptions;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace ExcelHelper.Reports.ExcelReports
{
    public class Sheet
    {
        public Sheet(string name)
        {
            Name = name;
        }

        [Required(ErrorMessage = "Sheet Name is required")]
        public string Name { get; set; }
        public List<Row> SheetRows { get; set; } = new();
        public List<ColumnStyle> ColumnsStyle { get; set; } = new();
        public List<Cell> SheetCells { get; set; } = new();
        public List<Table> SheetTables { get; set; } = new();
        public List<string> MergedCells { get; set; } = new();
        // TODO: New Property
        public SheetStyle SheetStyle { get; set; } = new();
        public bool? IsSheetLocked { get; set; }
        public ProtectionLevels SheetProtectionLevels { get; set; } = new();
    }
}
