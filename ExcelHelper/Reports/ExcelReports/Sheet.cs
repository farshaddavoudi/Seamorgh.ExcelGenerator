using ExcelHelper.Reports.ExcelReports.PropertyOptions;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace ExcelHelper.Reports.ExcelReports
{
    public class Sheet
    {
        public Sheet(string name, ProtectionOptions protectionOptions)
        {
            Name = name;
            ProtectionOptions = protectionOptions;
        }

        [Required(ErrorMessage = "Sheet Name is required")]
        public string Name { get; set; }
        public List<Row> Rows { get; set; } = new();
        public List<ColumnProps> Columns { get; set; } = new();
        public List<Cell> Cells { get; set; } = new();
        public List<Table> Tables { get; set; } = new();
        public List<string> MergedCells { get; set; } = new();
        // TODO: New Property
        public WSProps WSProps { get; set; } = new();
        public bool IsLocked { get; set; } = false;
        public ProtectionOptions ProtectionOptions { get; set; }

    }
}
