using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace ExcelHelper.Reports.ExcelReports
{
    public class EasyExcelModel
    {
        [Required(ErrorMessage = "FileName is required")]
        public string FileName { get; set; }

        public List<Sheet> Sheets { get; set; } = new();

        public SheetsDefaultStyle SheetsDefaultStyles { get; set; } = new();

        public bool SheetsDefaultIsLocked { get; set; } = false;
    }
}
