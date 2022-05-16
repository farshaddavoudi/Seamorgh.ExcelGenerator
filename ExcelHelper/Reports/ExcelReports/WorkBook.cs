using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace ExcelHelper.Reports.ExcelReports
{
    public class WorkBook
    {
        [Required(ErrorMessage = "FileName is required")]
        public string FileName { get; set; }
        public string Path { get; set; } //TODO: Remove this property
        public List<Sheet> Sheets { get; set; } = new();

        // TODO: New Property 
        public SheetsDefaultStyles SheetsDefaultStyles { get; set; } = new();
    }
}
