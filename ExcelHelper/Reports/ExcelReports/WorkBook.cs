using System.Collections.Generic;

namespace ExcelHelper.Reports.ExcelReports
{
    public class WorkBook
    {
        public WorkBook(string fileName)
        {
            FileName = fileName;
        }

        public string FileName { get; set; }
        public string Path { get; set; } //TODO: Remove this property
        public List<Sheet> Sheets { get; set; } = new();

        // TODO: New Property 
        public WBProps WBProps { get; set; } = new();
    }
}
