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
        public string Path { get; set; }
        public List<Sheet> Sheets { get; set; } = new();
    }
}
