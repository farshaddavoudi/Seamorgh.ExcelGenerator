using System.Collections.Generic;

namespace ExcelHelper.Reports.ExcelReports
{
    public class WorkBook
    {
        public WorkBook(string fileName, string path)
        {
            FileName = fileName;
            Path = path;
            Sheets = new();
        }

        public string FileName { get; set; }
        public string Path { get; set; }
        public List<Sheet> Sheets { get; set; }
    }
}
