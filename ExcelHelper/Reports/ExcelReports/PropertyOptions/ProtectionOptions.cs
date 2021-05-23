using System.ComponentModel;

namespace ExcelHelper.Reports.ExcelReports.PropertyOptions
{
    public class ProtectionOptions
    {
        public string Password { get; set; }
        [DefaultValue(true)]
        public bool Selectlockedcells { get; set; }
        public bool Selectunlockedcells { get; set; }
        public bool Formatcells { get; set; }
        public bool Formatcolumns { get; set; }
        public bool Formatrows { get; set; }
        public bool Insertcolumns { get; set; }
        public bool Insertrows { get; set; }
        public bool Inserthyperlinks { get; set; }
        public bool Deletecolumns { get; set; }
        public bool Deleterows { get; set; }
        public bool Sort { get; set; }
        public bool UseAutoFilter { get; set; }
        public bool UsePivotTablereports { get; set; }
        public bool Editobjects { get; set; }
        public bool Editscenarios { get; set; }
    }
}
