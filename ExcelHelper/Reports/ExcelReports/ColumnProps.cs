using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace ExcelHelper.Reports.ExcelReports
{
    public class ColumnProps : IValidatableObject
    {
        public int ColumnNumber { get; set; }
        public ColumnWidth Width { get; set; } = null; //If not specified, default would be considered
        public bool IsHidden { get; set; } = false;
        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
        {
            if (Width is not null)
            {
                if (Width.CalculateType == ColumnWidthCalculateType.ExplicitValue && Width.Value is null)
                    yield return new ValidationResult(
                        "Column width value should be specified when CalculateType is set to explicit value",
                        new List<string> { nameof(Width.Value) });
            }
        }
    }

    public enum ColumnWidthCalculateType
    {
        ExplicitValue,
        AdjustToContents
    }

    public class ColumnWidth
    {
        public ColumnWidthCalculateType CalculateType { get; set; } = ColumnWidthCalculateType.ExplicitValue;
        public double? Value { get; set; }
    }
}