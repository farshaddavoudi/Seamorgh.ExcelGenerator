using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace ExcelHelper.Reports.ExcelReports
{
    public class ColumnProps : IValidatableObject
    {
        [Required(ErrorMessage = "ColumnNo is required")]
        public int ColumnNo { get; set; }
        public ColumnWidth Width { get; set; } = null; //If not specified, default would be considered
        public TextAlign TextAlign { get; set; } = TextAlign.Right; //Default RTL direction
        public bool IsHidden { get; set; } = false;
        // TODO: Add MergedCells for Columns property
        public bool AutoFit { get; set; } = false;

        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
        {
            if (ColumnNo == default)
                yield return new ValidationResult("ColumnNo is required", new List<string> { nameof(ColumnNo) });

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
        public ColumnWidth()
        {

        }

        public ColumnWidth(double width)
        {
            Value = width;
        }

        public ColumnWidthCalculateType CalculateType { get; set; } = ColumnWidthCalculateType.ExplicitValue;

        public double? Value { get; set; }
    }
}