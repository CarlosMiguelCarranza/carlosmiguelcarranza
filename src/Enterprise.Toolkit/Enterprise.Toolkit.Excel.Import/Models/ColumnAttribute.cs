namespace Enterprise.Toolkit.Excel.Import.Models;

/// <summary>
/// Specifies which column in the Excel sheet maps to this property.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
public class ColumnAttribute : Attribute
{
    /// <summary>
    /// The 1-based index of the column in the Excel sheet.
    /// </summary>
    public int ColumnIndex { get; }

    /// <summary>
    /// Indicates if this column is required. If true, the import will fail for rows where this column is empty.
    /// </summary>
    public bool IsRequired { get; set; }

    public ColumnAttribute(int columnIndex)
    {
        if (columnIndex <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(columnIndex), "Column index must be a positive integer.");
        }
        ColumnIndex = columnIndex;
    }
}
