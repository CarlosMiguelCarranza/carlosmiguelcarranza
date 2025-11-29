namespace Enterprise.Toolkit.Excel.Import.Models;

/// <summary>
/// A wrapper for the result of mapping a single row.
/// </summary>
/// <typeparam name="T">The type of the resulting object.</typeparam>
public class RowMappingResult<T> where T : class, new()
{
    /// <summary>
    /// The resulting object from the mapping. Null if mapping failed critically.
    /// </summary>
    public T? Result { get; set; }

    /// <summary>
    /// The validation result containing any errors found during mapping.
    /// </summary>
    public ValidationResult Validation { get; } = new();

    /// <summary>
    /// Indicates if the mapping was successful and the data is valid.
    /// </summary>
    public bool IsSuccess => Validation.IsValid;
}
