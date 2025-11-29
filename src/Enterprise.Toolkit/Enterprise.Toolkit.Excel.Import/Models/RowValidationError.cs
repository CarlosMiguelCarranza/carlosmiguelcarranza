namespace Enterprise.Toolkit.Excel.Import.Models;

/// <summary>
/// Represents a specific validation error for a given row and field.
/// </summary>
/// <param name="RowNumber">The row number in the Excel sheet where the error occurred.</param>
/// <param name="FieldName">The name of the property/field that failed validation.</param>
/// <param name="ErrorMessage">The specific error message.</param>
public record RowValidationError(int RowNumber, string FieldName, string ErrorMessage);
