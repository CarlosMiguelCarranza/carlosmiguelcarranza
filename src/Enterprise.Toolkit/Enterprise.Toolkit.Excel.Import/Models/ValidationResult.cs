using System.Collections.Generic;
using System.Linq;

namespace Enterprise.Toolkit.Excel.Import.Models;

/// <summary>
/// A simple class to hold validation errors for a processed row.
/// </summary>
public class ValidationResult
{
    private readonly List<RowValidationError> _errors = new();

    /// <summary>
    /// Indicates whether the row has any validation errors.
    /// </summary>
    public bool IsValid => !_errors.Any();

    /// <summary>
    /// A collection of structured error messages for the row.
    /// </summary>
    public IReadOnlyCollection<RowValidationError> Errors => _errors;

    /// <summary>
    /// Adds a new structured error.
    /// </summary>
    /// <param name="rowNumber">The row number where the error occurred.</param>
    /// <param name="fieldName">The name of the field that failed validation.</param>
    /// <param name="errorMessage">The validation error message.</param>
    public void AddError(int rowNumber, string fieldName, string errorMessage)
    {
        _errors.Add(new RowValidationError(rowNumber, fieldName, errorMessage));
    }
}
