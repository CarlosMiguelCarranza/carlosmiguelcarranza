using System;
using System.Collections.Generic;
using System.Linq;
using Enterprise.Toolkit.Excel.Import.Models;

namespace Enterprise.Toolkit.Excel.Import;

/// <summary>
/// Represents a snapshot of the Excel import process at a specific moment in time.
/// This record is immutable.
/// </summary>
/// <param name="TotalRowsRead">The total number of rows read from the Excel file so far.</param>
/// <param name="SuccessfulRows">The count of rows that have been successfully validated and processed.</param>
/// <param name="FailedRows">The count of rows that failed validation or processing.</param>
/// <param name="IsCompleted">Indicates if the entire import process has finished.</param>
/// <param name="RecentErrors">A collection of errors that have occurred in the latest batch of processing.</param>
/// <param name="TotalRowsInFile">The total number of rows in the file, if known. Used to calculate percentage.</param>
public record ImportProgressReport(
    long TotalRowsRead,
    long SuccessfulRows,
    long FailedRows,
    bool IsCompleted = false,
    IReadOnlyCollection<RowValidationError>? RecentErrors = null,
    long? TotalRowsInFile = null)
{
    /// <summary>
    /// Calculates the estimated progress percentage. Returns null if the total number of rows is unknown.
    /// </summary>
    public double? ProgressPercentage => TotalRowsInFile.HasValue && TotalRowsInFile > 0
        ? (double)TotalRowsRead / TotalRowsInFile.Value * 100.0
        : null;
}
