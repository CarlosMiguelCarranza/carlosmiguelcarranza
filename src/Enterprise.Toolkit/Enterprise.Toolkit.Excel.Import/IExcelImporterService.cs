using System;
using System.Collections.Generic;
using System.Text;

namespace Enterprise.Toolkit.Excel.Import;

/// <summary>
/// Represents the final result of an import operation.
/// </summary>
/// <param name="TotalRowsRead">The total number of rows read from the Excel file.</param>
/// <param name="SuccessfulRows">The final count of successfully processed rows.</param>
/// <param name="FailedRows">The final count of failed rows.</param>
public record ImportResult(long TotalRowsRead, long SuccessfulRows, long FailedRows);

/// <summary>
/// Defines the contract for a high-performance, observable Excel importer service.
/// </summary>
/// <typeparam name="T">The type of the entity to be created from each Excel row.</typeparam>
public interface IExcelImporterService<T> where T : class, new()
{
    /// <summary>
    /// Processes an Excel file from a stream, providing real-time progress updates.
    /// </summary>
    /// <param name="stream">The stream containing the Excel file (.xlsx).</param>
    /// <param name="enableLogging">Optional: If set to false (default), logging messages from the service will be suppressed. Set to true to enable logging.</param>
    /// <param name="cancellationToken">A token to cancel the operation.</param>
    /// <returns>
    /// A tuple containing:
    /// 1. A Task representing the entire import operation, which completes with a final <see cref="ImportResult"/>.
    /// 2. An IObservable stream of <see cref="ImportProgressReport"/> for real-time progress tracking.
    /// </returns>
    (Task<ImportResult> CompletionTask, IObservable<ImportProgressReport> ProgressStream) ProcessExcelStream(
        Stream stream,
        bool enableLogging = false,
        CancellationToken cancellationToken = default);
}
