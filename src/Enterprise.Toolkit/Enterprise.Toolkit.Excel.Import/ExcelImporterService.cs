using Enterprise.Toolkit.Excel.Import.Models;
using Enterprise.Toolkit.Excel.Import.Services;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reactive.Subjects;
using System.Threading;
using System.Threading.Channels;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;

namespace Enterprise.Toolkit.Excel.Import;

/// <summary>
/// Implements a high-performance, observable Excel importer service using a Producer-Consumer pattern
/// with System.Threading.Channels for data flow and Rx.NET for progress reporting.
/// </summary>
public class ExcelImporterService<T> : IExcelImporterService<T> where T : class, new()
{
    private readonly int _degreeOfParallelism;
    private readonly CultureInfo _cultureInfo;
    private readonly ILogger<ExcelImporterService<T>> _logger;

    public ExcelImporterService(
        int degreeOfParallelism = 4, 
        CultureInfo? cultureInfo = null, 
        ILogger<ExcelImporterService<T>>? logger = null)
    {
        _degreeOfParallelism = Math.Max(1, degreeOfParallelism);
        _cultureInfo = cultureInfo ?? CultureInfo.InvariantCulture;
        _logger = logger ?? NullLogger<ExcelImporterService<T>>.Instance;

        // Required for ExcelDataReader on .NET Core.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
    }

    /// <inheritdoc />
    public (Task<ImportResult> CompletionTask, IObservable<ImportProgressReport> ProgressStream) ProcessExcelStream(
        Stream stream, 
        bool enableLogging = false, // New parameter
        CancellationToken cancellationToken = default)
    {
        if (enableLogging)
        {
            _logger.LogInformation("Starting Excel import process for type {TypeName} with {Parallelism} degree of parallelism.", typeof(T).Name, _degreeOfParallelism);
        }

        var progressSubject = new ReplaySubject<ImportProgressReport>();
        var channel = Channel.CreateBounded<(object[] Data, int RowNumber)>(new BoundedChannelOptions(_degreeOfParallelism * 2)
        {
            FullMode = BoundedChannelFullMode.Wait
        });

        long totalRowsRead = 0;
        long successfulRows = 0;
        long failedRows = 0;

        var producerTask = Task.Run(async () =>
        {
            try
            {
                using var reader = ExcelReaderFactory.CreateReader(stream);
                
                // Start counter at 1 for the header row.
                var rowCounter = 1;

                // Skip header row
                if (reader.Read())
                {
                    rowCounter++;
                }
                
                while (reader.Read())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var rowData = new object[reader.FieldCount];
                    reader.GetValues(rowData);
                    
                    // Filter out completely empty rows
                    if (rowData.All(cell => cell == null || string.IsNullOrWhiteSpace(cell.ToString())))
                    {
                        if (enableLogging)
                        {
                            _logger.LogTrace("Skipping empty row at index {RowIndex}.", rowCounter);
                        }
                        rowCounter++;
                        continue;
                    }
                    
                    await channel.Writer.WriteAsync((rowData, rowCounter), cancellationToken);
                    Interlocked.Increment(ref totalRowsRead);
                    rowCounter++;
                }
            }
            catch (Exception ex)
            {
                if (enableLogging)
                {
                    _logger.LogError(ex, "A critical error occurred in the Excel reader producer task.");
                }
                progressSubject.OnError(ex);
            }
            finally
            {
                channel.Writer.Complete();
                if (enableLogging)
                {
                    _logger.LogDebug("Excel reader producer task completed.");
                }
            }
        }, cancellationToken);

        var consumerTasks = Enumerable.Range(0, _degreeOfParallelism).Select(_ => Task.Run(async () =>
        {
            // NOTE: RowMapper is not thread-safe if custom converters are not.
            // A new instance is created per consumer task to ensure thread safety.
            var mapper = new RowMapper<T>(_cultureInfo);
            await foreach (var (data, rowNum) in channel.Reader.ReadAllAsync(cancellationToken))
            {
                var mappingResult = mapper.Map(data, rowNum);
                if (mappingResult.IsSuccess)
                {
                    Interlocked.Increment(ref successfulRows);
                }
                else
                {
                    var failedCount = Interlocked.Increment(ref failedRows);
                    var errors = mappingResult.Validation.Errors;
                    if (enableLogging)
                    {
                        _logger.LogWarning("Row {RowNumber} failed validation. Total failures: {FailureCount}. Errors: {ValidationErrors}", 
                            rowNum, failedCount, string.Join(", ", errors.Select(e => e.ErrorMessage)));
                    }

                    progressSubject.OnNext(new ImportProgressReport(
                        Interlocked.Read(ref totalRowsRead),
                        Interlocked.Read(ref successfulRows),
                        failedCount,
                        RecentErrors: errors
                    ));
                }

                // Report progress periodically to avoid flooding the observer.
                if (rowNum % 100 == 0)
                {
                    progressSubject.OnNext(new ImportProgressReport(
                        Interlocked.Read(ref totalRowsRead),
                        Interlocked.Read(ref successfulRows),
                        Interlocked.Read(ref failedRows)
                    ));
                }
            }
        }, cancellationToken)).ToList();

        var completionTask = Task.Run(async () =>
        {
            await Task.WhenAll(producerTask, Task.WhenAll(consumerTasks));
            
            if (enableLogging)
            {
                _logger.LogInformation("Excel import process finished. Total rows read: {TotalRowsRead}, Successful: {SuccessfulRows}, Failed: {FailedRows}.",
                    totalRowsRead, successfulRows, failedRows);
            }

            var finalResult = new ImportResult(totalRowsRead, successfulRows, failedRows);
            progressSubject.OnNext(new ImportProgressReport(totalRowsRead, successfulRows, failedRows, IsCompleted: true));
            progressSubject.OnCompleted();
            return finalResult;
        }, cancellationToken);

        return (completionTask, progressSubject);
    }
}
