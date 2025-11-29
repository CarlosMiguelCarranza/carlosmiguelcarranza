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

namespace Enterprise.Toolkit.Excel.Import;

/// <summary>
/// Implements a high-performance, observable Excel importer service using a Producer-Consumer pattern
/// with System.Threading.Channels for data flow and Rx.NET for progress reporting.
/// </summary>
public class ExcelImporterService<T> : IExcelImporterService<T> where T : class, new()
{
    private readonly int _degreeOfParallelism;
    private readonly CultureInfo _cultureInfo;

    public ExcelImporterService(int degreeOfParallelism = 4, CultureInfo? cultureInfo = null)
    {
        _degreeOfParallelism = Math.Max(1, degreeOfParallelism);
        _cultureInfo = cultureInfo ?? CultureInfo.InvariantCulture;
        
        // Required for ExcelDataReader on .NET Core.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
    }

    /// <inheritdoc />
    public (Task<ImportResult> CompletionTask, IObservable<ImportProgressReport> ProgressStream) ProcessExcelStream(Stream stream, CancellationToken cancellationToken = default)
    {
        var progressSubject = new Subject<ImportProgressReport>();
        var channel = Channel.CreateBounded<(object?[] Data, int RowNumber)>(new BoundedChannelOptions(_degreeOfParallelism * 2)
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
                // Skip header row
                reader.Read(); 
                
                while (reader.Read())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var rowData = new object?[reader.FieldCount];
                    reader.GetValues(rowData);
                    
                    // Filter out completely empty rows
                    if (rowData.All(cell => cell == null || string.IsNullOrWhiteSpace(cell.ToString())))
                    {
                        continue;
                    }
                    
                    await channel.Writer.WriteAsync((rowData, reader.RowCount), cancellationToken);
                    Interlocked.Increment(ref totalRowsRead);
                }
            }
            catch (Exception ex)
            {
                progressSubject.OnError(ex);
            }
            finally
            {
                channel.Writer.Complete();
            }
        }, cancellationToken);

        var consumerTasks = Enumerable.Range(0, _degreeOfParallelism).Select(_ => Task.Run(async () =>
        {
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
                    Interlocked.Increment(ref failedRows);
                    progressSubject.OnNext(new ImportProgressReport(
                        Interlocked.Read(ref totalRowsRead),
                        Interlocked.Read(ref successfulRows),
                        Interlocked.Read(ref failedRows),
                        RecentErrors: mappingResult.Validation.Errors
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
            var finalResult = new ImportResult(totalRowsRead, successfulRows, failedRows);
            progressSubject.OnNext(new ImportProgressReport(totalRowsRead, successfulRows, failedRows, IsCompleted: true));
            progressSubject.OnCompleted();
            return finalResult;
        }, cancellationToken);

        return (completionTask, progressSubject);
    }
}
