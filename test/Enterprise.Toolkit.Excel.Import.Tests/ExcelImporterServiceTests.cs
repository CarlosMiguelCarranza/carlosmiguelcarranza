using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reactive.Linq;
using System.Threading;
using System.Threading.Tasks;
using Enterprise.Toolkit.Excel.Import.Models;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using Xunit;
using Xunit.Abstractions;

using Microsoft.Extensions.Logging.Abstractions;
using Moq;

namespace Enterprise.Toolkit.Excel.Import.Tests;

public class ExcelImporterServiceTests
{
    private readonly ITestOutputHelper _output;
    private readonly ILogger<ExcelImporterService<TestModel>> _logger;

    public ExcelImporterServiceTests(ITestOutputHelper output)
    {
        _output = output;
        _logger = NullLogger<ExcelImporterService<TestModel>>.Instance;
        // Set the license context for EPPlus
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    public record TestModel
    {
        [Column(1, IsRequired = true)]
        public string Name { get; set; } = string.Empty;

        [Column(2)]
        public int Age { get; set; }
        
        [Column(3)]
        public decimal? Salary { get; set; }
    }

    private MemoryStream CreateSampleExcelStream()
    {
        var stream = new MemoryStream();
        using var package = new ExcelPackage(stream);
        var worksheet = package.Workbook.Worksheets.Add("Sheet1");

        // Header
        worksheet.Cells[1, 1].Value = "Name";
        worksheet.Cells[1, 2].Value = "Age";
        worksheet.Cells[1, 3].Value = "Salary";

        // Data Rows
        worksheet.Cells[2, 1].Value = "John Doe";    // Good
        worksheet.Cells[2, 2].Value = 30;
        worksheet.Cells[2, 3].Value = 50000.50m;

        worksheet.Cells[3, 1].Value = "Jane Smith";  // Good
        worksheet.Cells[3, 2].Value = 25;
        worksheet.Cells[3, 3].Value = 45000.75m;
        
        worksheet.Cells[4, 1].Value = null;          // Bad (Name is required)
        worksheet.Cells[4, 2].Value = 35;
        worksheet.Cells[4, 3].Value = 60000.00m;
        
        worksheet.Cells[5, 1].Value = "Peter Jones"; // Bad (Age is not a number)
        worksheet.Cells[5, 2].Value = "forty";
        worksheet.Cells[5, 3].Value = 70000.00m;
        
        worksheet.Cells[6, 1].Value = "Alice Brown"; // Good
        worksheet.Cells[6, 2].Value = 28;
        worksheet.Cells[6, 3].Value = null;

        package.Save();
        stream.Position = 0;
        return stream;
    }

    [Fact]
    public async Task ProcessExcelStream_WithGeneratedFile_ReturnsCorrectCounts()
    {
        // Arrange
        var service = new ExcelImporterService<TestModel>(logger: _logger);
        await using var testStream = CreateSampleExcelStream();

        // Act
        var (completionTask, _) = service.ProcessExcelStream(testStream, enableLogging: false);
        var result = await completionTask;

        // Assert
        Assert.Equal(5, result.TotalRowsRead);
        Assert.Equal(3, result.SuccessfulRows);
        Assert.Equal(2, result.FailedRows);
    }

    [Fact]
    public async Task ProcessExcelStream_WithInvalidRows_ReportsDetailedErrors()
    {
        // Arrange
        var service = new ExcelImporterService<TestModel>(logger: _logger);
        var allErrors = new List<RowValidationError>();
        var errorLock = new object();
        await using var testStream = CreateSampleExcelStream();

        // Act
        var (completionTask, progressStream) = service.ProcessExcelStream(testStream, enableLogging: false);
        
        progressStream
            .Where(report => report.RecentErrors != null && report.RecentErrors.Any())
            .Subscribe(report =>
            {
                lock (errorLock)
                {
                    allErrors.AddRange(report.RecentErrors!);
                }
            });
        
        await completionTask;

        // Print collected errors for debugging
        _output.WriteLine($"Collected {allErrors.Count} errors:");
        foreach (var error in allErrors.OrderBy(e => e.RowNumber))
        {
            _output.WriteLine($"- Row: {error.RowNumber}, Field: {error.FieldName}, Msg: {error.ErrorMessage}");
        }

        // Assert
        Assert.Equal(2, allErrors.Count);

        // Check for the missing required 'Name' (Row 4 in Excel, so RowNumber 4)
        Assert.Contains(allErrors, error => error.RowNumber == 4 && error.FieldName == "Name" && error.ErrorMessage.Contains("required"));
        
        // Check for the invalid 'Age' type (Row 5 in Excel, so RowNumber 5)
        Assert.Contains(allErrors, error => error.RowNumber == 5 && error.FieldName == "Age" && error.ErrorMessage.Contains("Could not convert"));
    }

    [Fact]
    public async Task ProcessExcelStream_WhenCancelled_ThrowsOperationCanceledException()
    {
        // Arrange
        var service = new ExcelImporterService<TestModel>(logger: _logger);
        var cts = new CancellationTokenSource();
        await using var testStream = CreateSampleExcelStream();

        // Act
        var (completionTask, _) = service.ProcessExcelStream(testStream, enableLogging: false, cts.Token);
        
        cts.Cancel();

        // Assert
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => completionTask);
    }

    [Fact]
    public async Task ProcessExcelStream_EnableLoggingFalse_NoLogsEmitted()
    {
        // Arrange
        var mockLogger = new Mock<ILogger<ExcelImporterService<TestModel>>>();
        var service = new ExcelImporterService<TestModel>(logger: mockLogger.Object);
        await using var testStream = CreateSampleExcelStream();

        // Act
        var (completionTask, _) = service.ProcessExcelStream(testStream, enableLogging: false);
        await completionTask;

        // Assert
        mockLogger.Verify(
            x => x.Log(
                It.IsAny<LogLevel>(),
                It.IsAny<EventId>(),
                It.IsAny<It.IsAnyType>(),
                It.IsAny<Exception>(),
                (Func<It.IsAnyType, Exception?, string>)It.IsAny<object>()),
            Times.Never());
    }

    [Fact]
    public async Task ProcessExcelStream_EnableLoggingTrue_LogsEmitted()
    {
        // Arrange
        var mockLogger = new Mock<ILogger<ExcelImporterService<TestModel>>>();
        var service = new ExcelImporterService<TestModel>(logger: mockLogger.Object);
        await using var testStream = CreateSampleExcelStream();

        // Act
        var (completionTask, _) = service.ProcessExcelStream(testStream, enableLogging: true);
        await completionTask;

        // Assert
        mockLogger.Verify(
            x => x.Log(
                LogLevel.Information,
                It.IsAny<EventId>(),
                It.Is<It.IsAnyType>((o, t) => o.ToString()!.Contains("Starting Excel import process")),
                It.IsAny<Exception>(),
                (Func<It.IsAnyType, Exception?, string>)It.IsAny<object>()),
            Times.Once());

        mockLogger.Verify(
            x => x.Log(
                LogLevel.Warning,
                It.IsAny<EventId>(),
                It.Is<It.IsAnyType>((o, t) => o.ToString()!.Contains("failed validation")),
                It.IsAny<Exception>(),
                (Func<It.IsAnyType, Exception?, string>)It.IsAny<object>()),
            Times.Exactly(2));
    }
}
