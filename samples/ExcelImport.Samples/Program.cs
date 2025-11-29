using System.Globalization;
using System.Reactive.Linq;
using Enterprise.Toolkit.Excel.Import;
using Enterprise.Toolkit.Excel.Import.Models;
using Enterprise.Toolkit.Excel.Import.Services;
using ExcelImport.Samples;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;

// --- Setup Generic Host for DI, Logging, etc. ---
var host = Host.CreateDefaultBuilder(args)
    .ConfigureServices((context, services) =>
    {
        // Register the importer service
        services.AddScoped(typeof(IExcelImporterService<>), typeof(ExcelImporterService<>));
        
        // Add the main application service
        services.AddHostedService<SampleRunner>();
    })
    .Build();

await host.RunAsync();

// --- Main Application Logic ---
public class SampleRunner : IHostedService
{
    private readonly IServiceProvider _serviceProvider;
    private readonly ILogger<SampleRunner> _logger;

    public SampleRunner(IServiceProvider serviceProvider, ILogger<SampleRunner> logger)
    {
        _serviceProvider = serviceProvider;
        _logger = logger;
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    public async Task StartAsync(CancellationToken cancellationToken)
    {
        _logger.LogInformation("--- Starting Excel Importer Samples ---");

        await RunBasicImportAsync();
        await RunCultureAndErrorHandlingAsync();
        await RunCustomConverterAsync();

        _logger.LogInformation("--- Samples finished. Application will now exit. ---");
        // Trigger application shutdown
        using var scope = _serviceProvider.CreateScope();
        var lifetime = scope.ServiceProvider.GetRequiredService<IHostApplicationLifetime>();
        lifetime.StopApplication();
    }

    public Task StopAsync(CancellationToken cancellationToken) => Task.CompletedTask;
    
    // --- Example 1: Basic Import ---
    private async Task RunBasicImportAsync()
    {
        _logger.LogInformation("\n--- Running Example 1: Basic Import ---");
        
        await using var stream = CreateBasicExcelFile();
        using var scope = _serviceProvider.CreateScope();
        var importer = scope.ServiceProvider.GetRequiredService<IExcelImporterService<Employee>>();

        var (completionTask, progressStream) = importer.ProcessExcelStream(stream, enableLogging: true);

        // We can optionally listen to progress, but for this simple case, we'll just await the final result.
        var result = await completionTask;
        
        _logger.LogInformation("Basic import complete. Success: {Success}, Failed: {Failed}", result.SuccessfulRows, result.FailedRows);
    }
    
    // --- Example 2: Culture and Error Handling ---
    private async Task RunCultureAndErrorHandlingAsync()
    {
        _logger.LogInformation("\n--- Running Example 2: Culture and Error Handling ---");
        
        await using var stream = CreateRegionalExcelFile();
        using var scope = _serviceProvider.CreateScope();

        // When creating the service, we can specify the culture for parsing.
        var germanCulture = new CultureInfo("de-DE");
        var importer = new ExcelImporterService<Employee>(cultureInfo: germanCulture, logger: _logger as ILogger<ExcelImporterService<Employee>>);
        
        var (completionTask, progressStream) = importer.ProcessExcelStream(stream, enableLogging: true);

        progressStream
            .Where(report => report.RecentErrors != null && report.RecentErrors.Any())
            .Subscribe(report =>
            {
                foreach(var error in report.RecentErrors!)
                {
                    _logger.LogWarning("Validation Error. Row: {Row}, Field: {Field}, Message: {Message}", 
                        error.RowNumber, error.FieldName, error.ErrorMessage);
                }
            });

        var result = await completionTask;
        _logger.LogInformation("Regional import complete. Success: {Success}, Failed: {Failed}", result.SuccessfulRows, result.FailedRows);
    }

    // --- Example 3: Custom Converters ---
    private async Task RunCustomConverterAsync()
    {
        _logger.LogInformation("\n--- Running Example 3: Custom GUID Converter ---");

        await using var stream = CreateProductExcelFile();
        using var scope = _serviceProvider.CreateScope();
        
        // 1. Define a custom converter function.
        var customConverters = new Dictionary<Type, Func<object?, CultureInfo, (object?, bool)>> 
        {
            {
                typeof(Guid), (val, _) =>
                {
                    if (val is string s && Guid.TryParse(s, out var guid))
                    {
                        return (guid, true);
                    }
                    return (default(Guid), false);
                }
            }
        };

        // 2. Create a RowMapper with the custom converters.
        var rowMapper = new RowMapper<Product>(customConverters: customConverters);
        
        // 3. To use the custom RowMapper, we can't use DI for the service in this example. 
        //    A more advanced DI setup could handle this. We'll instantiate it directly.
        //    (Note: The service itself could be extended to accept a RowMapper factory).
        var logger = scope.ServiceProvider.GetRequiredService<ILogger<ExcelImporterService<Product>>>();
        var importer = new ExcelImporterService<Product>(logger: logger);
        
        // This is a limitation we'll work around for the sample.
        // We'll manually adapt the service logic slightly here to show the RowMapper usage.
        // The correct way would be to inject the mapper into the service.
        _logger.LogInformation("For this example, we are using the RowMapper directly to show custom conversion.");
        var reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);
        reader.Read(); // skip header
        
        int successful = 0;
        int failed = 0;
        int rowNum = 2;
        while(reader.Read())
        {
            var rowData = new object[reader.FieldCount];
            reader.GetValues(rowData);
            var result = rowMapper.Map(rowData, rowNum);
            if (result.IsSuccess)
            {
                _logger.LogInformation("Successfully mapped Product: {ProductId}", result.Result!.ProductId);
                successful++;
            }
            else
            {
                failed++;
            }
            rowNum++;
        }
        _logger.LogInformation("Custom converter import complete. Success: {Success}, Failed: {Failed}", successful, failed);
    }

    // --- Helper methods to create sample Excel files ---
    private MemoryStream CreateBasicExcelFile()
    {
        var stream = new MemoryStream();
        using var package = new ExcelPackage(stream);
        var ws = package.Workbook.Worksheets.Add("Employees");
        ws.Cells[1, 1].Value = "Name";
        ws.Cells[1, 2].Value = "Department";
        ws.Cells[1, 3].Value = "Employee ID";
        ws.Cells[1, 4].Value = "Hire Date";
        ws.Cells[1, 5].Value = "Salary";

        ws.Cells[2, 1].Value = "John Doe";
        ws.Cells[2, 2].Value = "Engineering";
        ws.Cells[2, 3].Value = 101;
        ws.Cells[2, 4].Value = new DateTime(2022, 1, 15);
        ws.Cells[2, 4].Style.Numberformat.Format = "yyyy-mm-dd";
        ws.Cells[2, 5].Value = 75000.50m;
        
        package.Save();
        stream.Position = 0;
        return stream;
    }
    
    private MemoryStream CreateRegionalExcelFile()
    {
        var stream = new MemoryStream();
        using var package = new ExcelPackage(stream);
        var ws = package.Workbook.Worksheets.Add("Mitarbeiter");
        ws.Cells[1, 1].Value = "Name";
        ws.Cells[1, 2].Value = "Abteilung";
        ws.Cells[1, 3].Value = "Mitarbeiter-ID";
        ws.Cells[1, 4].Value = "Einstellungsdatum";
        ws.Cells[1, 5].Value = "Gehalt";

        ws.Cells[2, 1].Value = "Hans Schmidt";
        ws.Cells[2, 2].Value = "Technik";
        ws.Cells[2, 3].Value = 102;
        ws.Cells[2, 4].Value = "15.01.2022"; // German date format
        ws.Cells[2, 5].Value = "85000,75";   // German decimal format
        
        ws.Cells[3, 1].Value = null; // Bad data: Name is required
        ws.Cells[3, 2].Value = "Marketing";
        ws.Cells[3, 3].Value = 103;
        ws.Cells[3, 4].Value = "01.03.2023";
        ws.Cells[3, 5].Value = "62000,00";

        package.Save();
        stream.Position = 0;
        return stream;
    }
    
    private MemoryStream CreateProductExcelFile()
    {
        var stream = new MemoryStream();
        using var package = new ExcelPackage(stream);
        var ws = package.Workbook.Worksheets.Add("Products");
        ws.Cells[1, 1].Value = "Product ID";
        ws.Cells[1, 2].Value = "Name";
        ws.Cells[1, 3].Value = "Stock";

        ws.Cells[2, 1].Value = Guid.NewGuid().ToString();
        ws.Cells[2, 2].Value = "Widget";
        ws.Cells[2, 3].Value = 150;
        
        ws.Cells[3, 1].Value = "not-a-guid"; // Bad data
        ws.Cells[3, 2].Value = "Gizmo";
        ws.Cells[3, 3].Value = 200;

        package.Save();
        stream.Position = 0;
        return stream;
    }
}
