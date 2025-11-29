# Enterprise Toolkit: Excel Importer

A high-performance, observable, and extensible library for importing data from Excel files into .NET objects.

This library is designed for server-side applications where performance, detailed error reporting, and handling of large files are critical. It uses a multi-threaded producer-consumer pattern to process Excel rows concurrently, providing significant performance gains on multi-core systems.

## Features

- **High-Performance Streaming**: Reads Excel files row-by-row without loading the entire file into memory, allowing for the processing of very large files with a minimal memory footprint.
- **Asynchronous & Observable**: Built with `async/await` and Rx.NET (`IObservable`), providing a modern, non-blocking way to handle import operations and report progress in real-time.
- **Concurrent Processing**: Uses a producer-consumer pattern to read from the Excel file on one thread and process rows on multiple consumer threads, maximizing CPU utilization.
- **Declarative Mapping**: Simple attribute-based mapping (`[Column(index)]`) to map Excel columns to properties on your C# model.
- **Detailed Error Reporting**: For each failed row, the library provides structured validation errors, including the row number, field name, and a descriptive error message.
- **DI-Friendly**: Designed to be easily integrated into modern .NET applications using Dependency Injection.
- **Extensible**:
  - **Custom Type Converters**: Allows you to register your own converters for custom data types or to override default conversion logic.
  - **Configurable Logging**: Provides control over logging, which can be enabled for detailed diagnostics or disabled for maximum performance.
- **Culture-Aware**: Correctly parses numbers and dates from different cultures by specifying a `CultureInfo` object.

**Note:** The project is currently targeting `.NET 10.0`, which is not a supported version. For production use, it is critical to change the target framework to a supported LTS version, such as `.NET 8.0`.

## Getting Started

### 1. Define Your Model

Create a C# class representing a row in your Excel file. Use the `[Column]` attribute to map properties to their corresponding 1-based column index in the Excel sheet.

```csharp
// In your project
public record SampleModel
{
    [Column(1, IsRequired = true)]
    public string Name { get; set; } = string.Empty;

    [Column(2)]
    public int Age { get; set; }

    [Column(3)]
    public decimal? Salary { get; set; }
}
```

### 2. Instantiate the Service

Create an instance of `ExcelImporterService<T>`, where `T` is your model type.

```csharp
using Enterprise.Toolkit.Excel.Import;

// ...

var importer = new ExcelImporterService<SampleModel>();
```

### 3. Process the File

Open a `Stream` to your Excel file and call `ProcessExcelStream`. The method returns a tuple containing:
- A `Task<ImportResult>` that completes when the entire operation is finished.
- An `IObservable<ImportProgressReport>` that emits progress updates.

```csharp
await using var fileStream = File.OpenRead("path/to/your/excel.xlsx");

var (completionTask, progressStream) = importer.ProcessExcelStream(fileStream);

// Subscribe to progress updates
progressStream.Subscribe(report =>
{
    // You can get row-specific errors from here
    if (report.RecentErrors != null && report.RecentErrors.Any())
    {
        foreach (var error in report.RecentErrors)
        {
            Console.WriteLine($"Error in Row {error.RowNumber}: {error.ErrorMessage}");
        }
    }
});

// Wait for the entire import to finish
var finalResult = await completionTask;

Console.WriteLine($"Import finished. Success: {finalResult.SuccessfulRows}, Failed: {finalResult.FailedRows}");
```

## Advanced Usage

### Handling Different Cultures

If your Excel file uses regional formats for numbers or dates (e.g., a comma for a decimal separator), you can provide a `CultureInfo` object.

```csharp
using System.Globalization;

var germanCulture = new CultureInfo("de-DE");
var importer = new ExcelImporterService<SampleModel>(cultureInfo: germanCulture);
```

### Custom Type Converters

You can handle custom types (like `Guid`) or override default parsing logic by providing your own converters.

```csharp
// 1. Define your model
public record ModelWithGuid
{
    [Column(1)]
    public Guid Id { get; set; }
}

// 2. Create a converter function
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

// 3. Create a RowMapper with the custom converter
var rowMapper = new RowMapper<ModelWithGuid>(customConverters: customConverters);

// Note: To use a custom RowMapper, you would currently need to adapt the service or use the RowMapper directly.
// The ExcelImporterService has been updated to accept custom converters in its RowMapper.
```

### Controlling Logging

For performance-critical scenarios, you can disable logging by passing `enableLogging: false`.

```csharp
var (completionTask, progressStream) = importer.ProcessExcelStream(
    fileStream, 
    enableLogging: false
);
```

### Dependency Injection

The service is designed for DI. In an `Microsoft.Extensions.Hosting` application (e.g., ASP.NET Core), you can register it as follows:

```csharp
// In Program.cs
builder.Services.AddScoped(typeof(IExcelImporterService<>), typeof(ExcelImporterService<>));

// In your service
public class MyService
{
    private readonly IExcelImporterService<SampleModel> _importer;

    public MyService(IExcelImporterService<SampleModel> importer)
    {
        _importer = importer;
    }
}
```

## Running the Samples

A sample console project with detailed examples is provided in the `/samples` directory. To run it, navigate to the directory and execute:

```bash
cd samples/ExcelImport.Samples
dotnet run
```
This will demonstrate a basic import, error handling, and custom type conversion.