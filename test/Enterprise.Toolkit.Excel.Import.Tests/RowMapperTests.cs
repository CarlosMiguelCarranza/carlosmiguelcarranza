using System.Globalization;
using Enterprise.Toolkit.Excel.Import.Models;
using Enterprise.Toolkit.Excel.Import.Services;
using Xunit;

namespace Enterprise.Toolkit.Excel.Import.Tests;

public class RowMapperTests
{
    private record TestModel
    {
        [Column(1, IsRequired = true)]
        public string Name { get; set; } = string.Empty;

        [Column(2)]
        public int Age { get; set; }
        
        [Column(3)]
        public decimal? Salary { get; set; }
        
        [Column(4)]
        public bool IsActive { get; set; }
        
        [Column(5, IsRequired = true)]
        public DateTime HireDate { get; set; }

        [Column(6)]
        public double? Height { get; set; }
    }

    [Fact]
    public void Map_ValidRow_ReturnsSuccessAndCorrectlyTypedData()
    {
        // Arrange
        var mapper = new RowMapper<TestModel>();
        var rowData = new object[] { "John Doe", 30, 50000.50m, "true", new DateTime(2023, 1, 15), 180.3 };
        const int rowNumber = 2;

        // Act
        var result = mapper.Map(rowData, rowNumber);

        // Assert
        Assert.True(result.IsSuccess);
        Assert.NotNull(result.Result);
        var model = result.Result;
        Assert.Equal("John Doe", model.Name);
        Assert.Equal(30, model.Age);
        Assert.Equal(50000.50m, model.Salary);
        Assert.True(model.IsActive);
        Assert.Equal(new DateTime(2023, 1, 15), model.HireDate);
        Assert.Equal(180.3, model.Height);
    }
    
    [Fact]
    public void Map_RequiredFieldIsEmpty_ReturnsFailureWithCorrectError()
    {
        // Arrange
        var mapper = new RowMapper<TestModel>();
        var rowData = new object?[] { null, 30, 50000.50, true, new DateTime(2023, 1, 15), 180.3 };
        const int rowNumber = 3;
        
        // Act
        var result = mapper.Map(rowData, rowNumber);
        
        // Assert
        Assert.False(result.IsSuccess);
        Assert.Single(result.Validation.Errors);
        var error = result.Validation.Errors.First();
        Assert.Equal(rowNumber, error.RowNumber);
        Assert.Equal(nameof(TestModel.Name), error.FieldName);
        Assert.Contains("required", error.ErrorMessage);
    }
    
    [Fact]
    public void Map_WrongDataType_ReturnsFailureWithCorrectError()
    {
        // Arrange
        var mapper = new RowMapper<TestModel>();
        var rowData = new object[] { "Jane Doe", "thirty", 50000.50, true, new DateTime(2023, 1, 15), 180.3 };
        const int rowNumber = 4;
        
        // Act
        var result = mapper.Map(rowData, rowNumber);
        
        // Assert
        Assert.False(result.IsSuccess);
        Assert.Single(result.Validation.Errors);
        var error = result.Validation.Errors.First();
        Assert.Equal(rowNumber, error.RowNumber);
        Assert.Equal(nameof(TestModel.Age), error.FieldName);
        Assert.Contains("Could not convert", error.ErrorMessage);
    }
    
    [Fact]
    public void Map_NullableFieldIsEmpty_ReturnsSuccessWithNullValue()
    {
        // Arrange
        var mapper = new RowMapper<TestModel>();
        var rowData = new object?[] { "John Doe", 30, null, "true", new DateTime(2023, 1, 15), null };
        const int rowNumber = 5;

        // Act
        var result = mapper.Map(rowData, rowNumber);

        // Assert
        Assert.True(result.IsSuccess);
        Assert.NotNull(result.Result);
        Assert.Null(result.Result.Salary);
        Assert.Null(result.Result.Height);
    }

    [Fact]
    public void Map_CultureWithCommaDecimal_ParsesCorrectly()
    {
        // Arrange
        var culture = new CultureInfo("de-DE"); // German culture uses a comma for decimals
        var mapper = new RowMapper<TestModel>(culture);
        var rowData = new object[] { "Hans Schmidt", 40, "75000,25", "true", "15.01.2023", "180,3" };
        const int rowNumber = 6;
        
        // Act
        var result = mapper.Map(rowData, rowNumber);
        
        // Assert
        Assert.True(result.IsSuccess, $"Validation failed with errors: {string.Join(", ", result.Validation.Errors.Select(e => e.ErrorMessage))}");
        Assert.NotNull(result.Result);
        var model = result.Result;
        Assert.Equal(75000.25m, model.Salary);
        Assert.Equal(180.3, model.Height);
    }

    private record TestModelWithGuid
    {
        [Column(1)]
        public Guid Id { get; set; }
    }

    [Fact]
    public void Map_WithCustomGuidConverter_ParsesCorrectly()
    {
        // Arrange
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
        
        var mapper = new RowMapper<TestModelWithGuid>(customConverters: customConverters);
        var testGuid = Guid.NewGuid();
        var rowData = new object[] { testGuid.ToString() };
        const int rowNumber = 7;

        // Act
        var result = mapper.Map(rowData, rowNumber);

        // Assert
        Assert.True(result.IsSuccess);
        Assert.NotNull(result.Result);
        Assert.Equal(testGuid, result.Result.Id);
    }
}
