using Enterprise.Toolkit.Excel.Import.Models;

namespace ExcelImport.Samples;

public record Employee
{
    [Column(1, IsRequired = true)]
    public string Name { get; set; } = string.Empty;

    [Column(2)]
    public string Department { get; set; } = string.Empty;

    [Column(3)]
    public int? EmployeeId { get; set; }

    [Column(4)]
    public DateTime HireDate { get; set; }
    
    [Column(5)]
    public decimal Salary { get; set; }
}

public record Product
{
    [Column(1, IsRequired = true)]
    public Guid ProductId { get; set; }

    [Column(2)]
    public string Name { get; set; } = string.Empty;

    [Column(3)]
    public int UnitsInStock { get; set; }
}
