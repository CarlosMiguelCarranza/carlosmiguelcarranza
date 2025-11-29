namespace Enterprise.Toolkit.Excel.Import.Models;

/// <summary>
/// Marks a property on the import model to store the original Excel row number.
/// The property should be of type int or long.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
public class RowKeyAttribute : Attribute
{
}
