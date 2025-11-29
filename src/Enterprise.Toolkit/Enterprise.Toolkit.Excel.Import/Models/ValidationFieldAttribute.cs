namespace Enterprise.Toolkit.Excel.Import.Models;

/// <summary>
/// Marks a property to store the validation results for the row.
/// The property should be of a type that can hold a collection of validation error messages.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
public class ValidationFieldAttribute : Attribute
{
}
