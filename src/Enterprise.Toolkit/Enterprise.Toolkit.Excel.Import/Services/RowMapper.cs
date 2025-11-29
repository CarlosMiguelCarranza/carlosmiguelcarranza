using Enterprise.Toolkit.Excel.Import.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;

namespace Enterprise.Toolkit.Excel.Import.Services;

public class RowMapper<T> where T : class, new()
{
    private readonly List<(PropertyInfo Property, ColumnAttribute Column)> _columnMappings;
    private readonly PropertyInfo? _rowKeyProperty;
    private readonly PropertyInfo? _validationProperty;

    // A dictionary-based type converter, which is much cleaner than a large if-else block.
    private static readonly Dictionary<Type, Func<object?, CultureInfo, (object?, bool)>> TypeConverters = new()
    {
        { typeof(string), (val, _) => ConvertToString(val) },
        { typeof(int), (val, culture) => ConvertToNumber<int>(val, culture, int.TryParse) },
        { typeof(int?), (val, culture) => ConvertToNullableNumber<int>(val, culture, int.TryParse) },
        { typeof(long), (val, culture) => ConvertToNumber<long>(val, culture, long.TryParse) },
        { typeof(long?), (val, culture) => ConvertToNullableNumber<long>(val, culture, long.TryParse) },
        { typeof(decimal), (val, culture) => ConvertToNumber<decimal>(val, culture, decimal.TryParse) },
        { typeof(decimal?), (val, culture) => ConvertToNullableNumber<decimal>(val, culture, decimal.TryParse) },
        { typeof(double), (val, culture) => ConvertToNumber<double>(val, culture, double.TryParse) },
        { typeof(double?), (val, culture) => ConvertToNullableNumber<double>(val, culture, double.TryParse) },
        { typeof(float), (val, culture) => ConvertToNumber<float>(val, culture, float.TryParse) },
        { typeof(float?), (val, culture) => ConvertToNullableNumber<float>(val, culture, float.TryParse) },
        { typeof(bool), (val, _) => ConvertToBool(val) },
        { typeof(bool?), (val, _) => ConvertToNullableBool(val) },
        { typeof(DateTime), (val, culture) => ConvertToDateTime(val, culture) },
        { typeof(DateTime?), (val, culture) => ConvertToNullableDateTime(val, culture) }
    };
    
    private readonly CultureInfo _cultureInfo;

    public RowMapper(CultureInfo? cultureInfo = null)
    {
        _cultureInfo = cultureInfo ?? CultureInfo.InvariantCulture;
        
        _columnMappings = typeof(T).GetProperties()
            .Where(p => p.IsDefined(typeof(ColumnAttribute), true))
            .Select(p => (Property: p, Column: p.GetCustomAttribute<ColumnAttribute>()!))
            .ToList();

        _rowKeyProperty = FindProperty<RowKeyAttribute>();
        _validationProperty = FindProperty<ValidationFieldAttribute>();
    }

    public RowMappingResult<T> Map(object?[] rowData, int rowNumber)
    {
        var result = new RowMappingResult<T>();
        var instance = new T();
        result.Result = instance;

        foreach (var (prop, attr) in _columnMappings)
        {
            // ExcelDataReader is 1-based, but our array is 0-based.
            int colIndex = attr.ColumnIndex - 1;
            object? cellValue = colIndex < rowData.Length ? rowData[colIndex] : null;

            if (cellValue == null || string.IsNullOrWhiteSpace(cellValue.ToString()))
            {
                if (attr.IsRequired)
                {
                    result.Validation.AddError(rowNumber, prop.Name, $"Field is required but the cell was empty.");
                }
                else
                {
                    prop.SetValue(instance, null);
                }
                continue;
            }

            if (TypeConverters.TryGetValue(prop.PropertyType, out var converter))
            {
                var (convertedValue, success) = converter(cellValue, _cultureInfo);
                if (success)
                {
                    prop.SetValue(instance, convertedValue);
                }
                else
                {
                    result.Validation.AddError(rowNumber, prop.Name, $"Could not convert value '{cellValue}' to the required type '{prop.PropertyType.Name}'.");
                }
            }
            else
            {
                // Fallback for unhandled types
                try
                {
                    var convertedValue = Convert.ChangeType(cellValue, prop.PropertyType, _cultureInfo);
                    prop.SetValue(instance, convertedValue);
                }
                catch (Exception)
                {
                    result.Validation.AddError(rowNumber, prop.Name, $"No converter was found for type '{prop.PropertyType.Name}'.");
                }
            }
        }
        
        _rowKeyProperty?.SetValue(instance, rowNumber);
        _validationProperty?.SetValue(instance, result.Validation);

        return result;
    }

    private PropertyInfo? FindProperty<TAttr>() where TAttr : Attribute
    {
        return typeof(T).GetProperties()
            .FirstOrDefault(p => p.IsDefined(typeof(TAttr), true));
    }
    
    #region Type Conversion Helpers

    private delegate bool TryParseHandler<T>(string s, NumberStyles style, IFormatProvider provider, out T result);

    private static (object?, bool) ConvertToString(object? value) => (value?.ToString(), true);

    private static (object?, bool) ConvertToNumber<T>(object? value, CultureInfo culture, TryParseHandler<T> handler) where T : struct
    {
        if (value is T number) return (number, true);
        if (handler(value?.ToString(), NumberStyles.Any, culture, out var result))
        {
            return (result, true);
        }
        return (default(T), false);
    }

    private static (object?, bool) ConvertToNullableNumber<T>(object? value, CultureInfo culture, TryParseHandler<T> handler) where T : struct
    {
        if (value == null || value is DBNull) return (null, true);
        if (value is T number) return (number, true);
        var strValue = value.ToString();
        if (string.IsNullOrWhiteSpace(strValue)) return (null, true);
        if (handler(strValue, NumberStyles.Any, culture, out var result))
        {
            return (result, true);
        }
        return (null, false);
    }
    
    private static (object?, bool) ConvertToBool(object? value)
    {
        if (value is bool b) return (b, true);
        var s = value?.ToString()?.Trim().ToLowerInvariant();
        switch (s)
        {
            case "true":
            case "yes":
            case "1":
                return (true, true);
            case "false":
            case "no":
            case "0":
                return (false, true);
            default:
                return (default(bool), false);
        }
    }

    private static (object?, bool) ConvertToNullableBool(object? value)
    {
        if (value == null || value is DBNull) return (null, true);
        var s = value?.ToString()?.Trim().ToLowerInvariant();
        if (string.IsNullOrWhiteSpace(s)) return (null, true);
        
        var (result, success) = ConvertToBool(value);
        return (success ? result : null, success);
    }
    
    private static (object?, bool) ConvertToDateTime(object? value, CultureInfo culture)
    {
        if (value is DateTime dt) return (dt, true);
        if (value is double oaDate) return (DateTime.FromOADate(oaDate), true); // Excel stores dates as numbers
        if (DateTime.TryParse(value?.ToString(), culture, DateTimeStyles.None, out var result))
        {
            return (result, true);
        }
        return (default(DateTime), false);
    }

    private static (object?, bool) ConvertToNullableDateTime(object? value, CultureInfo culture)
    {
        if (value == null || value is DBNull) return (null, true);
        var strValue = value.ToString();
        if (string.IsNullOrWhiteSpace(strValue)) return (null, true);
        
        var (result, success) = ConvertToDateTime(value, culture);
        return (success ? result : null, success);
    }

    #endregion
}
