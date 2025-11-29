using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reactive.Linq;
using System.Reactive.Subjects;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace cmc.ExcelUtilities
{
    public class ExcelReaderBase<TSchema> where TSchema : ExcelSheetSchemaBase
    {
        private readonly ExcelReaderOptions _settings;
        private readonly IDictionary<int, TSchema> _data = new Dictionary<int, TSchema>();
        private readonly Subject<TSchema> _dataSubject = new Subject<TSchema>();

        public IObservable<TSchema> DataStream => _dataSubject.AsObservable();

        public ExcelReaderBase(ExcelReaderOptions options)
        {
            _settings = options;
        }

        public Task<IDictionary<int, TSchema>> GetData()
        {
            return Task.FromResult(_data);
        }

        public async Task ReadAsync(string filePath, string? sheetName = null)
        {
            using var stream = System.IO.File.OpenRead(filePath);
            await ReadAsync(stream, sheetName);
        }

        public async Task ReadAsync(Stream stream, string? sheetName = null)
        {
            _data.Clear();

            var package = new ExcelPackage(stream);

            try
            {
                var worksheet = sheetName == null
                    ? package.Workbook.Worksheets.First()
                    : package.Workbook.Worksheets[sheetName];

                if (worksheet == null)
                    throw new ArgumentException($"Sheet '{sheetName}' not found");

                int startRow = _settings.HasHeaderRow
                    ? _settings.StartRowData ?? 2
                    : 1;

                int endRow = _settings.EndRowData ?? worksheet.Dimension.Rows;

                await ProcessRowsAsync(worksheet, startRow, endRow);

            }
            catch (ExcelReaderException ex)
            {
                throw new ExcelReaderException("", ex);
            }
            finally
            {
                package.Dispose();
                _dataSubject.OnCompleted();                
            }
        }


        private Task ProcessRowsAsync(ExcelWorksheet worksheet, int startRow, int endRow)
        {
            Type schemaType = typeof(TSchema);

            // Pre-mapeo de propiedades
            var propertyMappings = schemaType.GetProperties()
                .Select(prop => (
                    Property: prop,
                    Attribute: prop.GetCustomAttribute<ExcelColumnAttribute>())
                )
                .Where(x => x.Attribute != null)
                .Select(x => (
                    x.Property,
                    ColumnNumber: ColumnLetterToNumber(x.Attribute.ColumnName),
                    PropertyType: Nullable.GetUnderlyingType(x.Property.PropertyType) ?? x.Property.PropertyType,
                    IsNullable: x.Property.PropertyType.IsGenericType &&
                               x.Property.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>)
                ))
                .ToList();

            // Procesar filas
            for (int row = startRow; row <= endRow; row++)
            {
                // Crear nueva instancia
                var schemaObject = Activator.CreateInstance<TSchema>();

                foreach (var mapping in propertyMappings)
                {
                    try
                    {
                        // Obtener valor de la celda
                        object cellValue = worksheet.GetValue(row, mapping.ColumnNumber);

                        // Convertir el valor
                        object? parsedValue = cellValue == null
                            ? (mapping.IsNullable ? null : throw new Exception(""))
                            : ConvertValue(cellValue, mapping.PropertyType);

                        // Asignar valor a la propiedad
                        mapping.Property.SetValue(schemaObject, parsedValue);
                        schemaObject.Key = row;
                    }
                    catch (Exception ex)
                    {
                        schemaObject.ValidationResult.Add(
                            ColumnNumberToLetter(mapping.ColumnNumber),
                            row,
                            $"Error en fila {row}, columna {mapping.ColumnNumber} ({mapping.Property.Name}): {ex.Message}");
                    }
                }
                _data.Add(row, schemaObject);
                _dataSubject.OnNext(schemaObject);
            }

            return Task.CompletedTask;
        }

        // Métodos auxiliares
        private static object ConvertValue(object value, Type targetType)
        {
            if (value == null) return null;

            // Manejo especial para fechas de Excel
            if (targetType == typeof(DateTime) && value is double excelDate)
            {
                return DateTime.FromOADate(excelDate);
            }

            return Convert.ChangeType(value, targetType);
        }

        public static int ColumnLetterToNumber(string columnLetters)
        {
            if (string.IsNullOrEmpty(columnLetters))
                throw new ArgumentException("La cadena no puede estar vacía");

            columnLetters = columnLetters.ToUpperInvariant();
            int result = 0;

            foreach (char c in columnLetters)
            {
                if (c < 'A' || c > 'Z')
                    throw new ArgumentException($"Carácter inválido: {c}");

                result = result * 26 + (c - 'A' + 1);
            }

            return result;
        }

        public static string ColumnNumberToLetter(int columnNumber)
        {
            if (columnNumber < 1)
                throw new ArgumentOutOfRangeException(nameof(columnNumber), "El número de columna debe ser mayor a 0");

            string columnName = "";

            while (columnNumber > 0)
            {
                columnNumber--; // Ajuste base 0
                int remainder = columnNumber % 26;
                columnName = (char)('A' + remainder) + columnName;
                columnNumber /= 26;
            }

            return columnName;
        }
    }
}
