using cmc.ExcelUtilities;
using cmd.TestUtilities.Models;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;
using System.Reflection;
using static Microsoft.IO.RecyclableMemoryStreamManager;

namespace cmd.TestUtilities
{
    public class Tests
    {
        private ServiceProvider _serviceProvider;

        [SetUp]
        public void Setup()
        {
            var services = new ServiceCollection();

            services.AddOptions();

            services.Configure<ExcelReaderOptions>(n =>
            {
                n.HasHeaderRow = true;
            });

            services.AddTransient<MyExcelReader>();

            _serviceProvider = services.BuildServiceProvider();
        }

        [TearDown]
        public void TearDown()
        {
            _serviceProvider.Dispose();
        }

        [Test]
        public async Task ReadExcelFile_ValidData_AllItemsAreValid()
        {
            // Arrange
            var filePath = Path.Combine(TestContext.CurrentContext.TestDirectory,
                                      "Resources",
                                      "ExcelTest1.xlsx");

            var service = _serviceProvider.GetRequiredService<MyExcelReader>();

            // Configuración inicial y verificaciones
            Assert.That(File.Exists(filePath), $"Archivo de prueba no encontrado: {filePath}");

            // Act
            using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                await service.ReadAsync(fileStream);
            }

            var result = await service.GetData();

            // Assert
            var validationErrors = result
                .Where(item => !item.Value.ValidationResult.IsValid)
                .ToList();

            Assert.Multiple(() =>
            {
                foreach (var error in validationErrors)
                {
                    var errorDetails = error.Value.ValidationResult.CellValidationResults
                        .Select(e => $"Fila {e.Row}, Columna {e.Column}: {e.Message}");

                    Assert.That(
                        errorDetails,
                        Is.Empty,
                        $"Errores en registro {error.Key}:\n{string.Join("\n", errorDetails)}");
                }
            });
        }
    }
}
