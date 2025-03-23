using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace cmc.ExcelUtilities
{
    public class ExcelSheetSchemaBase
    {
        public required int Key { get; set; }
        public ExcelRowValidationResults ValidationResult { get; } = new();
    }
}
