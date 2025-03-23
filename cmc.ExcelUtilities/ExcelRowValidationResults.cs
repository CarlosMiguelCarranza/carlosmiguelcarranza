using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace cmc.ExcelUtilities
{
    public class ExcelRowValidationResults
    {
        public bool IsValid => CellValidationResults.Count == 0;
        public List<ExcelCellValidationResult> CellValidationResults { get; set; } = new();
        public void Add(string column, int row, string message)
        {
            CellValidationResults.Add(new ExcelCellValidationResult(column, row, message));
        }
    }

    public record ExcelCellValidationResult(string Column, int Row, string Message);
}
