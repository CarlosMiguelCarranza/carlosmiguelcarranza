using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace cmc.ExcelUtilities
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ExcelColumnAttribute : Attribute
    {
        public string ColumnName { get; set; }
        public ExcelColumnAttribute(string columnName)
        {
            ColumnName = columnName;
        }
    }
}
