using cmc.ExcelUtilities;
using Microsoft.Extensions.Options;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace cmd.TestUtilities.Models
{
    public class MyExcelReader : ExcelReaderBase<ExcelTest1>
    {
        public MyExcelReader(IOptions<ExcelReaderOptions> options) : base(options.Value)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }
    }
}
