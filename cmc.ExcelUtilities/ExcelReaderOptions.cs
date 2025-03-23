using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace cmc.ExcelUtilities
{
    public class ExcelReaderOptions
    {
        public bool HasHeaderRow { get; set; }
        public int? StartRowData { get; set; }
        public int? EndRowData { get; set; }
    }
}
