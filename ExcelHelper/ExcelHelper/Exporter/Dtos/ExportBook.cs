using System.Collections.Generic;
using ExcelHelper.Common;

namespace ExcelHelper.Exporter.Dtos
{
    public class ExportBook
    {
        public IEnumerable<BookSheet> Sheets { get; set; }

        public ExtEnum Ext { get; set; } = ExtEnum.Xlsx;
    }
}
