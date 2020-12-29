using System.Collections.Generic;

namespace ExcelHelper.Exporter.Dtos
{
    public class ExportBook
    {
        public IEnumerable<BookSheet> Sheets { get; set; }

        public ExtEnum Ext { get; set; } = ExtEnum.XLSX;
    }
}
