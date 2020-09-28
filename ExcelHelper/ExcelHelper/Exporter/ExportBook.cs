using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelHelper.Excel.Exporter
{
    public class ExportBook
    {
        public IEnumerable<ExportSheet> Sheets { get; set; }

        public ExtEnum Ext { get; set; } = ExtEnum.XLSX;
    }
}
