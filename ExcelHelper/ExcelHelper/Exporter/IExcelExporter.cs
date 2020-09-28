using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExcelHelper.Exporter
{
    public interface IExcelExporter
    {
        Stream Export(ExportBook book, Stream stream);
    }
}
