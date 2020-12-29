using ExcelHelper.Exporter.Dtos;
using System.IO;

namespace ExcelHelper.Exporter
{
    public interface IExcelExporter
    {
        Stream Export(ExportBook book, Stream stream);
    }
}
