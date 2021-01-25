using ExcelHelper.Exporter.Dtos;
using System.IO;

namespace ExcelHelper.Exporter
{
    public interface IExcelExporter
    {
        void Export(ExportBook book, Stream stream);
        
        byte[] Export(ExportBook book);
    }
}
