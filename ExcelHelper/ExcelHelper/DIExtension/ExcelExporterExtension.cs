using ExcelHelper.Exporter;
using Microsoft.Extensions.DependencyInjection;

namespace ExcelHelper.DIExtension
{
    public static class ExcelExporterExtension
    {
        public static void AddExcelExporter(this IServiceCollection service)
        {
            service.AddSingleton(typeof(IExcelExporter), typeof(DefaultExcelExporter));
        }
    }
}
