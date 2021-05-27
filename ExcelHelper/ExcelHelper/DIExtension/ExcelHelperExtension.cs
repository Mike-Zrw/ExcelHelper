using ExcelHelper.Exporter;
using ExcelHelper.Importer;
using Microsoft.Extensions.DependencyInjection;

namespace ExcelHelper.DIExtension
{
    public static class ExcelHelperExtension
    {
        public static IServiceCollection AddExcelHelper(this IServiceCollection service)
        {
            service.AddSingleton(typeof(IExcelExporter), typeof(DefaultExcelExporter));
            service.AddSingleton(typeof(IExcelImporter), typeof(DefaultExcelImporter));
            return service;
        }
    }
}
