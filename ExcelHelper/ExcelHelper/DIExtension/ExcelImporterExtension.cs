using ExcelHelper.Importer;
using Microsoft.Extensions.DependencyInjection;

namespace ExcelHelper.DIExtension
{
    public static class ExcelImporterExtension
    {
        public static IServiceCollection AddExcelImporter(this IServiceCollection service)
        {
            service.AddSingleton(typeof(IExcelImporter), typeof(DefaultExcelImporter));
            return service;
        }
    }
}
