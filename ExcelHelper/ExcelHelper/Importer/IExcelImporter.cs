using System.IO;
using ExcelHelper.Excel.Importer.Models.Import;
using ExcelHelper.Excel.Importer.Models.Result;

namespace ExcelHelper.Excel.Importer
{
    /// <summary>
    /// Excel导入
    /// </summary>
    public interface IExcelImporter
    {
        /// <summary>
        /// 导入Excel
        /// </summary>
        /// <param name="fileStream">excel文件流</param>
        /// <param name="ext">excel后缀</param>
        /// <param name="importBook">导入模型</param>
        /// <param name="outPutErrorStream">错误输出流</param>
        /// <returns>导入结果</returns>
        ImportResult ImportExcel(Stream fileStream, ExtEnum ext, ImportBook importBook, Stream outPutErrorStream = null);
    }
}
