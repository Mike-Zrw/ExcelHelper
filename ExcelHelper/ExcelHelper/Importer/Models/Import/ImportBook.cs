using System.Collections.Generic;
using NPOI.HSSF.Util;

namespace ExcelHelper.Importer.Models.Import
{
    public class ImportBook
    {
        /// <summary>
        /// 错误前景色
        /// </summary>
        public short DataErrorForegroundColor { get; set; } = HSSFColor.Red.Index;

        /// <summary>
        /// 重复前景色
        /// </summary>
        public short RepeatedErrorForegroundColor { get; set; } = HSSFColor.Yellow.Index;

        /// <summary>
        /// 默认前景色
        /// </summary>
        public short DefaultForegroundColor { get; set; } = HSSFColor.White.Index;

        public IEnumerable<IImportSheet> Sheets { get; set; }

        public void SetSheetModels(params IImportSheet[] list)
        {
            this.Sheets = list;
        }
    }
}
