using NPOI.HSSF.Util;
using System.Collections.Generic;

namespace ExcelHelper.Importer.Dtos
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

        public IEnumerable<IBookSheet> Sheets { get; set; }

        public ImportBook SetSheets(params IBookSheet[] list)
        {
            this.Sheets = list;
            return this;
        }
    }
}
