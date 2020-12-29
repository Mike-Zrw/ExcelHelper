using System;
using System.Collections.Generic;

namespace ExcelHelper.Importer.Dtos
{
    public class BookSheet<T> : IBookSheet
        where T : SheetRow
    {
        public int HeaderRowIndex { get; set; }

        public int SheetIndex { get; set; }

        public string SheetName { get; set; }

        /// <summary>
        /// 是否需要唯一验证
        /// </summary>
        public bool NeedUniqueValidation { get; set; } = true;

        /// <summary>
        /// 唯一验证提示
        /// </summary>
        public string UniqueValidationPrompt { get; set; } = "数据重复";

        public Action<List<T>> ValidateHandler { get; set; }
    }
}
