using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelHelper.Importer.Models.Import
{
    public interface IImportSheet
    {
        int HeaderRowIndex { get; set; }

        int SheetIndex { get; set; }

        string SheetName { get; set; }

        /// <summary>
        /// 是否需要唯一验证
        /// </summary>
        bool NeedUniqueValidation { get; set; }

        /// <summary>
        /// 唯一验证提示
        /// </summary>
        string UniqueValidationPrompt { get; set; }
    }
}
