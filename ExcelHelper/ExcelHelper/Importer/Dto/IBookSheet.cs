namespace ExcelHelper.Importer.Dto
{
    public interface IBookSheet
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
