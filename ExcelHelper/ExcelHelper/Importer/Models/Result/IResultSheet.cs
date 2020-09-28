using System.Collections.Generic;
using ExcelHelper.Excel.Importer.Models.Import;

namespace ExcelHelper.Excel.Importer.Models.Result
{
    public interface IResultSheet : IImportSheet
    {
        bool IsValidated { get; }

        bool IsUniqueValidated { get; set; }

        string SheetFormatErrorMessage { get; set; }

        string UniqueValidateErrorMessage { get; }

        IEnumerable<ImportModel> ErrorRows { get; }

        List<List<RepeatRow>> RepeatedRowIndexes { get; }

        void Validate();

        string GetSummaryErrorMessage();
    }
}
