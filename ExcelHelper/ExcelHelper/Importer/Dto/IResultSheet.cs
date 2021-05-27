using System.Collections.Generic;

namespace ExcelHelper.Importer.Dto
{
    public interface IResultSheet : IBookSheet
    {
        bool IsValidated { get; }

        bool IsUniqueValidated { get; set; }

        string SheetFormatErrorMessage { get; set; }

        string UniqueValidateErrorMessage { get; }

        IEnumerable<SheetRow> ErrorRows { get; }

        List<List<RepeatRow>> RepeatedRowIndexes { get; }

        IEnumerable<T> GetData<T>();

        void Validate();

        string GetSummaryErrorMessage();
    }
}
