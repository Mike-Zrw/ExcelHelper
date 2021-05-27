using System.Collections.Generic;

namespace ExcelHelper.Importer.Dto
{
    public class RepeatRow
    {
        public int RowIndex { get; set; }

        public List<int> ColumnIndexes { get; set; }
    }
}
