using System.Collections.Generic;

namespace ExcelHelper.Importer.Dtos
{
    public class RepeatRow
    {
        public int RowIndex { get; set; }

        public List<int> ColumnIndexes { get; set; }
    }
}
