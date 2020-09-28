using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelHelper.Excel.Importer.Models.Result
{
    public class RepeatRow
    {
        public int RowIndex { get; set; }

        public List<int> ColumnIndexes { get; set; }
    }
}
