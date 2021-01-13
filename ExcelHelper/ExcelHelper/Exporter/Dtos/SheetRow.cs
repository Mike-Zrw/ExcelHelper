using System.Collections.Generic;

namespace ExcelHelper.Exporter.Dtos
{
    public class SheetRow
    {
        /// <summary>
        /// 唯一键
        /// </summary>
        public string ExportPrimaryKey { get; set; }

        public int ExportRowIndex { get; set; }

        public Dictionary<string, CellStyle> CellStyles { get; private set; }

        public void SetCellStyle(string cellName, CellStyle style)
        {
            if (CellStyles == null) CellStyles = new Dictionary<string, CellStyle>();
            if (CellStyles.ContainsKey(cellName))
                CellStyles[cellName] = style;
            else
                CellStyles.Add(cellName, style);
        }
    }
}