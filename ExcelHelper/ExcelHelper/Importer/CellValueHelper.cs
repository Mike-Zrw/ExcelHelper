using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelHelper.Importer
{
    /// <summary>
    /// 单元格帮助
    /// </summary>
    public class CellValueHelper
    {
        /// <summary>
        /// 根据Excel列类型获取列的值.
        /// </summary>
        /// <param name="cell">cell.</param>
        /// <returns>值.</returns>
        public static string GetCellValue(ICell cell)
        {
            if (cell == null)
            {
                return string.Empty;
            }

            switch (cell.CellType)
            {
                case CellType.Blank:
                    return string.Empty;
                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                case CellType.Error:
                    return cell.ErrorCellValue.ToString();
                case CellType.Numeric:
                    return HSSFDateUtil.IsCellDateFormatted(cell) ? $"{cell.DateCellValue:G}" : cell.NumericCellValue.ToString();

                case CellType.Unknown:
                default:
                    return cell.ToString();
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Formula:
                    try
                    {
                        var da = cell.Sheet.Workbook.GetType().Name;
                        if (da == "HSSFWorkbook")
                        {
                            var e = new HSSFFormulaEvaluator(cell.Sheet.Workbook);
                            e.EvaluateInCell(cell);
                            return cell.ToString();
                        }
                        else
                        {
                            var e = new XSSFFormulaEvaluator(cell.Sheet.Workbook);
                            e.EvaluateInCell(cell);
                            return cell.ToString();
                        }
                    }
                    catch
                    {
                        return cell.NumericCellValue.ToString();
                    }
            }
        }
    }
}
