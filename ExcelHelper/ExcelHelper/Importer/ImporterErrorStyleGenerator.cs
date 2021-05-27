using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Linq;
using ExcelHelper.Importer.Dto;

namespace ExcelHelper.Importer
{
    public class ImporterErrorStyleGenerator
    {
        private readonly IWorkbook _workBook;
        private readonly ICellStyle _defaultStyle;
        private readonly ICellStyle _dataErrorStyle;
        private readonly ICellStyle _rowRepeatedErrorStyle;
        private readonly ICreationHelper _commentFactory;
        private readonly IClientAnchor _commentAnchor;
        public ImporterErrorStyleGenerator(IWorkbook workBook, short dataErrorForegroundColor, short repeatedErrorForegroundColor, short defaultForegroundColor)
        {

            _workBook = workBook;
            _defaultStyle = _workBook.CreateCellStyle();
            _defaultStyle.FillForegroundColor = defaultForegroundColor;

            _dataErrorStyle = _workBook.CreateCellStyle();
            _dataErrorStyle.FillForegroundColor = dataErrorForegroundColor;
            _dataErrorStyle.FillPattern = FillPattern.SolidForeground;

            _rowRepeatedErrorStyle = _workBook.CreateCellStyle();
            _rowRepeatedErrorStyle.FillForegroundColor = repeatedErrorForegroundColor;
            _rowRepeatedErrorStyle.FillPattern = FillPattern.SolidForeground;

            _commentFactory = _workBook.GetCreationHelper();
            _commentAnchor = _commentFactory.CreateClientAnchor();
        }

        public void InitStyle(ISheet sheet, int headerRowIndex)
        {
            for (var rowIndex = headerRowIndex + 1; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                foreach (var cell in row.Cells)
                {
                    cell.CellComment = null;
                    if (cell.CellStyle == null || (cell.CellStyle.FillForegroundColor != _dataErrorStyle.FillForegroundColor && cell.CellStyle.FillForegroundColor != _rowRepeatedErrorStyle.FillForegroundColor))
                    {
                        continue;
                    }

                    if (cell.CellStyle != null)
                        cell.CellStyle.FillForegroundColor = _defaultStyle.FillForegroundColor;
                }
            }
        }

        public void SetSheetErrorStyle(ISheet sheet, IResultSheet sheetModel)
        {
            var commentDrawing = sheet.CreateDrawingPatriarch();
            if (sheetModel.ErrorRows?.Any() ?? false)
            {
                SetSheetDataErrorStyle(sheet, commentDrawing, sheetModel.ErrorRows);
            }

            if (!sheetModel.IsUniqueValidated)
            {
                SetSheetRowRepeatedErrorStyle(sheet, commentDrawing, sheetModel.RepeatedRowIndexes, sheetModel.UniqueValidationPrompt);
            }
        }

        private void SetSheetDataErrorStyle(ISheet sheet, IDrawing commentDrawing, IEnumerable<SheetRow> errorRows)
        {
            foreach (var item in errorRows)
            {
                foreach (var errorColumn in item.ColumnIndexError)
                {
                    var cell = sheet.GetRow(item.RowIndex).GetCell(errorColumn.Key) ?? sheet.GetRow(item.RowIndex).CreateCell(errorColumn.Key);

                    SetCellErrorStyle(cell, _dataErrorStyle);

                    if (cell.CellComment == null)
                    {
                        cell.CellComment = commentDrawing.CreateCellComment(this._commentAnchor);
                    }

                    cell.CellComment.String = _commentFactory.CreateRichTextString(errorColumn.Value);
                }
            }
        }

        private void SetSheetRowRepeatedErrorStyle(ISheet sheet, IDrawing commentDrawing, List<List<RepeatRow>> rowGroups, string errorPrompt)
        {
            foreach (var rows in rowGroups)
            {
                foreach (var repeatRow in rows)
                {
                    var row = sheet.GetRow(repeatRow.RowIndex);

                    var commented = false;
                    foreach (var columnIndex in repeatRow.ColumnIndexes)
                    {
                        var cell = row.GetCell(columnIndex);

                        SetCellErrorStyle(cell, _rowRepeatedErrorStyle);

                        if (commented) continue;

                        if (cell.CellComment == null) cell.CellComment = commentDrawing.CreateCellComment(this._commentAnchor);

                        cell.CellComment.String = _commentFactory.CreateRichTextString($"{errorPrompt}.重复行：{string.Join(",", rows.Where(m => m.RowIndex != repeatRow.RowIndex).Select(m => m.RowIndex + 1))}");
                        commented = true;
                    }
                }
            }
        }

        private void SetCellErrorStyle(ICell cell, ICellStyle errorStyle)
        {
            if (cell.CellStyle == null)
            {
                cell.CellStyle = errorStyle;
            }
            else
            {
                var style = _workBook.CreateCellStyle();
                style.CloneStyleFrom(cell.CellStyle);
                style.FillForegroundColor = errorStyle.FillForegroundColor;
                style.FillPattern = errorStyle.FillPattern;
                cell.CellStyle = style;
            }
        }
    }
}
