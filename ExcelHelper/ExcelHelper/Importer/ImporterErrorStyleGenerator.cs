using System.Collections.Generic;
using System.Linq;
using ExcelHelper.Excel.Importer.Models.Import;
using ExcelHelper.Excel.Importer.Models.Result;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelHelper.Excel.Importer
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
            this.DataErrorForegroundColor = dataErrorForegroundColor;
            this.RepeatedErrorForegroundColor = repeatedErrorForegroundColor;
            this.DefaultForegroundColor = defaultForegroundColor;

            this._workBook = workBook;
            this._defaultStyle = this._workBook.CreateCellStyle();
            this._defaultStyle.FillForegroundColor = defaultForegroundColor;

            this._dataErrorStyle = this._workBook.CreateCellStyle();
            this._dataErrorStyle.FillForegroundColor = dataErrorForegroundColor;
            this._dataErrorStyle.FillPattern = FillPattern.SolidForeground;

            this._rowRepeatedErrorStyle = this._workBook.CreateCellStyle();
            this._rowRepeatedErrorStyle.FillForegroundColor = repeatedErrorForegroundColor;
            this._rowRepeatedErrorStyle.FillPattern = FillPattern.SolidForeground;

            this._commentFactory = this._workBook.GetCreationHelper();
            this._commentAnchor = this._commentFactory.CreateClientAnchor();
        }

        public short DataErrorForegroundColor { get; set; }

        public short RepeatedErrorForegroundColor { get; set; }

        public short DefaultForegroundColor { get; set; }

        public void InitStyle(ISheet sheet, int headerRowIndex)
        {
            for (var rowIndex = headerRowIndex + 1; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                foreach (var cell in row.Cells)
                {
                    cell.CellComment = null;
                    if (cell.CellStyle == null || (cell.CellStyle.FillForegroundColor != this.DataErrorForegroundColor && cell.CellStyle.FillForegroundColor != this.RepeatedErrorForegroundColor))
                    {
                        continue;
                    }

                    cell.CellStyle = this._defaultStyle;
                }
            }
        }

        public void SetErrorStyle(ISheet sheet, IResultSheet sheetModel)
        {
            var commentDrawing = sheet.CreateDrawingPatriarch();
            if (sheetModel.ErrorRows?.Any() ?? false)
            {
                this.SetDataErrorStyle(sheet, commentDrawing, sheetModel.ErrorRows);
            }

            if (!sheetModel.IsUniqueValidated)
            {
                this.SetRowRepeatedErrorStyle(sheet, commentDrawing, sheetModel.RepeatedRowIndexes, sheetModel.UniqueValidationPrompt);
            }
        }

        private void SetDataErrorStyle(ISheet sheet, IDrawing commentDrawing, IEnumerable<ImportModel> errorRows)
        {
            foreach (var item in errorRows)
            {
                foreach (var errorColumn in item.ColumnIndexError)
                {
                    var cell = sheet.GetRow(item.RowIndex).GetCell(errorColumn.Key);
                    if (cell == null)
                    {
                        cell = sheet.GetRow(item.RowIndex).CreateCell(errorColumn.Key);
                    }

                    cell.CellStyle = this._dataErrorStyle;

                    if (cell.CellComment == null)
                    {
                        cell.CellComment = commentDrawing.CreateCellComment(this._commentAnchor);
                    }

                    cell.CellComment.String = this._commentFactory.CreateRichTextString(errorColumn.Value);
                }
            }
        }

        private void SetRowRepeatedErrorStyle(ISheet sheet, IDrawing commentDrawing, List<List<RepeatRow>> rowGroups, string errorPrompt)
        {
            foreach (var rows in rowGroups)
            {
                foreach (var repeatRow in rows)
                {
                    var row = sheet.GetRow(repeatRow.RowIndex);

                    var setedComment = false;
                    foreach (var columnIndex in repeatRow.ColumnIndexes)
                    {
                        var cell = row.GetCell(columnIndex);
                        cell.CellStyle = this._rowRepeatedErrorStyle;
                        if (setedComment)
                        {
                            continue;
                        }

                        if (cell.CellComment == null)
                        {
                            cell.CellComment = commentDrawing.CreateCellComment(this._commentAnchor);
                        }

                        cell.CellComment.String = this._commentFactory.CreateRichTextString($"{errorPrompt}.重复行：{string.Join(",", rows.Where(m => m.RowIndex != repeatRow.RowIndex).Select(m => m.RowIndex + 1))}");
                        setedComment = true;
                    }
                }
            }
        }
    }
}
