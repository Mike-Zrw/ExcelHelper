using ExcelHelper.Common;
using ExcelHelper.Exporter.Dtos;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ExcelHelper.Exporter
{
    public class DefaultExcelExporter : IExcelExporter, IDisposable
    {
        private const string DefaultDateFormat = "yyyy-MM-dd HH:mm:ss";
        private IWorkbook _workBook;
        public void Export(ExportBook book, Stream stream)
        {
            if (stream == null) throw new Exception("stream is null");

            FillWorkBook(book);
            _workBook.Write(stream);
        }

        public byte[] Export(ExportBook book)
        {
            FillWorkBook(book);

            using (var stream = new MemoryStream())
            {
                _workBook.Write(stream);
                return stream.ToArray();
            }
        }

        private void FillWorkBook(ExportBook book)
        {
            if (!book.Sheets?.Any() ?? true) throw new Exception("sheets is empty");

            _workBook = WorkbookGenerator.GetIWorkbook(book.Ext);
            foreach (var item in book.Sheets)
            {
                this.CreateSheet(item.Data.GetType().GetGenericArguments()[0], item.SheetName, item.Title, item.Data?.ToList(), item.FilterColumn?.ToList());
            }
        }

        public ISheet CreateSheet<T>(Type dateType, string sheetName, SheetTitle title, List<T> data, List<string> filterColumn)
            where T : SheetRow
        {
            var sheet = string.IsNullOrWhiteSpace(sheetName) ? _workBook.CreateSheet() : _workBook.CreateSheet(sheetName);
            if (data == null) return sheet;

            var columnProperties = GetColumnProperties(dateType, filterColumn);

            int rowIndex = 0;
            if (title != null)
            {
                this.CreateTitle(title, sheet, columnProperties.Count);
                rowIndex++;
            }

            this.CreateHeader(sheet, columnProperties, rowIndex++);

            var cacheStyles = new Dictionary<string, ICellStyle>();
            foreach (var item in data)
            {
                var columnId = 0;
                var dataRow = sheet.CreateRow(rowIndex);
                item.ExportRowIndex = rowIndex;
                rowIndex++;
                foreach (var property in columnProperties)
                {
                    var cell = dataRow.CreateCell(columnId++);
                    var value = property.PropertyInfo.GetValue(item);
                    if (value == null)
                    {
                        cell.SetCellValue(string.Empty);
                        continue;
                    }

                    string valueStr;
                    if (property.PropertyInfo.PropertyType == typeof(DateTime) || property.PropertyInfo.PropertyType == typeof(DateTime?))
                    {
                        var format = string.IsNullOrWhiteSpace(property.StringFormat) ? DefaultDateFormat : property.StringFormat;
                        valueStr = ((DateTime)property.PropertyInfo.GetValue(item)).ToString(format);
                    }
                    else
                    {
                        valueStr = property.PropertyInfo.GetValue(item).ToString();
                    }

                    cell.SetCellValue(valueStr);
                    IBaseStyle cellStyle = null;
                    if ((item.CellStyles?.Any() ?? false) && item.CellStyles.ContainsKey(property.PropertyInfo.Name))
                        cellStyle = item.CellStyles[property.PropertyInfo.Name];
                    SetCellStyle(cell, cacheStyles, property, cellStyle);
                }
            }

            this.RowMerged(data, sheet, columnProperties);

            this.SetColumnWidth(sheet, columnProperties);

            return sheet;
        }

        private void RowMerged<T>(IEnumerable<T> listData, ISheet sheet, List<ExportColumnProperty> columnProperties)
            where T : SheetRow
        {
            if (columnProperties.Any(x => x.RowMerged))
            {
                var columnIds = columnProperties.Where(x => x.RowMerged).Select(m => m.ColumnIndex).OrderBy(m => m).ToList();
                var beginRow = 0;
                var endRow = 0;
                T preData = null;
                foreach (var curData in listData)
                {
                    if (preData == null)
                    {
                        beginRow = curData.ExportRowIndex;
                        preData = curData;
                        continue;
                    }

                    if (curData.ExportPrimaryKey == preData.ExportPrimaryKey)
                    {
                        endRow = curData.ExportRowIndex;
                    }
                    else
                    {
                        TryMergeRow(sheet, columnIds, beginRow, endRow);
                        beginRow = curData.ExportRowIndex;
                        endRow = 0;
                    }

                    preData = curData;
                }

                TryMergeRow(sheet, columnIds, beginRow, endRow);
            }

            void TryMergeRow(ISheet sheet2, List<int> columnIds, int beginRow, int endRow)
            {
                if (endRow == 0) return;

                foreach (var columnId in columnIds)
                {
                    sheet2.AddMergedRegion(new CellRangeAddress(beginRow, endRow, columnId, columnId));
                }
            }
        }

        private List<ExportColumnProperty> GetColumnProperties(Type type, List<string> filterColumn)
        {
            var columnProperties = type.GetProperties()
                                            .Where(x => x.GetCustomAttribute(typeof(ColumnNameAttribute)) is ColumnNameAttribute)
                                            .Select(m => new ExportColumnProperty(m)).ToList();
            if (filterColumn?.Any() ?? false)
            {
                columnProperties.RemoveAll(m => !(filterColumn.Any(p => String.Compare(p, m.Name, StringComparison.Ordinal) == 0) || filterColumn.Any(p => String.Compare(p, m.PropertyInfo.Name, StringComparison.Ordinal) == 0)));
            }

            for (int i = 0; i < columnProperties.Count; i++)
            {
                columnProperties[i].ColumnIndex = i;
            }

            return columnProperties;
        }

        private void CreateTitle(SheetTitle title, ISheet sheet, int columnCount)
        {
            var cellStyle = _workBook.CreateCellStyle();
            cellStyle.Alignment = (HorizontalAlignment)title.HorizontalAlign;
            var font = _workBook.CreateFont();
            font.IsBold = title.IsBold;
            font.FontName = title.FontName;
            font.FontHeightInPoints = title.FontSize;
            font.Color = title.FontColor;
            cellStyle.SetFont(font);

            var titleRow = sheet.CreateRow(0);
            titleRow.CreateCell(0).SetCellValue(title.Title);
            if (columnCount > 1)
            {
                CellRangeAddress region = new CellRangeAddress(0, 0, 0, columnCount - 1);
                sheet.AddMergedRegion(region);
            }
            sheet.GetRow(0).GetCell(0).CellStyle = cellStyle;
        }

        private void CreateHeader(ISheet sheet, IEnumerable<ExportColumnProperty> columnProperties, int rowIndex)
        {
            IRow headerRow = sheet.CreateRow(rowIndex);
            foreach (var item in columnProperties)
            {
                var cell = headerRow.CreateCell(item.ColumnIndex);
                cell.SetCellValue(item.Name);

                var fontStyle = _workBook.CreateCellStyle();
                var font = _workBook.CreateFont();
                font.IsBold = item.HeaderStyle.IsBold;
                font.FontName = item.HeaderStyle.FontName;
                font.FontHeightInPoints = item.HeaderStyle.FontSize;
                font.Color = item.HeaderStyle.FontColor;
                fontStyle.SetFont(font);
                fontStyle.VerticalAlignment = (VerticalAlignment)item.HeaderStyle.VerticalAlign;
                fontStyle.Alignment = (HorizontalAlignment)item.HeaderStyle.HorizontalAlign;
                fontStyle.WrapText = item.HeaderStyle.WrapText;
                if (item.HeaderStyle.FillForegroundColor != 0)
                {
                    fontStyle.FillPattern = FillPattern.SolidForeground;
                    fontStyle.FillForegroundColor = item.HeaderStyle.FillForegroundColor;
                }
                cell.CellStyle = fontStyle;
            }
        }

        private void SetCellStyle(ICell cell, Dictionary<string, ICellStyle> cacheStyles, ExportColumnProperty property, IBaseStyle cellStyle = null)
        {
            var style = cellStyle ?? property.ColumnStyle;
            if (cacheStyles.ContainsKey(style.ToString()))
            {
                cell.CellStyle = cacheStyles[style.ToString()];
                return;
            }
            var fontStyle = _workBook.CreateCellStyle();
            var font = _workBook.CreateFont();
            font.IsBold = style.IsBold;
            font.FontName = style.FontName;
            font.FontHeightInPoints = style.FontSize;
            font.Color = style.FontColor;
            fontStyle.SetFont(font);
            fontStyle.VerticalAlignment = (VerticalAlignment)style.VerticalAlign;
            fontStyle.Alignment = (HorizontalAlignment)style.HorizontalAlign;
            if (style.FillForegroundColor != 0)
            {
                fontStyle.FillPattern = FillPattern.SolidForeground;
                fontStyle.FillForegroundColor = style.FillForegroundColor;
            }
            fontStyle.WrapText = style.WrapText;
            cell.CellStyle = fontStyle;
            cacheStyles.Add(style.ToString(), fontStyle);
        }

        private void SetColumnWidth(ISheet sheet, List<ExportColumnProperty> columnProperties)
        {
            foreach (var item in columnProperties)
            {
                sheet.AutoSizeColumn(item.ColumnIndex);
                var width = (int)(sheet.GetColumnWidth(item.ColumnIndex) * 1.2);
                if (item.MinWidth > 0 && width < item.MinWidth)
                {
                    width = item.MinWidth;
                }
                else if (item.MaxWidth > 0 && width > item.MaxWidth)
                {
                    width = item.MaxWidth;
                }

                sheet.SetColumnWidth(item.ColumnIndex, width);
            }
        }

        public void Dispose()
        {
            if (_workBook != null)
                _workBook.Close();
        }
    }
}
