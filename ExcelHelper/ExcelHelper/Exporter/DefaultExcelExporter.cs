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
    public class DefaultExcelExporter : IExcelExporter
    {
        private const string DefaultDateFormat = "yyyy-MM-dd HH:mm:ss";

        public Stream Export(ExportBook book, Stream stream)
        {
            if (stream == null)
            {
                throw new Exception("stream is null");
            }

            if (!book.Sheets?.Any() ?? true)
            {
                throw new Exception("sheets is empty");
            }

            IWorkbook workbook = WorkbookGenerator.GetIWorkbook(book.Ext);
            foreach (var item in book.Sheets)
            {
                this.CreateSheet(workbook, item.SheetName, item.Title, item.Data, item.FilterColumn);
            }

            workbook.Write(stream);
            workbook.Close();
            return stream;
        }

        public ISheet CreateSheet<T>(IWorkbook workbook, string sheetName, SheetTitle title, IEnumerable<T> data, IEnumerable<string> filterColumn)
            where T : SheetRow
        {
            var sheet = string.IsNullOrWhiteSpace(sheetName) ? workbook.CreateSheet() : workbook.CreateSheet(sheetName);
            if (data == null)
            {
                return sheet;
            }

            var columnProperties = this.GetColumnProperties(data, filterColumn);

            int rowIndex = 0;
            if (title != null)
            {
                this.CreatetTitle(workbook, title, sheet, columnProperties.Count - 1);
                rowIndex++;
            }

            this.CreateHeader(workbook, sheet, columnProperties, rowIndex++);

            var columnStyles = this.CreateColumnStyles(workbook, columnProperties);
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
                    cell.CellStyle = columnStyles[property.ColumnIndex];
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
                if (endRow == 0)
                {
                    return;
                }

                foreach (var columnId in columnIds)
                {
                    sheet2.AddMergedRegion(new CellRangeAddress(beginRow, endRow, columnId, columnId));
                }
            }
        }

        private List<ExportColumnProperty> GetColumnProperties<T>(IEnumerable<T> data, IEnumerable<string> filterColumn)
            where T : SheetRow
        {
            var columnProperties = data.GetType().GetGenericArguments()[0].GetProperties()
                                            .Where(x => x.GetCustomAttribute(typeof(ColumnNameAttribute)) is ColumnNameAttribute)
                                            .Select(m => new ExportColumnProperty(m)).ToList();
            if (filterColumn?.Any() ?? false)
            {
                columnProperties.RemoveAll(m => !(filterColumn.Any(p => string.Compare(p, m.Name, true) == 0) || filterColumn.Any(p => string.Compare(p, m.PropertyInfo.Name, true) == 0)));
            }

            for (int i = 0; i < columnProperties.Count; i++)
            {
                columnProperties[i].ColumnIndex = i;
            }

            return columnProperties;
        }

        private void CreatetTitle(IWorkbook workbook, SheetTitle title, ISheet sheet, int columnCount)
        {
            ICellStyle cellstyle = workbook.CreateCellStyle();
            cellstyle.Alignment = (HorizontalAlignment)title.HorizontalAlign;
            var font = workbook.CreateFont();
            font.IsBold = title.IsBold;
            font.FontName = title.FontName;
            font.FontHeightInPoints = title.FontSize;
            font.Color = title.FontColor;
            cellstyle.SetFont(font);

            var titleRow = sheet.CreateRow(0);
            titleRow.CreateCell(0).SetCellValue(title.Title);
            CellRangeAddress region = new CellRangeAddress(0, 0, 0, columnCount);
            sheet.AddMergedRegion(region);
            sheet.GetRow(0).GetCell(0).CellStyle = cellstyle;
        }

        private int CreateHeader(IWorkbook workbook, ISheet sheet, IEnumerable<ExportColumnProperty> columnProperties, int rowIndex)
        {
            IRow headerRow = sheet.CreateRow(rowIndex++);
            foreach (var item in columnProperties)
            {
                var cell = headerRow.CreateCell(item.ColumnIndex);
                cell.SetCellValue(item.Name);

                var fontStyle = workbook.CreateCellStyle();
                var font = workbook.CreateFont();
                font.IsBold = item.HeaderStyle.IsBold;
                font.FontName = item.HeaderStyle.FontName;
                font.FontHeightInPoints = item.HeaderStyle.FontSize;
                font.Color = item.HeaderStyle.FontColor;
                fontStyle.SetFont(font);
                fontStyle.VerticalAlignment = (VerticalAlignment)item.HeaderStyle.VerticalAlign;
                fontStyle.Alignment = (HorizontalAlignment)item.HeaderStyle.HorizontalAlign;
                cell.CellStyle = fontStyle;
            }

            return rowIndex;
        }

        private Dictionary<int, ICellStyle> CreateColumnStyles(IWorkbook workbook, List<ExportColumnProperty> columnProperties)
        {
            var cellStyles = new Dictionary<int, ICellStyle>();
            columnProperties.ForEach((property) =>
            {
                var fontStyle = workbook.CreateCellStyle();
                var font = workbook.CreateFont();
                font.IsBold = property.ColumnStyle.IsBold;
                font.FontName = property.ColumnStyle.FontName;
                font.FontHeightInPoints = property.ColumnStyle.FontSize;
                font.Color = property.ColumnStyle.FontColor;
                fontStyle.SetFont(font);
                fontStyle.VerticalAlignment = (VerticalAlignment)property.ColumnStyle.VerticalAlign;
                fontStyle.Alignment = (HorizontalAlignment)property.ColumnStyle.HorizontalAlign;
                fontStyle.WrapText = property.ColumnStyle.WrapText;
                cellStyles.Add(property.ColumnIndex, fontStyle);
            });
            return cellStyles;
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
    }
}
