using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using ExcelHelper.Attributes;
using ExcelHelper.Importer.Models;
using ExcelHelper.Importer.Models.Import;
using ExcelHelper.Importer.Models.Result;
using NPOI.SS.UserModel;

namespace ExcelHelper.Importer
{
    /// <summary>
    /// excel导入
    /// </summary>
    public class DefaultExcelImporter : IExcelImporter
    {
        /// <summary>
        /// 导入Excel
        /// </summary>
        /// <param name="fileStream">excel文件流</param>
        /// <param name="ext">excel后缀</param>
        /// <param name="importBook">导入模型</param>
        /// <param name="outPutErrorStream">错误输出流</param>
        /// <returns>导入结果</returns>
        public ImportResult ImportExcel(Stream fileStream, ExtEnum ext, ImportBook importBook, Stream outPutErrorStream = null)
        {
            var ret = new ImportResult();
            var sheets = importBook.Sheets.Select(m => this.CreateResultSheetInstance(m.GetType().GenericTypeArguments[0], m)).ToArray();
            ret.SetSheets(sheets);
            IWorkbook workbook;
            try
            {
                workbook = WorkbookGenerator.GetIWorkbook(fileStream, ext);
            }
            catch (Exception ex)
            {
                ret.SetBookFormatErrorMessage(ex.Message, ex);
                return ret;
            }

            var errorStyleGenerator = new ImporterErrorStyleGenerator(workbook, importBook.DataErrorForegroundColor, importBook.RepeatedErrorForegroundColor, importBook.DefaultForegroundColor);

            for (var i = 0; i < workbook.NumberOfSheets; i++)
            {
                var sheet = workbook.GetSheetAt(i);
                var sheetModel = ret.Sheets.FirstOrDefault(m => m.SheetIndex == i || m.SheetName == sheet.SheetName);
                if (sheetModel == null)
                {
                    continue;
                }

                sheetModel.SheetIndex = i;
                sheetModel.SheetName = sheet.SheetName;

                this.ParseSheetToModel(sheet, sheetModel);
                if (outPutErrorStream != null)
                {
                    errorStyleGenerator.InitStyle(sheet, sheetModel.HeaderRowIndex);
                    errorStyleGenerator.SetErrorStyle(sheet, sheetModel);
                }
            }

            if (outPutErrorStream != null)
            {
                workbook.Write(outPutErrorStream);
            }

            return ret;
        }

        private IResultSheet CreateResultSheetInstance(Type genericType, IImportSheet sheetModel)
        {
            var instance = (IResultSheet)Activator.CreateInstance(typeof(ResultSheet<>).MakeGenericType(genericType));

            // 属性值拷贝到新实例
            var parentProperties = sheetModel.GetType().GetProperties().ToList();
            parentProperties.ForEach((property) =>
            {
                property.SetValue(instance, property.GetValue(sheetModel));
            });
            return instance;
        }

        private IResultSheet ParseSheetToModel(ISheet sheet, IResultSheet sheetModel)
        {
            var sheetModelType = sheetModel.GetType();
            var validateSheetMethod = this.GetType().GetMethod(nameof(this.ValidateSheetFormat), BindingFlags.NonPublic | BindingFlags.Instance).MakeGenericMethod(sheetModelType.GetGenericArguments()[0]);
            var validateSheetResult = validateSheetMethod.Invoke(this, new object[] { sheet, sheetModel.HeaderRowIndex }).ToString();
            if (!string.IsNullOrWhiteSpace(validateSheetResult))
            {
                sheetModel.SheetFormatErrorMessage = validateSheetResult;
                return sheetModel;
            }

            var sheetToModelMethod = this.GetType().GetMethod(nameof(this.FillSheetRow), BindingFlags.NonPublic | BindingFlags.Instance).MakeGenericMethod(sheetModelType.GetGenericArguments()[0]);
            var data = sheetToModelMethod.Invoke(this, new object[] { sheet, sheetModel.HeaderRowIndex });

            sheetModelType.GetMethod("SetData", BindingFlags.Public | BindingFlags.Instance).Invoke(sheetModel, new object[] { data });

            sheetModel.Validate();
            return sheetModel;
        }

        private IEnumerable<T> FillSheetRow<T>(ISheet sheet, int headerRowIndex)
            where T : ImportModel
        {
            var headerRow = sheet.GetRow(headerRowIndex);
            var modelType = typeof(T);
            var columnProperties = this.GetCellProperties(modelType, headerRow);

            for (var rowIndex = headerRowIndex + 1; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                if (row == null || row.FirstCellNum < 0)
                {
                    continue;
                }

                var rowModel = (ImportModel)Activator.CreateInstance(modelType, null);

                rowModel.SetColumnProperties(columnProperties);
                rowModel.RowIndex = rowIndex;

                var indextoProperty = columnProperties.ToDictionary(m => m.ColumnIndex, m => m);
                for (int cellIndex = 0; cellIndex < headerRow.LastCellNum; cellIndex++)
                {
                    if (!indextoProperty.TryGetValue(cellIndex, out ImportColumnProperty columnProperty))
                    {
                        continue;
                    }

                    var cell = row.GetCell(cellIndex);
                    try
                    {
                        this.FillPropertyValue(rowModel, columnProperty, cell);
                    }
                    catch (Exception ex)
                    {
                        rowModel.SetError(columnProperty.PropertyInfo.Name, $"数据解析异常.{ex.Message}");
                    }
                }

                rowModel.GenerateUniqueSign();

                yield return (T)rowModel;
            }
        }

        private string ValidateSheetFormat<T>(ISheet sheet, int headerRowIndex)
        {
            var headerRow = sheet.GetRow(headerRowIndex);
            if (headerRow == null || sheet.LastRowNum == 0)
            {
                return "表格没有数据";
            }

            var modelType = typeof(T);
            var columnProperties = this.GetCellProperties(modelType, headerRow);
            if (columnProperties.Any(m => m.ColumnIndex == -1))
            {
                var notFindNames = columnProperties.Where(m => m.ColumnIndex == -1).Select(m => m.Name);
                return $"未找到列: {string.Join(",", notFindNames)}";
            }

            return string.Empty;
        }

        private void FillPropertyValue(ImportModel rowModel, ImportColumnProperty columnProperty, ICell cell)
        {
            var cellValue = cell == null ? null : CellValueHelper.GetCellValue(cell);

            if (string.IsNullOrWhiteSpace(cellValue))
            {
                if (columnProperty.IsRequired || (!columnProperty.PropertyInfo.IsNullable() && columnProperty.PropertyInfo.PropertyType != typeof(string)))
                {
                    rowModel.SetError(columnProperty.PropertyInfo.Name, columnProperty.EmptyErrorMessage);
                }
            }
            else
            {
                if (columnProperty.HasRegex && !Regex.IsMatch(cellValue, columnProperty.RegexPattern, columnProperty.RegexOptions))
                {
                    rowModel.SetError(columnProperty.PropertyInfo.Name, columnProperty.RegexErrorMessage);
                    return;
                }

                var value = columnProperty.PropertyInfo.IsNullable() ? Convert.ChangeType(cellValue, columnProperty.PropertyInfo.PropertyType.GetGenericArguments()[0]) : Convert.ChangeType(cellValue, columnProperty.PropertyInfo.PropertyType);
                columnProperty.PropertyInfo.SetValue(rowModel, value);
            }
        }

        private List<ImportColumnProperty> GetCellProperties(Type modelType, IRow headerRow)
        {
            var properties = modelType.GetProperties();
            var columnProperties = new List<ImportColumnProperty>();
            foreach (var p in properties)
            {
                if (p.GetCustomAttribute(typeof(ColumnNameAttribute)) is ColumnNameAttribute attr)
                {
                    columnProperties.Add(new ImportColumnProperty(p, attr));
                }
            }

            for (int i = headerRow.FirstCellNum; i < headerRow.LastCellNum; i++)
            {
                var cell = headerRow.GetCell(i);
                if (cell == null)
                {
                    continue;
                }

                var headColumnName = CellValueHelper.GetCellValue(cell);
                if (string.IsNullOrWhiteSpace(headColumnName))
                {
                    continue;
                }

                headColumnName = headColumnName.Trim();
                var columnProperty = columnProperties.FirstOrDefault(x => x.Name == headColumnName.Trim());
                if (columnProperty != null)
                {
                    columnProperty.ColumnIndex = i;
                }
            }

            // 没有字段设置唯一认证，则给所有字段加上唯一认证
            if (!columnProperties.Any(m => m.IsUnique))
            {
                columnProperties.ForEach(m => m.IsUnique = true);
            }

            return columnProperties;
        }
    }
}
