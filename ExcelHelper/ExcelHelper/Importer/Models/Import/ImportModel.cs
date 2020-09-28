using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelHelper.Excel.Importer.Models.Import
{
    public class ImportModel
    {
        /// <summary>
        /// 字段名和excel中坐标
        /// </summary>
        private Dictionary<string, int> columnNameToIndex;

        public IEnumerable<ImportColumnProperty> ColumnProperties { get; set; }

        /// <summary>
        /// 字段名和错误
        /// </summary>
        public Dictionary<string, string> ColumnNameError { get; private set; }

        public int RowIndex { get; set; }

        /// <summary>
        /// 数据唯一标识
        /// </summary>
        public string UniqueSign { get; private set; }

        public List<int> UniqueColumnIndexes { get; private set; }

        /// <summary>
        /// 字段在excel中列坐标和错误
        /// </summary>
        public Dictionary<int, string> ColumnIndexError { get; private set; }

        public bool IsValidated { get => this.ColumnNameError == null; }

        public void SetError(string columnName, string errorMsg)
        {
            if (this.ColumnNameError == null)
            {
                this.ColumnNameError = new Dictionary<string, string>();
            }

            if (this.ColumnNameError.TryGetValue(columnName, out string oldError))
            {
                this.ColumnNameError[columnName] = $"{errorMsg},{oldError}";
            }
            else
            {
                this.ColumnNameError.Add(columnName, errorMsg);
            }

            this.ColumnIndexError = this.ColumnNameError.ToDictionary(m => this.columnNameToIndex[m.Key], m => m.Value);
        }

        public void SetColumnProperties(IEnumerable<ImportColumnProperty> columnProperties)
        {
            this.ColumnProperties = columnProperties;
            this.columnNameToIndex = columnProperties.ToDictionary(m => m.PropertyInfo.Name, m => m.ColumnIndex);
        }

        public void GenerateUniqueSign()
        {
            this.UniqueColumnIndexes = new List<int>();
            var uniqueProperties = this.ColumnProperties.Where(m => m.IsUnique);
            if (this.ColumnIndexError?.Any() ?? false)
            {
                if (uniqueProperties.Any(p => this.ColumnIndexError.Keys.Contains(p.ColumnIndex)))
                {
                    this.UniqueSign = Guid.NewGuid().ToString();
                    return;
                }
            }

            var sign = new StringBuilder();
            foreach (var item in uniqueProperties)
            {
                this.UniqueColumnIndexes.Add(item.ColumnIndex);
                var value = item.PropertyInfo.GetValue(this);
                sign.Append(value?.ToString() ?? string.Empty).Append(",");
            }

            this.UniqueSign = sign.ToString();
        }
    }
}
