using ExcelHelper.Attributes;
using ExcelHelper.Importer.Attributes;
using System.Reflection;
using System.Text.RegularExpressions;

namespace ExcelHelper.Importer
{
    public class ImportColumnProperty
    {
        public ImportColumnProperty(string name, PropertyInfo propertyInfo)
        {
            this.Name = name;
            this.PropertyInfo = propertyInfo;
        }

        public ImportColumnProperty(PropertyInfo property, ColumnNameAttribute attr)
        {
            this.Name = attr.Name;
            this.PropertyInfo = property;
            if (property.GetCustomAttribute(typeof(ColumnRegexAttribute)) is ColumnRegexAttribute attr2)
            {
                this.SetRegexData(attr2);
            }

            if (property.GetCustomAttribute(typeof(ColumnRequiredAttribute)) is ColumnRequiredAttribute attr3)
            {
                this.SetRequiredData(attr3);
            }

            if (property.GetCustomAttribute(typeof(ColumnUniqueAttribute)) is ColumnUniqueAttribute)
            {
                this.IsUnique = true;
            }
        }

        /// <summary>
        /// 展示名称
        /// </summary>
        public string Name { get; set; }

        public int ColumnIndex { get; set; } = -1;

        public PropertyInfo PropertyInfo { get; set; }

        /// <summary>
        /// 是否需要正则验证
        /// </summary>
        public bool HasRegex { get; set; }

        public string RegexPattern { get; set; }

        public RegexOptions RegexOptions { get; set; }

        public string RegexErrorMessage { get; set; }

        /// <summary>
        /// 是否是必须的（不可为空）
        /// </summary>
        public bool IsRequired { get; set; }

        /// <summary>
        /// 是否是唯一的
        /// </summary>
        public bool IsUnique { get; set; }

        public string EmptyErrorMessage { get; set; } = "数据不可为空";

        public void SetRegexData(ColumnRegexAttribute attr)
        {
            this.HasRegex = attr.Pattern != null;
            this.RegexPattern = attr.Pattern;
            this.RegexErrorMessage = attr.ErrorMessage;
            this.RegexOptions = attr.RegexOptions;
        }

        public void SetRequiredData(ColumnRequiredAttribute attr)
        {
            this.IsRequired = true;
            this.EmptyErrorMessage = attr.EmptyErrorMessage;
        }
    }
}
