using System.Reflection;
using ExcelHelper.Attributes;

namespace ExcelHelper.Exporter
{
    public class ExportColumnProperty
    {
        public ExportColumnProperty(PropertyInfo property)
        {
            this.Name = property.GetCustomAttribute(typeof(ColumnNameAttribute)) is ColumnNameAttribute nameAttr ? nameAttr.Name : property.Name;
            if (property.GetCustomAttribute(typeof(ColumnWidthAttribute)) is ColumnWidthAttribute widthAttr)
            {
                this.MinWidth = widthAttr.MinWidth;
                this.MaxWidth = widthAttr.MaxWidth;
            }

            if (property.GetCustomAttribute(typeof(HeaderStyleAttribute)) is HeaderStyleAttribute headerAttr)
            {
                this.HeaderStyle = headerAttr;
            }

            if (property.GetCustomAttribute(typeof(ColumnStyleAttribute)) is ColumnStyleAttribute columnAttr)
            {
                this.ColumnStyle = columnAttr;
            }

            if (property.GetCustomAttribute(typeof(StringFormatterAttribute)) is StringFormatterAttribute formatAttr)
            {
                this.StringFormat = formatAttr.Format;
            }

            if (property.GetCustomAttribute(typeof(RowMergedAttribute)) is RowMergedAttribute)
            {
                this.RowMerged = true;
            }

            this.PropertyInfo = property;
        }

        public int MinWidth { get; set; }

        public int MaxWidth { get; set; }

        public int ColumnIndex { get; set; }

        public string Name { get; set; }

        public HeaderStyleAttribute HeaderStyle { get; set; } = new HeaderStyleAttribute();

        public ColumnStyleAttribute ColumnStyle { get; set; } = new ColumnStyleAttribute();

        public string StringFormat { get; set; }

        public bool RowMerged { get; set; }

        public PropertyInfo PropertyInfo { get; set; }
    }
}
