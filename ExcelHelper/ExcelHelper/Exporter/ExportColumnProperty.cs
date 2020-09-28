using System.Reflection;
using ExcelHelper.Excel.Attributes;

namespace ExcelHelper.Excel.Exporter
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
                this.HeaderIsBold = headerAttr.IsBold;
                this.HeaderFontColor = headerAttr.FontColor;
                this.HeaderFontSize = headerAttr.FontSize;
                this.HeaderFontName = headerAttr.FontName;
            }

            if (property.GetCustomAttribute(typeof(ColumnStyleAttribute)) is ColumnStyleAttribute columnAttr)
            {
                this.IsBold = columnAttr.IsBold;
                this.FontColor = columnAttr.FontColor;
                this.FontSize = columnAttr.FontSize;
                this.FontName = columnAttr.FontName;
            }

            if (property.GetCustomAttribute(typeof(StringFormatterAttribute)) is StringFormatterAttribute formatAttr)
            {
                this.StringFormat = formatAttr.Format;
            }

            this.PropertyInfo = property;
        }

        public int MinWidth { get; set; }

        public int MaxWidth { get; set; }

        public int ColumnIndex { get; set; }

        public string Name { get; set; }

        public bool HeaderIsBold { get; set; }

        public short HeaderFontColor { get; set; } = DefaultStyle.FontColor;

        public int HeaderFontSize { get; set; } = DefaultStyle.FontSize;

        public string HeaderFontName { get; set; } = DefaultStyle.FontName;

        public bool IsBold { get; set; }

        public short FontColor { get; set; } = DefaultStyle.FontColor;

        public int FontSize { get; set; } = DefaultStyle.FontSize;

        public string FontName { get; set; } = DefaultStyle.FontName;

        public string StringFormat { get; set; }

        public PropertyInfo PropertyInfo { get; set; }
    }
}
