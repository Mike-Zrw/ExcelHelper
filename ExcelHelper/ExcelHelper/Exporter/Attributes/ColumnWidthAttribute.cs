using System;

namespace ExcelHelper.Exporter.Attributes
{
    /// <summary>
    /// 列宽
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, Inherited = false, AllowMultiple = false)]
    public class ColumnWidthAttribute : Attribute
    {
        public ColumnWidthAttribute(int minWidth, int maxWidth)
        {
            this.MinWidth = minWidth;
            this.MaxWidth = maxWidth;
        }

        public int MinWidth { get; set; }

        public int MaxWidth { get; set; }
    }
}
