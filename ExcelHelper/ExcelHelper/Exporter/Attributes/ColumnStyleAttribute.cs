using System;

namespace ExcelHelper.Exporter.Attributes
{
    /// <summary>
    /// 内容样式
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, Inherited = false, AllowMultiple = false)]
    public class ColumnStyleAttribute : StyleAttribute
    {
        public ColumnStyleAttribute(bool isBold)
        {
            this.Style.IsBold = isBold;
        }

        public ColumnStyleAttribute()
        {
        }
    }
}