using System;

namespace ExcelHelper.Exporter.Attributes
{
    /// <summary>
    /// 表头样式
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, Inherited = false, AllowMultiple = false)]
    public class HeaderStyleAttribute : StyleAttribute
    {
        public HeaderStyleAttribute(bool isBold)
        {
            this.Style.IsBold = isBold;
        }

        public HeaderStyleAttribute()
        {
        }
    }
}