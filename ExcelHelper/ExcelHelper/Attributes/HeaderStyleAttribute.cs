using System;

namespace ExcelHelper.Excel.Attributes
{
    /// <summary>
    /// 表头样式
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, Inherited = false, AllowMultiple = false)]
    public class HeaderStyleAttribute : StyleAttribute
    {
        public HeaderStyleAttribute(bool isBold)
        {
            this.IsBold = isBold;
        }

        public HeaderStyleAttribute()
        {
        }
    }
}
