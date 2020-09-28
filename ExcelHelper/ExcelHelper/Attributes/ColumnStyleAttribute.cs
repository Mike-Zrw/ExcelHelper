using System;

namespace ExcelHelper.Excel.Attributes
{
    /// <summary>
    /// 内容样式
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, Inherited = false, AllowMultiple = false)]
    public class ColumnStyleAttribute : StyleAttribute
    {
        public ColumnStyleAttribute(bool isBold)
        {
            this.IsBold = isBold;
        }

        public ColumnStyleAttribute()
        {
        }
    }
}
