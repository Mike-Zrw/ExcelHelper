using System;

namespace ExcelHelper.Attributes
{
    /// <summary>
    /// 非空判断
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, Inherited = false, AllowMultiple = false)]
    public class ColumnRequiredAttribute : Attribute
    {
        public ColumnRequiredAttribute()
        {
        }

        public ColumnRequiredAttribute(string emptyErrorMessage)
        {
            this.EmptyErrorMessage = emptyErrorMessage;
        }

        public string EmptyErrorMessage { get; set; } = "数据不可为空";
    }
}
