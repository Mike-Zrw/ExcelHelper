using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelHelper.Attributes
{
    /// <summary>
    /// 行合并，根据ParentKey合并
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, Inherited = false, AllowMultiple = false)]
    public class RowMergedAttribute : Attribute
    {
    }
}
