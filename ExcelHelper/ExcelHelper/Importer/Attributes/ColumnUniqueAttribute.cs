using System;

namespace ExcelHelper.Importer.Attributes
{
    /// <summary>
    /// 唯一验证
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, Inherited = false, AllowMultiple = false)]
    public class ColumnUniqueAttribute : Attribute
    {
    }
}
