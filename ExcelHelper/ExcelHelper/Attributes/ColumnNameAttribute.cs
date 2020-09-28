using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelHelper.Excel.Attributes
{
    /// <summary>
    /// 列名
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, Inherited = false, AllowMultiple = false)]
    public class ColumnNameAttribute : Attribute
    {
        public ColumnNameAttribute(string name)
        {
            this.Name = name;
        }

        public string Name { get; set; }
    }
}
