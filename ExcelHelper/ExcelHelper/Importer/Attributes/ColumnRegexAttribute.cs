using System;
using System.Text.RegularExpressions;

namespace ExcelHelper.Importer.Attributes
{
    /// <summary>
    /// 正则判断
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, Inherited = false, AllowMultiple = false)]
    public class ColumnRegexAttribute : Attribute
    {
        public ColumnRegexAttribute(string pattern)
        {
            this.Pattern = pattern;
        }

        public ColumnRegexAttribute(string pattern, string errorMessage)
            : this(pattern)
        {
            this.Pattern = pattern;
            this.ErrorMessage = errorMessage;
        }

        public ColumnRegexAttribute(string pattern, string errorMessage, RegexOptions regexOptions)
            : this(pattern, errorMessage)
        {
            this.RegexOptions = regexOptions;
        }

        public string Pattern { get; set; }

        public RegexOptions RegexOptions { get; set; } = default;

        public string ErrorMessage { get; set; } = "数据格式不正确";
    }
}
