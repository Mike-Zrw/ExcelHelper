using System;

namespace ExcelHelper.Exporter.Attributes
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, Inherited = false, AllowMultiple = false)]
    public class StringFormatterAttribute : Attribute
    {
        public StringFormatterAttribute(string format)
        {
            this.Format = format;
        }

        public string Format { get; set; }
    }
}
