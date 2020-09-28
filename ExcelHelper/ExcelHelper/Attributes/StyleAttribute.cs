using System;

namespace ExcelHelper.Excel.Attributes
{
    public class StyleAttribute : Attribute
    {
        /// <summary>
        /// 加粗
        /// </summary>
        public bool IsBold { get; set; }

        public short FontColor { get; set; } = DefaultStyle.FontColor;

        public int FontSize { get; set; } = DefaultStyle.FontSize;

        /// <summary>
        /// 字体名称
        /// </summary>
        public string FontName { get; set; } = DefaultStyle.FontName;
    }
}
