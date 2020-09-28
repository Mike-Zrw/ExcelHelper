using System;

namespace ExcelHelper.Attributes
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

        public HorizontalAlignEnum HorizontalAlign { get; set; } = HorizontalAlignEnum.Left;

        public VerticalAlignmentEnum VerticalAlign { get; set; } = VerticalAlignmentEnum.Center;
    }
}
