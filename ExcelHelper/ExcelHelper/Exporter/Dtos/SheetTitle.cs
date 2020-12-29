using ExcelHelper.Exporter.Enums;

namespace ExcelHelper.Exporter.Dtos
{
    /// <summary>
    /// 导出的第一行标题
    /// </summary>
    public class SheetTitle
    {
        public SheetTitle()
        {
        }

        public SheetTitle(string title)
            : this()
        {
            this.Title = title;
        }

        public SheetTitle(string title, bool isBold, int fontSize = DefaultStyle.FontSize, short fontColor = DefaultStyle.FontColor, HorizontalAlignEnum horizontalAlign = HorizontalAlignEnum.Center, string fontName = DefaultStyle.FontName)
            : this(title)
        {
            this.IsBold = isBold;
            this.FontSize = fontSize;
            this.FontColor = fontColor;
            this.HorizontalAlign = horizontalAlign;
            this.FontName = fontName;
        }

        public string Title { get; set; }

        public bool IsBold { get; set; }

        public int FontSize { get; set; } = DefaultStyle.FontSize;

        public short FontColor { get; set; } = DefaultStyle.FontColor;

        /// <summary>
        /// 字体
        /// </summary>
        public string FontName { get; set; } = DefaultStyle.FontName;

        public HorizontalAlignEnum HorizontalAlign { get; set; } = HorizontalAlignEnum.Center;
    }
}
