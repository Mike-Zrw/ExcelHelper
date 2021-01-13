using ExcelHelper.Exporter.Enums;

namespace ExcelHelper.Exporter
{
    public class CellStyle : IBaseStyle
    {
        /// <summary>
        /// 加粗
        /// </summary>
        public bool IsBold { get; set; } = false;


        /// <summary>
        /// 自动换行
        /// </summary>
        public bool WrapText { get; set; } = true;

        public short FontColor { get; set; } = DefaultStyle.FontColor;

        public int FontSize { get; set; } = DefaultStyle.FontSize;

        /// <summary>
        /// 字体名称
        /// </summary>
        public string FontName { get; set; } = DefaultStyle.FontName;

        public HorizontalAlignEnum HorizontalAlign { get; set; } = DefaultStyle.HorizontalAlign;

        public VerticalAlignmentEnum VerticalAlign { get; set; } = DefaultStyle.VerticalAlign;
        public short FillForegroundColor { get; set; }
    }
}
