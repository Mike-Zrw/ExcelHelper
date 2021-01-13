using ExcelHelper.Exporter.Enums;

namespace ExcelHelper.Exporter
{
    public interface IBaseStyle
    {
        /// <summary>
        /// 加粗
        /// </summary>
        bool IsBold { get; set; }
        /// <summary>
        /// 自动换行
        /// </summary>
        bool WrapText { get; set; }
        short FontColor { get; set; }

        int FontSize { get; set; }

        /// <summary>
        /// 字体名称
        /// </summary>
        string FontName { get; set; }

        short FillForegroundColor { get; set; }

        HorizontalAlignEnum HorizontalAlign { get; set; }

        VerticalAlignmentEnum VerticalAlign { get; set; }
    }
}
