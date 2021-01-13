using ExcelHelper.Exporter.Enums;
using System;

namespace ExcelHelper.Exporter.Attributes
{
    public class StyleAttribute : Attribute
    {
        public StyleAttribute()
        {
            Style = new CellStyle();
        }
        public IBaseStyle Style { get; set; }
        public bool IsBold { get => Style.IsBold; set => Style.IsBold = value; }
        public bool WrapText { get => Style.WrapText; set => Style.WrapText = value; }
        public short FontColor { get => Style.FontColor; set => Style.FontColor = value; }
        public short FillForegroundColor { get => Style.FillForegroundColor; set => Style.FillForegroundColor = value; }
        public int FontSize { get => Style.FontSize; set => Style.FontSize = value; }
        public string FontName { get => Style.FontName; set => Style.FontName = value; }
        public HorizontalAlignEnum HorizontalAlign { get => Style.HorizontalAlign; set => Style.HorizontalAlign = value; }
        public VerticalAlignmentEnum VerticalAlign { get => Style.VerticalAlign; set => Style.VerticalAlign = value; }
    }
}