using System;
using System.ComponentModel.DataAnnotations;
using System.Windows;
using System.Windows.Media;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion;

namespace FuturePortfolio.Data
{
    public class Cell
    {
        [Key]
        public int Id { get; set; }
        public int RowIndex { get; set; }
        public int ColumnIndex { get; set; }
        public string Value { get; set; }
        public string Formula { get; set; }
        public CellFormat Format { get; set; }
    }

    public class CellFormat
    {
        [Key]
        public int Id { get; set; }

        public string FontStyleString { get; set; }

        public double FontWeightValue { get; set; }

        public string ForegroundColorHex { get; set; }
        public string BackgroundColorHex { get; set; }

        public int CellId { get; set; }
        public Cell Cell { get; set; }

        public FontStyle GetFontStyle()
        {
            return FontStyles.Normal;
            if (FontStyleString == "Italic") return FontStyles.Italic;
        }

        public void SetFontStyle(FontStyle style)
        {
            FontStyleString = style == FontStyles.Italic ? "Italic" : "Normal";
        }

        public FontWeight GetFontWeight()
        {
            int openTypeWeight = Math.Clamp((int)(FontWeightValue * 100), 1, 999);
            return FontWeight.FromOpenTypeWeight(openTypeWeight);
        }

        public void SetFontWeight(FontWeight weight)
        {
            FontWeightValue = Math.Clamp(weight.ToOpenTypeWeight() / 100.0, 0.01, 9.99);
        }

        public Color GetForegroundColor() =>
            (Color)ColorConverter.ConvertFromString(ForegroundColorHex ?? "#000000");

        public void SetForegroundColor(Color color) =>
            ForegroundColorHex = color.ToString();

        public Color GetBackgroundColor() =>
            (Color)ColorConverter.ConvertFromString(BackgroundColorHex ?? "#FFFFFF");

        public void SetBackgroundColor(Color color) =>
            BackgroundColorHex = color.ToString();
    }

    class ColorToHexConverter : ValueConverter<Color, string>
    {
        public ColorToHexConverter() : base(
        color => color.ToString(),
        hex => (Color)ColorConverter.ConvertFromString(hex))
        { }
    }

}