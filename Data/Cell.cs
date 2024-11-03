using System;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Windows;
using System.Windows.Media;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion;
using Color = System.Windows.Media.Color;
using ColorConverter = System.Windows.Media.ColorConverter;
using FontStyle = System.Windows.FontStyle;


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

    public class WpfCell : INotifyPropertyChanged
    {
        private string _value;
        private string _formula;
        private WpfCellFormat _format;
        public int Id { get; set; }
        public int RowIndex { get; set; }
        public int ColumnIndex { get; set; }

        public string Value
        {
            get => _value;
            set
            {
                _value = value;
                OnPropertyChanged(nameof(Value));
            }
        }

        public string Formula
        {
            get => _formula;
            set
            {
                _formula = value;
                OnPropertyChanged(nameof(Formula));
            }
        }

        public WpfCellFormat Format
        {
            get => _format;
            set
            {
                _format = value;
                OnPropertyChanged(nameof(Format));
            }
        }

        public WpfCell()
        {
            Format = new WpfCellFormat();
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class WpfCellFormat : INotifyPropertyChanged
    {
        private FontStyle _fontStyle = FontStyles.Normal;
        private FontWeight _fontWeight = FontWeights.Normal;
        private Color _foregroundColor = Colors.Black;
        private Color _backgroundColor = Colors.White;

        public FontStyle FontStyle
        {
            get => _fontStyle;
            set
            {
                _fontStyle = value;
                OnPropertyChanged(nameof(FontStyle));
            }
        }

        public FontWeight FontWeight
        {
            get => _fontWeight;
            set
            {
                _fontWeight = value;
                OnPropertyChanged(nameof(FontWeight));
            }
        }

        public Color ForegroundColor
        {
            get => _foregroundColor;
            set
            {
                _foregroundColor = value;
                OnPropertyChanged(nameof(ForegroundColor));
            }
        }

        public Color BackgroundColor
        {
            get => _backgroundColor;
            set
            {
                _backgroundColor = value;
                OnPropertyChanged(nameof(BackgroundColor));
            }
        }

        public FontStyle GetFontStyle() => FontStyle;
        public FontWeight GetFontWeight() => FontWeight;
        public Color GetForegroundColor() => ForegroundColor;
        public Color GetBackgroundColor() => BackgroundColor;

        public void SetFontStyle(FontStyle style)
        {
            FontStyle = style;
        }

        public void SetFontWeight(FontWeight weight)
        {
            FontWeight = weight;
        }

        public void SetForegroundColor(Color color)
        {
            ForegroundColor = color;
        }

        public void SetBackgroundColor(Color color)
        {
            BackgroundColor = color;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public static class DataConverter
    {
        public static WpfCell ToWpfCell(Cell dbCell)
        {
            if (dbCell == null) return null;

            return new WpfCell
            {
                Id = dbCell.Id,
                RowIndex = dbCell.RowIndex,
                ColumnIndex = dbCell.ColumnIndex,
                Value = dbCell.Value ?? string.Empty,
                Formula = dbCell.Formula ?? string.Empty,
                Format = ToWpfCellFormat(dbCell.Format) ?? new WpfCellFormat()
            };
        }

        public static Cell ToDbCell(WpfCell wpfCell)
        {
            if (wpfCell == null) return null;

            var dbCell = new Cell
            {
                RowIndex = wpfCell.RowIndex,
                ColumnIndex = wpfCell.ColumnIndex,
                Value = wpfCell.Value,
                Formula = wpfCell.Formula
            };

            dbCell.Format = ToDbCellFormat(wpfCell.Format);
            return dbCell;
        }

        public static WpfCellFormat ToWpfCellFormat(CellFormat dbFormat)
        {
            if (dbFormat == null) return new WpfCellFormat();

            return new WpfCellFormat
            {
                FontStyle = dbFormat.GetFontStyle(),
                FontWeight = dbFormat.GetFontWeight(),
                ForegroundColor = dbFormat.GetForegroundColor(),
                BackgroundColor = dbFormat.GetBackgroundColor()
            };
        }

        public static CellFormat ToDbCellFormat(WpfCellFormat wpfFormat)
        {
            if (wpfFormat == null) return null;

            var dbFormat = new CellFormat();
            dbFormat.SetFontStyle(wpfFormat.FontStyle);
            dbFormat.SetFontWeight(wpfFormat.FontWeight);
            dbFormat.SetForegroundColor(wpfFormat.ForegroundColor);
            dbFormat.SetBackgroundColor(wpfFormat.BackgroundColor);
            return dbFormat;
        }
    }

    class ColorToHexConverter : ValueConverter<Color, string>
    {
        public ColorToHexConverter() : base(
        color => color.ToString(),
        hex => (Color)ColorConverter.ConvertFromString(hex))
        { }
    }


}