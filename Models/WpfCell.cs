using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Media;
using System.Linq;
using FuturePortfolio.Data;
using Microsoft.EntityFrameworkCore;
using Color = System.Windows.Media.Color;
using FontStyle = System.Windows.FontStyle;

namespace FuturePortfolio.Models
{
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

    // WPF CellFormat Model
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

        // Getter methods
        public FontStyle GetFontStyle() => FontStyle;
        public FontWeight GetFontWeight() => FontWeight;
        public Color GetForegroundColor() => ForegroundColor;
        public Color GetBackgroundColor() => BackgroundColor;

        // Setter methods
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

    // Data Transfer Objects (DTO) Converter
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

    // SpreadsheetData with database support
    public class SpreadsheetDataWithDb : ObservableCollection<ObservableCollection<WpfCell>>
    {
        private readonly SpreadSheetContext _context;

        public SpreadsheetDataWithDb(SpreadSheetContext context)
        {
            _context = context;
            LoadFromDatabase();
        }

        private void LoadFromDatabase()
        {
            // Load cells with explicit format inclusion
            var cells = _context.Cells
                .Include(c => c.Format)  // Ensure format is loaded
                .AsNoTracking()  // Improve performance for read-only data
                .OrderBy(c => c.RowIndex)
                .ThenBy(c => c.ColumnIndex)
                .ToList();

            // Find max row and column indices
            int maxRow = cells.Any() ? cells.Max(c => c.RowIndex) : 0;
            int maxCol = cells.Any() ? cells.Max(c => c.ColumnIndex) : 0;

            // Create matrix with default cells
            for (int i = 0; i <= maxRow; i++)
            {
                var wpfRow = new ObservableCollection<WpfCell>();
                for (int j = 0; j <= maxCol; j++)
                {
                    // Find existing cell or create new one
                    var dbCell = cells.FirstOrDefault(c => c.RowIndex == i && c.ColumnIndex == j);
                    if (dbCell == null)
                    {
                        dbCell = new Cell
                        {
                            RowIndex = i,
                            ColumnIndex = j,
                            Format = new CellFormat()  // Create default format
                        };
                    }
                    wpfRow.Add(DataConverter.ToWpfCell(dbCell));
                }
                Add(wpfRow);
            }
        }

        public void SaveToDatabase()
        {
            _context.CellFormats.RemoveRange(_context.CellFormats);
            _context.Cells.RemoveRange(_context.Cells);
            _context.SaveChanges(); 
            try
            {
                // Remove existing cells and their formats
                _context.CellFormats.RemoveRange(_context.CellFormats);
                _context.Cells.RemoveRange(_context.Cells);
                _context.SaveChanges();

                // Add current cells
                for (int rowIndex = 0; rowIndex < Count; rowIndex++)
                {
                    for (int colIndex = 0; colIndex < this[rowIndex].Count; colIndex++)
                    {
                        var wpfCell = this[rowIndex][colIndex];
                        if (!string.IsNullOrEmpty(wpfCell.Value) || !string.IsNullOrEmpty(wpfCell.Formula))
                        {
                            wpfCell.RowIndex = rowIndex;
                            wpfCell.ColumnIndex = colIndex;
                            var dbCell = DataConverter.ToDbCell(wpfCell);
                            _context.Cells.Add(dbCell);
                        }
                    }
                }
                _context.SaveChanges();
            }
            catch (Exception ex)
            {
                // Handle the exception as needed, logging or rethrowing
                throw;
            }
        }

        public void AddRow()
        {
            var newRow = new ObservableCollection<WpfCell>();
            int columnCount = this.FirstOrDefault()?.Count ?? 0;

            for (int j = 0; j < columnCount; j++)
            {
                newRow.Add(new WpfCell
                {
                    RowIndex = Count,
                    ColumnIndex = j,
                    Format = new WpfCellFormat()  // Ensure format is initialized
                });
            }
            Add(newRow);
        }

        public void AddColumn()
        {
            int newColumnIndex = this.FirstOrDefault()?.Count ?? 0;

            foreach (var row in this)
            {
                row.Add(new WpfCell
                {
                    RowIndex = this.IndexOf(row),
                    ColumnIndex = newColumnIndex,
                    Format = new WpfCellFormat()  // Ensure format is initialized
                });
            }
        }

        public void RemoveLastRow()
        {
            if (Count > 1)
            {
                RemoveAt(Count - 1);
            }
        }

        public void RemoveLastColumn()
        {
            if (this.FirstOrDefault()?.Count > 1)
            {
                foreach (var row in this)
                {
                    row.RemoveAt(row.Count - 1);
                }
            }
        }
    }
}