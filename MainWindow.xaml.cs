using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Media;
using System.IO;
using System.Windows.Controls;
using System.Data;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace FuturePortfolio
{
    public class CellTemplateSelector : DataTemplateSelector
    {
        public DataTemplate DefaultTemplate { get; set; }
        public DataTemplate EditingTemplate { get; set; }

        public override DataTemplate SelectTemplate(object item, DependencyObject container)
        {
            var cell = container as DataGridCell;
            return cell != null && cell.IsEditing ? EditingTemplate : DefaultTemplate;
        }
    }

    public partial class MainWindow : Window
    {
        private SpreadsheetData _data;
        private const string SaveFilePath = "spreadsheet_data.json";
        private Cell _selectedCell;

        public MainWindow()
        {
            InitializeComponent();
            LoadSpreadsheet();
            this.DataContext = _data;

            GenerateColumns();
            ExcelLikeGrid.ItemsSource = _data;

            ExcelLikeGrid.LoadingRow += ExcelLikeGrid_LoadingRow;
            this.Closing += MainWindow_Closing;

            ExcelLikeGrid.SelectionChanged += (s, e) =>
            {
                MessageBox.Show($"Selection changed. Selected cells: {ExcelLikeGrid.SelectedCells.Count}");
            };
        }

        private void GenerateColumns()
        {
            ExcelLikeGrid.Columns.Clear();
            for (int i = 0; i < _data[0].Count; i++)
            {
                var column = new DataGridTextColumn
                {
                    Header = ((char)('A' + i)).ToString(),
                    Width = 100,
                    Binding = new System.Windows.Data.Binding($"[{i}].Value")
                    {
                        UpdateSourceTrigger = System.Windows.Data.UpdateSourceTrigger.PropertyChanged
                    },
                    EditingElementStyle = new Style(typeof(TextBox))
                    {
                        Setters =
                {
                    new Setter(TextBox.BorderThicknessProperty, new Thickness(0)),
                    new Setter(TextBox.BackgroundProperty, Brushes.White),
                    new Setter(TextBox.PaddingProperty, new Thickness(2))
                }
                    }
                };

                ExcelLikeGrid.Columns.Add(column);
            }
        }

        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            SaveSpreadsheet();
        }

        private void SaveSpreadsheet()
        {
            var jsonString = JsonSerializer.Serialize(_data);
            File.WriteAllText(SaveFilePath, jsonString);
        }

        private void LoadSpreadsheet()
        {
            if (File.Exists(SaveFilePath))
            {
                var jsonString = File.ReadAllText(SaveFilePath);
                _data = JsonSerializer.Deserialize<SpreadsheetData>(jsonString);
            }
            else
            {
                _data = new SpreadsheetData(20, 5);
            }
        }

        private void ExcelLikeGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = (e.Row.GetIndex() + 1).ToString();
        }

        public void ApplyFormatToCell(int row, int column, CellFormat format)
        {
            var cell = _data.GetCell(row, column);
            cell.Format = format;
        }

        private void ExcelLikeGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.EditAction == DataGridEditAction.Commit)
            {
                var column = e.Column.DisplayIndex;
                var row = e.Row.GetIndex();
                var editedTextBox = e.EditingElement as TextBox;

                if (editedTextBox != null)
                {
                    string newValue = editedTextBox.Text;
                    UpdateCellValue(row, column, newValue);
                }
            }
        }

        private void ExcelLikeGrid_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
        {
            if (ExcelLikeGrid.SelectedCells.Count > 0)
            {
                var cellInfo = ExcelLikeGrid.SelectedCells[0];
                var row = cellInfo.Item as ObservableCollection<Cell>;
                var column = cellInfo.Column.DisplayIndex;
                _selectedCell = row[column];
            }
            else
            {
                _selectedCell = null;
            }
        }

        private void BoldButton_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedCell != null)
            {
                _selectedCell.Format ??= new CellFormat();
                _selectedCell.Format.FontWeight = _selectedCell.Format.FontWeight == FontWeights.Bold ? FontWeights.Normal : FontWeights.Bold;
                RefreshCell();
            }
        }

        private void ItalicButton_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedCell != null)
            {
                _selectedCell.Format ??= new CellFormat();
                _selectedCell.Format.FontStyle = _selectedCell.Format.FontStyle == FontStyles.Italic ? FontStyles.Normal : FontStyles.Italic;
                RefreshCell();
            }
        }

        private void ColorButton_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedCell != null)
            {
                _selectedCell.Format ??= new CellFormat();
                _selectedCell.Format.ForegroundColor = _selectedCell.Format.ForegroundColor == Colors.Red ? Colors.Black : Colors.Red;
                RefreshCell();
            }
        }

        private void RefreshCell()
        {
            ExcelLikeGrid.Items.Refresh();
        }

        private void UpdateCellValue(int row, int column, string newValue)
        {
            var cell = _data[row][column];

            if (newValue.StartsWith("="))
            {
                cell.Formula = newValue;
                cell.Value = EvaluateFormula(newValue.Substring(1));
            }
            else
            {
                cell.Formula = null;
                cell.Value = newValue;
            }
        }

        private string EvaluateFormula(string formula)
        {
            try
            {
                DataTable dt = new DataTable();
                var result = dt.Compute(formula, "");
                return result.ToString();
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }
    }
}

public class Cell : INotifyPropertyChanged
{
    private string _value;
    private string _formula;
    private CellFormat _format;

    [JsonPropertyName("Value")]
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

    public CellFormat Format
    {
        get => _format;
        set
        {
            _format = value;
            OnPropertyChanged(nameof(Format));
        }
    }

    public Cell()
    {
        Format = new CellFormat();
    }

    public event PropertyChangedEventHandler PropertyChanged;

    protected virtual void OnPropertyChanged(string propertyName)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}

public class CellFormat : INotifyPropertyChanged
{
    public FontStyle FontStyle { get; set; } = FontStyles.Normal;
    public FontWeight FontWeight { get; set; } = FontWeights.Normal;
    public Color ForegroundColor { get; set; } = Colors.Black;
    public Color BackgroundColor { get; set; } = Colors.White;

    public event PropertyChangedEventHandler PropertyChanged;

    protected virtual void OnPropertyChanged(string propertyName)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}

public class SpreadsheetData : ObservableCollection<ObservableCollection<Cell>>
{
    [JsonConstructor]
    public SpreadsheetData() { }

    public SpreadsheetData(int rows, int columns)
    {
        for (int i = 0; i < rows; i++)
        {
            var row = new ObservableCollection<Cell>();
            for (int j = 0; j < columns; j++)
            {
                row.Add(new Cell());
            }
            Add(row);
        }
    }

    public Cell GetCell(int row, int column)
    {
        return this[row][column];
    }

    public void SetCellValue(int row, int column, string value)
    {
        this[row][column].Value = value;
    }
}

public static class FormulaCalculator
{
    public static string CalculateFormula(string formula)
    {
        if (string.IsNullOrWhiteSpace(formula) || !formula.StartsWith("="))
            return formula;

        formula = formula.Substring(1);

        try
        {
            DataTable dt = new DataTable();
            var result = dt.Compute(formula, "");
            return result.ToString();
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }
}