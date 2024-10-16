using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Media;
using System.IO;
using System.Text.Json;
using System.Windows.Controls;
using System.Data;

namespace FuturePortfolio
{
    public partial class MainWindow : Window
    {
        private SpreadsheetData _data;

        public MainWindow()
        {
            InitializeComponent();
            _data = new SpreadsheetData(100, 3);
            ExcelLikeGrid.ItemsSource = _data;

            ExcelLikeGrid.LoadingRow += ExcelLikeGrid_LoadingRow;
        }

        private void ExcelLikeGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = (e.Row.GetIndex() + 1).ToString();
        }

        public void SaveSpreadsheet(string filePath)
        {
            var jsonString = JsonSerializer.Serialize(_data);
            File.WriteAllText(filePath, jsonString);
        }

        public void LoadSpreadsheet(string filePath)
        {
            var jsonString = File.ReadAllText(filePath);
            _data = JsonSerializer.Deserialize<SpreadsheetData>(jsonString);
            ExcelLikeGrid.ItemsSource = _data;
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
                var cell = (e.Row.Item as ObservableCollection<Cell>)[e.Column.DisplayIndex];
                var editedValue = (e.EditingElement as TextBox).Text;

                if (editedValue.StartsWith("="))
                {
                    cell.Formula = editedValue;
                    cell.Value = FormulaCalculator.CalculateFormula(editedValue);
                }
                else
                {
                    cell.Value = editedValue;
                    cell.Formula = null;
                }
            }
        }
    }
}

public class Cell : INotifyPropertyChanged
{
    private string _value;
    private string _formula;
    private CellFormat _format;

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

    public event PropertyChangedEventHandler PropertyChanged;

    protected virtual void OnPropertyChanged(string propertyName)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}

public class CellFormat
{
    public FontStyle FontStyle { get; set; }
    public FontWeight FontWeight { get; set; }
    public Color ForegroundColor { get; set; }
    public Color BackgroundColor { get; set; }
}

public class SpreadsheetData : ObservableCollection<ObservableCollection<Cell>>
{
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