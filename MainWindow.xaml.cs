using System.Windows;
using System.Windows.Controls;
using System.Data;
using FuturePortfolio.Data;
using Microsoft.Win32;
using System.Text.RegularExpressions;
using static FuturePortfolio.Data.SpreadSheetContext;
using System.Windows.Media;

namespace FuturePortfolio
{
    public partial class MainWindow : Window
    {
        private readonly SpreadsheetDataWithDb _data;
        private readonly SpreadSheetContext _context;
        private WpfCell? _selectedCell;
        private static readonly Regex FormulaPattern = new(@"^=.*", RegexOptions.Compiled);
        private const int DefaultColumnCount = 5;

        public MainWindow()
        {
            InitializeComponent();
            _context = new SpreadSheetContext();
            _data = new SpreadsheetDataWithDb(_context);

            if (_data.Count == 0)
            {
                _data.AddRow();
            }

            DataContext = _data;
            InitializeGrid();

            ExcelLikeGrid.LoadingRow += ExcelLikeGrid_LoadingRow;
            Closing += MainWindow_Closing;
        }

        private void InitializeGrid()
        {
            GenerateColumns();
            ExcelLikeGrid.ItemsSource = _data;
        }

        private void ExcelLikeGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var cell = ExcelLikeGrid.CurrentCell;
            if (cell.Column != null && cell.Item != null)
            {
                var row = ExcelLikeGrid.Items.IndexOf(cell.Item);
                var column = cell.Column.DisplayIndex;
                _selectedCell = _data[row][column];
            }
        }

        private void GenerateColumns()
        {
            ExcelLikeGrid.Columns.Clear();

            int columnCount = _data.FirstOrDefault()?.Count ?? DefaultColumnCount;

            if (columnCount == 0)
            {
                columnCount = DefaultColumnCount;
                for (int i = 0; i < columnCount; i++)
                {
                    _data[0].Add(new WpfCell
                    {
                        RowIndex = 0,
                        ColumnIndex = i
                    });
                }
            }

            for (int i = 0; i < columnCount; i++)
            {
                ExcelLikeGrid.Columns.Add(CreateColumn(i));
            }
        }

        private static DataGridTextColumn CreateColumn(int index) =>
            new()
            {
                Header = GetColumnHeader(index),
                Width = 100,
                Binding = new System.Windows.Data.Binding($"[{index}].Value")
                {
                    UpdateSourceTrigger = System.Windows.Data.UpdateSourceTrigger.PropertyChanged
                },
                EditingElementStyle = CreateEditingStyle()
            };

        private static string GetColumnHeader(int index)
        {
            const int alphabetLength = 26;
            string header = string.Empty;

            do
            {
                header = (char)('A' + (index % alphabetLength)) + header;
                index = (index / alphabetLength) - 1;
            } while (index >= 0);

            return header;
        }

        private static Style CreateEditingStyle() =>
            new(typeof(TextBox))
            {
                Setters =
                {
                    new Setter(TextBox.BorderThicknessProperty, new Thickness(0)),
                    new Setter(TextBox.BackgroundProperty, Brushes.White),
                    new Setter(TextBox.PaddingProperty, new Thickness(2))
                }
            };

        private async void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                await _data.SaveToDatabaseAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving data: {ex.Message}", "Save Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                _data.Dispose();
            }
        }

        private void ExcelLikeGrid_LoadingRow(object? sender, DataGridRowEventArgs e)
        {
            e.Row.Header = (e.Row.GetIndex() + 1).ToString();
        }

        private void ExcelLikeGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.EditAction == DataGridEditAction.Commit && e.EditingElement is TextBox editedTextBox)
            {
                UpdateCellValue(e.Row.GetIndex(), e.Column.DisplayIndex, editedTextBox.Text);
            }
        }

        private void UpdateCellValue(int row, int column, string newValue)
        {
            var cell = _data[row][column];

            if (FormulaPattern.IsMatch(newValue))
            {
                cell.Formula = newValue;
                cell.Value = EvaluateFormula(newValue[1..]);
            }
            else
            {
                cell.Formula = null;
                cell.Value = newValue;
            }
        }

        private static string EvaluateFormula(string formula)
        {
            try
            {
                var dt = new DataTable();
                var result = dt.Compute(formula, string.Empty);
                return result.ToString() ?? string.Empty;
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }

        private void AddRowButton_Click(object sender, RoutedEventArgs e) => _data.AddRow();
        private void AddColumnButton_Click(object sender, RoutedEventArgs e)
        {
            _data.AddColumn();
            GenerateColumns();
        }
        private void RemoveRowButton_Click(object sender, RoutedEventArgs e) => _data.RemoveLastRow();
        private void RemoveColumnButton_Click(object sender, RoutedEventArgs e)
        {
            _data.RemoveLastColumn();
            GenerateColumns();
        }
    }
}