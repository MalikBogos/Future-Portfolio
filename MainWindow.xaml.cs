
using System.Windows;
using System.IO;
using System.Windows.Controls;
using System.Data;
using System.Text.Json;
using FuturePortfolio.Data;
using TextBox = System.Windows.Controls.TextBox;
using Brushes = System.Windows.Media.Brushes;
using static FuturePortfolio.Data.SpreadSheetContext;


namespace FuturePortfolio
{

    public partial class MainWindow : Window
    {
        private SpreadsheetDataWithDb _data;
        private readonly SpreadSheetContext _context;
        private WpfCell? _selectedCell;
        private const string SaveFilePath = "spreadsheet.json";


        public MainWindow()
        {
            InitializeComponent();
            _context = new SpreadSheetContext();
            _data = new SpreadsheetDataWithDb(_context);

            this.DataContext = _data;
            GenerateColumns();
            ExcelLikeGrid.ItemsSource = _data;

            ExcelLikeGrid.LoadingRow += ExcelLikeGrid_LoadingRow;
            this.Closing += MainWindow_Closing;
        }

        

        private void ExcelLikeGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var cell = ExcelLikeGrid.CurrentCell;
            if (cell != null && cell.Item != null)
            {
                var row = ExcelLikeGrid.Items.IndexOf(cell.Item);
                var column = cell.Column.DisplayIndex;
                _selectedCell = _data[row][column];

            }
        }

        private void AddRowButton_Click(object sender, RoutedEventArgs e)
        {
            _data.AddRow();
            UpdateRowHeaders();
        }

        private void AddColumnButton_Click(object sender, RoutedEventArgs e)
        {
            _data.AddColumn();
            GenerateColumns();
        }

        private void RemoveRowButton_Click(object sender, RoutedEventArgs e)
        {
            _data.RemoveLastRow();
            UpdateRowHeaders();
        }

        private void RemoveColumnButton_Click(object sender, RoutedEventArgs e)
        {
            _data.RemoveLastColumn();
            GenerateColumns();
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
            _data.SaveToDatabase();
            _context.Dispose();
        }

        private void UpdateRowHeaders()
        {
            for (int i = 0; i < _data.Count; i++)
            {
                var rowContainer = ExcelLikeGrid.ItemContainerGenerator.ContainerFromIndex(i) as DataGridRow;
                if (rowContainer != null)
                {
                    rowContainer.Header = (i + 1).ToString();
                }
            }
        }

        private void SaveSpreadsheet()
        {
            var options = new JsonSerializerOptions
            {
                WriteIndented = true
            };
            var jsonString = JsonSerializer.Serialize(_data, options);
            File.WriteAllText(SaveFilePath, jsonString);
        }

        private void ExcelLikeGrid_LoadingRow(object? sender, DataGridRowEventArgs e)
        {
            e.Row.Header = (e.Row.GetIndex() + 1).ToString();
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
