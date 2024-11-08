using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using FuturePortfolio.Core;
using Microsoft.Extensions.DependencyInjection;

namespace FuturePortfolio
{
    public partial class MainWindow : Window
    {
        private readonly MainViewModel _viewModel;
        private readonly IServiceScope _scope;
        private bool _isClosing;

        public MainWindow()
        {
            InitializeComponent();

            if (App.Host != null)
            {
                _scope = App.Host.Services.CreateScope();
                _viewModel = _scope.ServiceProvider.GetRequiredService<MainViewModel>();
                DataContext = _viewModel;
            }
            else
            {
                throw new InvalidOperationException("Application host is not initialized.");
            }

            Loaded += MainWindow_Loaded;
            Closing += MainWindow_Closing;
        }

        private void SpreadsheetGrid_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            var column = new DataGridTextColumn();
            var header = ((PropertyDescriptor)e.PropertyDescriptor).DisplayName;

            if (header.StartsWith("Column"))
            {
                int columnIndex = int.Parse(header.Replace("Column", ""));
                column.Header = new CellPosition(0, columnIndex).ToColumnName();
                column.Binding = new Binding($"Cells[{columnIndex}].DisplayValue")
                {
                    Mode = BindingMode.TwoWay,
                    UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
                };
                e.Column = column;
            }
            else
            {
                e.Cancel = true;
            }
        }

        private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                await _viewModel.InitializeAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Failed to load data: {ex.Message}",
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private async void MainWindow_Closing(object sender, CancelEventArgs e)
        {
            if (_isClosing) return;

            try
            {
                e.Cancel = true;
                _isClosing = true;

                var result = MessageBox.Show(
                    "Do you want to save changes before closing?",
                    "Save Changes",
                    MessageBoxButton.YesNoCancel,
                    MessageBoxImage.Question);

                if (result == MessageBoxResult.Cancel)
                {
                    _isClosing = false;
                    return;
                }

                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        await _viewModel.Save();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(
                            $"Failed to save changes: {ex.Message}",
                            "Save Error",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error);
                    }
                }

                _scope?.Dispose();
                Application.Current.Shutdown();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error while closing: {ex.Message}",
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                Application.Current.Shutdown();
            }
        }

        private async void SaveAndClose()
        {
            try
            {
                await _viewModel.Save();
                Close();
            }
            catch (Exception ex)
            {
                HandleClosingError(ex, new CancelEventArgs());
            }
        }

        private void HandleClosingError(Exception ex, CancelEventArgs e)
        {
            var result = MessageBox.Show(
                $"Failed to save changes: {ex.Message}\n\nDo you want to close without saving?",
                "Save Error",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            e.Cancel = (result == MessageBoxResult.No);

            if (!e.Cancel)
            {
                Close();
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            _scope?.Dispose();
        }

        private void SpreadsheetGrid_LoadingRow(object? sender, DataGridRowEventArgs e)
        {
            e.Row.Header = (e.Row.GetIndex() + 1).ToString();
        }

        private void SpreadsheetGrid_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
        {
            if (SpreadsheetGrid.SelectedCells.Count > 0)
            {
                var selectedCell = SpreadsheetGrid.SelectedCells[0];
                var row = (ObservableCollection<CellViewModel>)selectedCell.Item;
                var columnIndex = selectedCell.Column.DisplayIndex;

                if (columnIndex >= 0 && columnIndex < row.Count)
                {
                    var cell = row[columnIndex];
                    _viewModel.OnCellSelected(cell);
                }
            }
        }

        private async void SpreadsheetGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.EditAction != DataGridEditAction.Commit)
                return;

            try
            {
                if (e.EditingElement is TextBox textBox)
                {
                    var row = (ObservableCollection<CellViewModel>)e.Row.Item;
                    var columnIndex = e.Column.DisplayIndex;

                    if (columnIndex >= 0 && columnIndex < row.Count)
                    {
                        var position = new CellPosition(e.Row.GetIndex(), columnIndex);

                        e.Cancel = true;
                        await _viewModel.UpdateCellValueAsync(position, textBox.Text);
                        e.Cancel = false;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Failed to update cell: {ex.Message}",
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }
    }
}