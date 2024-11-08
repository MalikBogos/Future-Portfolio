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
            if (e.PropertyName.StartsWith("Item[") && e.PropertyName.EndsWith("]"))
            {
                var column = new DataGridTextColumn();

                // Extract index from property name (format is "Item[X]")
                string indexStr = e.PropertyName.Substring(5, e.PropertyName.Length - 6);
                if (int.TryParse(indexStr, out int columnIndex))
                {
                    // Convert to Excel-style column header (A, B, C, etc.)
                    var position = new CellPosition(0, columnIndex);
                    column.Header = position.ToColumnName();

                    // Set the binding to DisplayValue of the CellViewModel
                    column.Binding = new Binding($"DisplayValue");

                    // Apply styles
                    column.ElementStyle = new Style(typeof(TextBlock))
                    {
                        Setters = { new Setter(TextBlock.PaddingProperty, new Thickness(4, 2, 4, 2)) }
                    };

                    column.EditingElementStyle = new Style(typeof(TextBox))
                    {
                        Setters = {
                            new Setter(TextBox.PaddingProperty, new Thickness(4, 2, 4, 2)),
                            new Setter(TextBox.BorderThicknessProperty, new Thickness(0))
                        }
                    };

                    e.Column = column;
                }
                else
                {
                    e.Cancel = true;
                }
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
                        await _viewModel.SaveAsync();
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
                await _viewModel.SaveAsync();
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
                    var cell = row[e.Column.DisplayIndex];
                    var position = new CellPosition(e.Row.GetIndex(), e.Column.DisplayIndex);

                    e.Cancel = true;
                    await _viewModel.UpdateCellValueAsync(position, textBox.Text);
                    e.Cancel = false;
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