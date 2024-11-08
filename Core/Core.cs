using System.ComponentModel;
using System.Collections.ObjectModel;
using Microsoft.EntityFrameworkCore;
using System.Data;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.Hosting;
using System.Text.RegularExpressions;
using Microsoft.Extensions.DependencyInjection;
using FuturePortfolio.Core;
using System.Runtime.InteropServices;
using System.IO;
using System.Windows.Controls;
using Microsoft.Win32;
using System.Windows;
using FuturePortfolio;
using System;

namespace FuturePortfolio.Core
{
    // Core Models
    public record CellPosition(int Row, int Column)
    {
        public static CellPosition FromIndex(int index, int columnCount) =>
            new(index / columnCount, index % columnCount);

        public int ToIndex(int columnCount) => Row * columnCount + Column;

        public string ToColumnName()
        {
            string columnName = string.Empty;
            int dividend = Column + 1;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }
    }

    public record CellFormatting
    {
        public bool IsBold { get; init; }
        public bool IsItalic { get; init; }
        public bool IsUnderlined { get; init; }

        public static CellFormatting Default => new();

        public CellFormatting WithBold(bool bold) => this with { IsBold = bold };
        public CellFormatting WithItalic(bool italic) => this with { IsItalic = italic };
        public CellFormatting WithUnderline(bool underline) => this with { IsUnderlined = underline };
    }

    public record CellValue
    {
        public string? DisplayValue { get; init; }
        public string? Formula { get; init; }
        public CellValueType Type { get; init; }
        public CellFormatting Formatting { get; init; }

        private CellValue(string? displayValue, string? formula, CellValueType type, CellFormatting? formatting = null)
        {
            DisplayValue = displayValue;
            Formula = formula;
            Type = type;
            Formatting = formatting ?? CellFormatting.Default;
        }

        public static CellValue Empty => new(null, null, CellValueType.Empty);
        public static CellValue FromText(string text, CellFormatting? formatting = null) =>
            new(text, null, CellValueType.Text, formatting);
        public static CellValue FromFormula(string formula, string result, CellFormatting? formatting = null) =>
            new(result, formula, CellValueType.Formula, formatting);

        public CellValue WithFormatting(CellFormatting formatting) =>
            this with { Formatting = formatting };
    }

    public enum CellValueType
    {
        Empty,
        Text,
        Formula
    }

    public class Cell : INotifyPropertyChanged
    {
        private CellValue _value = CellValue.Empty;

        public CellPosition Position { get; }

        public CellValue Value
        {
            get => _value;
            set
            {
                if (_value != value)
                {
                    _value = value;
                    OnPropertyChanged(nameof(Value));
                    OnPropertyChanged(nameof(DisplayValue));
                }
            }
        }

        public string? DisplayValue => Value.DisplayValue;

        public Cell(CellPosition position) => Position = position;

        public event PropertyChangedEventHandler? PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    public class CellEntity
    {
        public int Id { get; set; }
        public int RowIndex { get; set; }
        public int ColumnIndex { get; set; }
        public string? DisplayValue { get; set; }
        public string? Formula { get; set; }
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public bool IsUnderlined { get; set; }
    }

    public class FuturePortfolioDbContext : DbContext
    {
        public DbSet<CellEntity> Cells { get; set; } = null!;

        public FuturePortfolioDbContext(DbContextOptions<FuturePortfolioDbContext> options)
            : base(options) { }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<CellEntity>(entity =>
            {
                entity.HasKey(e => e.Id);
                entity.HasIndex(e => new { e.RowIndex, e.ColumnIndex })
                      .IsUnique();
                entity.Property(e => e.DisplayValue)
                      .IsRequired(false);
                entity.Property(e => e.Formula)
                      .IsRequired(false);
                entity.Property(e => e.IsBold)
                      .IsRequired()
                      .HasDefaultValue(false);
                entity.Property(e => e.IsItalic)
                      .IsRequired()
                      .HasDefaultValue(false);
                entity.Property(e => e.IsUnderlined)
                      .IsRequired()
                      .HasDefaultValue(false);
            });
        }
    }

    public class CellChangedEventArgs : EventArgs
    {
        public Cell Cell { get; }
        public CellChangedEventArgs(Cell cell) => Cell = cell;
    }

    public interface ISpreadsheetService
    {
        event EventHandler<CellChangedEventArgs>? CellChanged;
        Task<IReadOnlyList<Cell>> GetCellsAsync();
        Task SetCellValueAsync(CellPosition position, string value);
        Task UpdateCellFormattingAsync(CellPosition position, CellFormatting formatting);
        Task LoadCellsAsync(IEnumerable<CellEntity> cells);
        Task AddRowAsync();
        Task AddColumnAsync();
        Task RemoveLastRowAsync();
        Task RemoveLastColumnAsync();
        Task SaveChangesAsync();
        Task ClearAsync();
        int RowCount { get; }
        int ColumnCount { get; }
    }

    public class SpreadsheetService : ISpreadsheetService
    {
        private readonly FuturePortfolioDbContext _context;
        private readonly List<List<Cell>> _grid;
        private readonly SemaphoreSlim _semaphore = new(1, 1);
        private const int DefaultRows = 10;
        private const int DefaultColumns = 10;
        private const int MaxColumns = 26;

        public event EventHandler<CellChangedEventArgs>? CellChanged;
        public int RowCount => _grid.Count;
        public int ColumnCount => _grid.FirstOrDefault()?.Count ?? DefaultColumns;

        public SpreadsheetService(FuturePortfolioDbContext context)
        {
            _context = context;
            _grid = new List<List<Cell>>();
            InitializeEmptyGrid();
        }

        private void InitializeEmptyGrid()
        {
            _grid.Clear();
            for (int i = 0; i < DefaultRows; i++)
            {
                var row = new List<Cell>();
                for (int j = 0; j < DefaultColumns; j++)
                {
                    var cell = new Cell(new CellPosition(i, j));
                    row.Add(cell);
                }
                _grid.Add(row);
            }

            // Notify all cells have been created
            foreach (var cell in _grid.SelectMany(row => row))
            {
                CellChanged?.Invoke(this, new CellChangedEventArgs(cell));
            }
        }

        public async Task<IReadOnlyList<Cell>> GetCellsAsync()
        {
            await _semaphore.WaitAsync();
            try
            {
                var cells = await _context.Cells
                    .AsNoTracking()
                    .OrderBy(c => c.RowIndex)
                    .ThenBy(c => c.ColumnIndex)
                    .ToListAsync();

                EnsureGridDimensions(cells);
                UpdateGridWithData(cells);
                return _grid.SelectMany(row => row).ToList();
            }
            finally
            {
                _semaphore.Release();
            }
        }

        private void EnsureGridDimensions(List<CellEntity> cells)
        {
            if (!cells.Any()) return;

            int maxRow = Math.Max(cells.Max(c => c.RowIndex) + 1, DefaultRows);
            int maxCol = Math.Max(cells.Max(c => c.ColumnIndex) + 1, DefaultColumns);

            // Ensure rows
            while (_grid.Count < maxRow)
            {
                var newRow = new List<Cell>();
                for (int j = 0; j < maxCol; j++)
                {
                    newRow.Add(new Cell(new CellPosition(_grid.Count, j)));
                }
                _grid.Add(newRow);
            }

            // Ensure columns
            foreach (var row in _grid)
            {
                while (row.Count < maxCol)
                {
                    row.Add(new Cell(new CellPosition(_grid.IndexOf(row), row.Count)));
                }
            }
        }

        private void UpdateGridWithData(List<CellEntity> cells)
        {
            foreach (var dbCell in cells)
            {
                if (dbCell.RowIndex < _grid.Count && dbCell.ColumnIndex < _grid[dbCell.RowIndex].Count)
                {
                    var formatting = new CellFormatting()
                        .WithBold(dbCell.IsBold)
                        .WithItalic(dbCell.IsItalic)
                        .WithUnderline(dbCell.IsUnderlined);

                    var cell = _grid[dbCell.RowIndex][dbCell.ColumnIndex];
                    cell.Value = dbCell.Formula != null
                        ? CellValue.FromFormula(dbCell.Formula, dbCell.DisplayValue ?? string.Empty, formatting)
                        : CellValue.FromText(dbCell.DisplayValue ?? string.Empty, formatting);

                    CellChanged?.Invoke(this, new CellChangedEventArgs(cell));
                }
            }
        }

        public async Task SetCellValueAsync(CellPosition position, string value)
        {
            await _semaphore.WaitAsync();
            try
            {
                var cell = _grid[position.Row][position.Column];
                var currentFormatting = cell.Value.Formatting;

                if (string.IsNullOrWhiteSpace(value))
                {
                    cell.Value = CellValue.Empty.WithFormatting(currentFormatting);
                }
                else if (value.StartsWith("="))
                {
                    string formula = value[1..];
                    var result = EvaluateFormula(formula);
                    cell.Value = CellValue.FromFormula(value, result, currentFormatting);
                }
                else
                {
                    cell.Value = CellValue.FromText(value, currentFormatting);
                }

                CellChanged?.Invoke(this, new CellChangedEventArgs(cell));
            }
            finally
            {
                _semaphore.Release();
            }
        }

        private string EvaluateFormula(string formula)
        {
            try
            {
                var dt = new DataTable();
                dt.Columns.Add("result", typeof(double));

                string evaluatableFormula = ReplaceReferences(formula);
                var result = dt.Compute(evaluatableFormula, string.Empty);
                return result?.ToString() ?? "#ERROR";
            }
            catch
            {
                return "#ERROR";
            }
        }

        private string ReplaceReferences(string formula)
        {
            var referencePattern = new Regex(@"([A-Z]+)(\d+)");
            return referencePattern.Replace(formula, match =>
            {
                try
                {
                    var columnName = match.Groups[1].Value;
                    var rowIndex = int.Parse(match.Groups[2].Value) - 1;
                    var columnIndex = ColumnNameToIndex(columnName);

                    if (rowIndex >= 0 && rowIndex < _grid.Count &&
                        columnIndex >= 0 && columnIndex < _grid[0].Count)
                    {
                        var referencedCell = _grid[rowIndex][columnIndex];
                        var value = referencedCell.DisplayValue;
                        return double.TryParse(value, out _) ? value : "0";
                    }
                }
                catch { }
                return "0";
            });
        }

        private int ColumnNameToIndex(string columnName)
        {
            int index = 0;
            for (int i = 0; i < columnName.Length; i++)
            {
                index = index * 26 + (columnName[i] - 'A' + 1);
            }
            return index - 1;
        }

        public async Task SaveChangesAsync()
        {
            await _semaphore.WaitAsync();
            try
            {
                var entities = _grid
                    .SelectMany(row => row)
                    .Where(cell => cell.Value != CellValue.Empty)
                    .Select(cell => new CellEntity
                    {
                        RowIndex = cell.Position.Row,
                        ColumnIndex = cell.Position.Column,
                        DisplayValue = cell.DisplayValue,
                        Formula = cell.Value.Formula,
                        IsBold = cell.Value.Formatting.IsBold,
                        IsItalic = cell.Value.Formatting.IsItalic,
                        IsUnderlined = cell.Value.Formatting.IsUnderlined
                    })
                    .ToList();

                using var transaction = await _context.Database.BeginTransactionAsync();
                try
                {
                    _context.Cells.RemoveRange(await _context.Cells.ToListAsync());
                    await _context.Cells.AddRangeAsync(entities);
                    await _context.SaveChangesAsync();
                    await transaction.CommitAsync();
                }
                catch
                {
                    await transaction.RollbackAsync();
                    throw;
                }
            }
            finally
            {
                _semaphore.Release();
            }
        }

        public async Task UpdateCellFormattingAsync(CellPosition position, CellFormatting formatting)
        {
            await _semaphore.WaitAsync();
            try
            {
                var cell = _grid[position.Row][position.Column];
                cell.Value = cell.Value.WithFormatting(formatting);
                CellChanged?.Invoke(this, new CellChangedEventArgs(cell));
            }
            finally
            {
                _semaphore.Release();
            }
        }

        public async Task LoadCellsAsync(IEnumerable<CellEntity> cells)
        {
            await _semaphore.WaitAsync();
            try
            {
                var cellsList = cells.ToList();
                var maxRow = Math.Max(cellsList.Any() ? cellsList.Max(c => c.RowIndex) + 1 : 0, DefaultRows);
                var maxCol = Math.Max(cellsList.Any() ? cellsList.Max(c => c.ColumnIndex) + 1 : 0, DefaultColumns);

                _grid.Clear();

                for (int i = 0; i < maxRow; i++)
                {
                    var row = new List<Cell>();
                    for (int j = 0; j < maxCol; j++)
                    {
                        var position = new CellPosition(i, j);
                        var cell = new Cell(position);
                        var dbCell = cellsList.FirstOrDefault(c => c.RowIndex == i && c.ColumnIndex == j);

                        if (dbCell != null)
                        {
                            var formatting = new CellFormatting()
                                .WithBold(dbCell.IsBold)
                                .WithItalic(dbCell.IsItalic)
                                .WithUnderline(dbCell.IsUnderlined);

                            cell.Value = dbCell.Formula != null
                                ? CellValue.FromFormula(dbCell.Formula, dbCell.DisplayValue ?? string.Empty, formatting)
                                : CellValue.FromText(dbCell.DisplayValue ?? string.Empty, formatting);
                        }

                        row.Add(cell);
                        CellChanged?.Invoke(this, new CellChangedEventArgs(cell));
                    }
                    _grid.Add(row);
                }
            }
            finally
            {
                _semaphore.Release();
            }
        }

        public async Task ClearAsync()
        {
            await _semaphore.WaitAsync();
            try
            {
                _grid.Clear();
                InitializeEmptyGrid();
            }
            finally
            {
                _semaphore.Release();
            }
        }

        public async Task AddRowAsync()
        {
            await _semaphore.WaitAsync();
            try
            {
                var newRow = new List<Cell>();
                int currentColumnCount = ColumnCount;

                for (int j = 0; j < currentColumnCount; j++)
                {
                    var cell = new Cell(new CellPosition(RowCount, j));
                    newRow.Add(cell);
                    CellChanged?.Invoke(this, new CellChangedEventArgs(cell));
                }
                _grid.Add(newRow);
            }
            finally
            {
                _semaphore.Release();
            }
        }

        public async Task AddColumnAsync()
        {
            await _semaphore.WaitAsync();
            try
            {
                if (ColumnCount >= MaxColumns) return;

                foreach (var row in _grid)
                {
                    var newCell = new Cell(new CellPosition(_grid.IndexOf(row), row.Count));
                    row.Add(newCell);
                }

                // Notify about all cells to force UI refresh
                foreach (var cell in _grid.SelectMany(row => row))
                {
                    CellChanged?.Invoke(this, new CellChangedEventArgs(cell));
                }
            }
            finally
            {
                _semaphore.Release();
            }
        }


        public async Task RemoveLastRowAsync()
        {
            await _semaphore.WaitAsync();
            try
            {
                // Changed minimum row check to 2 to prevent deleting the last row
                if (_grid.Count > 2)
                {
                    _grid.RemoveAt(_grid.Count - 1);

                    // Notify about the change more explicitly
                    foreach (var cell in _grid.SelectMany(row => row))
                    {
                        CellChanged?.Invoke(this, new CellChangedEventArgs(cell));
                    }
                }
            }
            finally
            {
                _semaphore.Release();
            }
        }

        public async Task RemoveLastColumnAsync()
        {
            await _semaphore.WaitAsync();
            try
            {
                if (ColumnCount <= 1) return;

                foreach (var row in _grid)
                {
                    if (row.Count > 0)
                    {
                        row.RemoveAt(row.Count - 1);
                    }
                }

                // Notify about remaining cells to force UI refresh
                foreach (var cell in _grid.SelectMany(row => row))
                {
                    CellChanged?.Invoke(this, new CellChangedEventArgs(cell));
                }
            }
            finally
            {
                _semaphore.Release();
            }
        }
    }

    public partial class MainViewModel : ObservableObject
    {
        private readonly ISpreadsheetService _spreadsheetService;
        private readonly IFileOperationsService _fileOperations;
        private string? _currentFilePath;
        private CellViewModel? _selectedCell;

        [ObservableProperty]
        private ObservableCollection<ObservableCollection<CellViewModel>> _cells = new();

        [ObservableProperty]
        private string _statusMessage = string.Empty;

        [ObservableProperty]
        private bool _isLoading;

        [ObservableProperty]
        private string _searchText = string.Empty;

        [ObservableProperty]
        private bool _isSelectedCellBold;

        [ObservableProperty]
        private bool _isSelectedCellItalic;

        [ObservableProperty]
        private bool _isSelectedCellUnderline;

        [ObservableProperty]
        private ObservableCollection<RowViewModel> _rows = new();

        public MainViewModel(ISpreadsheetService spreadsheetService, IFileOperationsService fileOperations)
        {
            _spreadsheetService = spreadsheetService;
            _fileOperations = fileOperations;
            InitializeGrid(10, 10); // Start with 10x10 grid
        }

        private void InitializeGrid(int rowCount, int columnCount)
        {
            Rows.Clear();
            for (int i = 0; i < rowCount; i++)
            {
                Rows.Add(new RowViewModel(i, columnCount));
            }
        }

        [RelayCommand]
        private void AddRow()
        {
            Rows.Add(new RowViewModel(Rows.Count, Rows[0].Cells.Count));
        }

        [RelayCommand]
        private void AddColumn()
        {
            int newColumnIndex = Rows[0].Cells.Count;
            foreach (var row in Rows)
            {
                row.Cells.Add(new CellViewModel(new Cell(new CellPosition(Rows.IndexOf(row), newColumnIndex))));
            }
            // Force DataGrid to refresh columns
            OnPropertyChanged(nameof(Rows));
        }

        [RelayCommand]
        private void RemoveRow()
        {
            if (Rows.Count > 1)
            {
                Rows.RemoveAt(Rows.Count - 1);
            }
        }

        [RelayCommand]
        private void RemoveColumn()
        {
            if (Rows[0].Cells.Count > 1)
            {
                foreach (var row in Rows)
                {
                    row.Cells.RemoveAt(row.Cells.Count - 1);
                }
                // Force DataGrid to refresh columns
                OnPropertyChanged(nameof(Rows));
            }
        }

        public async Task InitializeAsync()
        {
            if (IsLoading) return;

            try
            {
                IsLoading = true;
                StatusMessage = "Loading...";

                var cells = await _spreadsheetService.GetCellsAsync();
                UpdateGrid(cells);

                StatusMessage = "Ready";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error loading data: {ex.Message}";
                throw;
            }
            finally
            {
                IsLoading = false;
            }
        }

        [RelayCommand]
        private async Task Search()
        {
            if (string.IsNullOrWhiteSpace(SearchText))
            {
                StatusMessage = "Please enter search text";
                return;
            }

            try
            {
                StatusMessage = "Searching...";
                var foundCells = new List<(int Row, int Col)>();

                // Search through all cells
                for (int i = 0; i < Cells.Count; i++)
                {
                    for (int j = 0; j < Cells[i].Count; j++)
                    {
                        var cellText = Cells[i][j].DisplayValue;
                        if (!string.IsNullOrEmpty(cellText) &&
                            cellText.Contains(SearchText, StringComparison.OrdinalIgnoreCase))
                        {
                            foundCells.Add((i, j));
                        }
                    }
                }

                if (foundCells.Count > 0)
                {
                    StatusMessage = $"Found {foundCells.Count} matches";
                    // Highlight first match
                    HighlightCell(foundCells[0].Row, foundCells[0].Col);
                }
                else
                {
                    StatusMessage = "No matches found";
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Search error: {ex.Message}";
            }
        }

        [RelayCommand]
        private async Task NewFile()
        {
            try
            {
                StatusMessage = "Creating new file...";
                await _spreadsheetService.ClearAsync();
                _currentFilePath = null;
                await InitializeAsync();
                StatusMessage = "New file created";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error creating new file: {ex.Message}";
            }
        }

        [RelayCommand]
        private Task FindReplace()
        {
            // Implement find and replace functionality
            StatusMessage = "Find and Replace not implemented yet";
            return Task.CompletedTask;
        }

        [RelayCommand]
        private Task Exit()
        {
            Application.Current.Shutdown();
            return Task.CompletedTask;
        }

        private void HighlightCell(int row, int col)
        {
            if (row < Cells.Count && col < Cells[row].Count)
            {
                _selectedCell = Cells[row][col];
                OnCellSelected(_selectedCell);
            }
        }

        public void OnCellSelected(CellViewModel cell)
        {
            _selectedCell = cell;
            IsSelectedCellBold = cell.Value.Formatting.IsBold;
            IsSelectedCellItalic = cell.Value.Formatting.IsItalic;
            IsSelectedCellUnderline = cell.Value.Formatting.IsUnderlined;
        }

        [RelayCommand]
        private async Task ToggleBold()
        {
            if (_selectedCell == null) return;

            try
            {
                var newFormatting = _selectedCell.Value.Formatting.WithBold(!IsSelectedCellBold);
                await UpdateCellFormattingAsync(_selectedCell.Position, newFormatting);
                IsSelectedCellBold = !IsSelectedCellBold;
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error updating formatting: {ex.Message}";
            }
        }

        [RelayCommand]
        private async Task ToggleItalic()
        {
            if (_selectedCell == null) return;

            try
            {
                var newFormatting = _selectedCell.Value.Formatting.WithItalic(!IsSelectedCellItalic);
                await UpdateCellFormattingAsync(_selectedCell.Position, newFormatting);
                IsSelectedCellItalic = !IsSelectedCellItalic;
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error updating formatting: {ex.Message}";
            }
        }

        [RelayCommand]
        private async Task ToggleUnderline()
        {
            if (_selectedCell == null) return;

            try
            {
                var newFormatting = _selectedCell.Value.Formatting.WithUnderline(!IsSelectedCellUnderline);
                await UpdateCellFormattingAsync(_selectedCell.Position, newFormatting);
                IsSelectedCellUnderline = !IsSelectedCellUnderline;
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error updating formatting: {ex.Message}";
            }
        }

        private async Task UpdateCellFormattingAsync(CellPosition position, CellFormatting formatting)
        {
            await _spreadsheetService.UpdateCellFormattingAsync(position, formatting);
        }

        [RelayCommand]
        private async Task OpenFile()
        {
            var filePath = _fileOperations.ShowOpenFileDialog();
            if (string.IsNullOrEmpty(filePath)) return;

            try
            {
                IsLoading = true;
                StatusMessage = "Loading file...";

                var cells = await _fileOperations.LoadFromFileAsync(filePath);
                await _spreadsheetService.LoadCellsAsync(cells);
                _currentFilePath = filePath;
                await InitializeAsync();

                StatusMessage = "File loaded successfully";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error loading file: {ex.Message}";
            }
            finally
            {
                IsLoading = false;
            }
        }

        [RelayCommand]
        public async Task Save()
        {
            if (string.IsNullOrEmpty(_currentFilePath))
            {
                await SaveAs();
                return;
            }

            await SaveToFile(_currentFilePath);
        }

        [RelayCommand]
        private async Task SaveAs()
        {
            var filePath = _fileOperations.ShowSaveFileDialog();
            if (string.IsNullOrEmpty(filePath)) return;

            await SaveToFile(filePath);
        }

        private async Task SaveToFile(string filePath)
        {
            try
            {
                StatusMessage = "Saving...";
                IsLoading = true;

                var cells = await _spreadsheetService.GetCellsAsync();
                await _fileOperations.SaveToFileAsync(filePath, cells.Select(c => new CellEntity
                {
                    RowIndex = c.Position.Row,
                    ColumnIndex = c.Position.Column,
                    DisplayValue = c.DisplayValue,
                    Formula = c.Value.Formula,
                    IsBold = c.Value.Formatting.IsBold,
                    IsItalic = c.Value.Formatting.IsItalic,
                    IsUnderlined = c.Value.Formatting.IsUnderlined
                }));

                _currentFilePath = filePath;
                StatusMessage = "File saved successfully";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error saving file: {ex.Message}";
            }
            finally
            {
                IsLoading = false;
            }
        }

        private void UpdateGrid(IReadOnlyList<Cell> cells)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                Cells.Clear();

                for (int i = 0; i < _spreadsheetService.RowCount; i++)
                {
                    var row = new ObservableCollection<CellViewModel>();
                    for (int j = 0; j < _spreadsheetService.ColumnCount; j++)
                    {
                        var cell = cells.FirstOrDefault(c => c.Position.Row == i && c.Position.Column == j)
                            ?? new Cell(new CellPosition(i, j));
                        row.Add(new CellViewModel(cell));
                    }
                    Cells.Add(row);
                }
            });
        }

        private void OnCellChanged(object? sender, CellChangedEventArgs e)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                if (e.Cell.Position.Row < Cells.Count &&
                    e.Cell.Position.Column < Cells[e.Cell.Position.Row].Count)
                {
                    var cellVm = Cells[e.Cell.Position.Row][e.Cell.Position.Column];
                    cellVm.UpdateFrom(e.Cell);
                }
                else
                {
                    // If cell position is out of current bounds, refresh the whole grid
                    Task.Run(async () =>
                    {
                        var cells = await _spreadsheetService.GetCellsAsync();
                        UpdateGrid(cells);
                    });
                }
            });
        }

        

        

        public async Task UpdateCellValueAsync(CellPosition position, string value)
        {
            try
            {
                await _spreadsheetService.SetCellValueAsync(position, value);
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error updating cell: {ex.Message}";
                throw;
            }
        }
    }

    public partial class CellViewModel : ObservableObject
    {
        private readonly Cell _cell;

        [ObservableProperty]
        private string _displayValue = string.Empty;

        [ObservableProperty]
        private string _editValue = string.Empty;

        public CellValue Value => _cell.Value;
        public CellPosition Position => _cell.Position;

        public CellViewModel(Cell cell)
        {
            _cell = cell;
            UpdateFrom(cell);
        }

        public void UpdateFrom(Cell cell)
        {
            DisplayValue = cell.DisplayValue ?? string.Empty;
            EditValue = cell.Value.Formula ?? DisplayValue;
            OnPropertyChanged(nameof(Value));
            OnPropertyChanged(nameof(DisplayValue));
            OnPropertyChanged(nameof(EditValue));
        }
    }

    public class RowViewModel : ObservableObject
    {
        private readonly ObservableCollection<CellViewModel> _cells = new();
        private readonly Dictionary<string, CellViewModel> _cellsByColumn = new();

        public ObservableCollection<CellViewModel> Cells => _cells;

        public CellViewModel this[int index]
        {
            get
            {
                var columnName = $"Column{index}";
                if (!_cellsByColumn.ContainsKey(columnName))
                {
                    var cell = new CellViewModel(new Cell(new CellPosition(_cells.Count, index)));
                    _cellsByColumn[columnName] = cell;
                    _cells.Add(cell);
                }
                return _cellsByColumn[columnName];
            }
        }

        public RowViewModel(int rowIndex, int columnCount)
        {
            for (int i = 0; i < columnCount; i++)
            {
                var cell = new CellViewModel(new Cell(new CellPosition(rowIndex, i)));
                var columnName = $"Column{i}";
                _cellsByColumn[columnName] = cell;
                _cells.Add(cell);
            }
        }
    }
}