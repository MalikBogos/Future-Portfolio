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
using System.Windows.Controls;
using System.Windows;
using FuturePortfolio;
using System;

namespace FuturePortfolio
{
    namespace Core
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

        public record CellValue
        {
            public string? DisplayValue { get; init; }
            public string? Formula { get; init; }
            public CellValueType Type { get; init; }

            private CellValue(string? displayValue, string? formula, CellValueType type)
            {
                DisplayValue = displayValue;
                Formula = formula;
                Type = type;
            }

            public static CellValue Empty => new(null, null, CellValueType.Empty);
            public static CellValue FromText(string text) => new(text, null, CellValueType.Text);
            public static CellValue FromFormula(string formula, string result) => new(result, formula, CellValueType.Formula);
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

        // Database Entity
        public class CellEntity
        {
            public int Id { get; set; }
            public int RowIndex { get; set; }
            public int ColumnIndex { get; set; }
            public string? DisplayValue { get; set; }
            public string? Formula { get; set; }
        }

        // DbContext
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
                });
            }
        }

        // Service Interfaces
        public interface ISpreadsheetService
        {
            event EventHandler<CellChangedEventArgs>? CellChanged;
            Task<IReadOnlyList<Cell>> GetCellsAsync();
            Task SetCellValueAsync(CellPosition position, string value);
            Task AddRowAsync();
            Task AddColumnAsync();
            Task RemoveLastRowAsync();
            Task RemoveLastColumnAsync();
            Task SaveChangesAsync();
            int RowCount { get; }
            int ColumnCount { get; }
        }

        public class CellChangedEventArgs : EventArgs
        {
            public Cell Cell { get; }
            public CellChangedEventArgs(Cell cell) => Cell = cell;
        }

        // Service Implementation
        public class SpreadsheetService : ISpreadsheetService
        {
            private readonly FuturePortfolioDbContext _context;
            private readonly List<List<Cell>> _grid;
            private readonly SemaphoreSlim _semaphore = new(1, 1);
            private const int DefaultRows = 10;
            private const int DefaultColumns = 10;

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
                for (int i = 0; i < DefaultRows; i++)
                {
                    var row = new List<Cell>();
                    for (int j = 0; j < DefaultColumns; j++)
                    {
                        row.Add(new Cell(new CellPosition(i, j)));
                    }
                    _grid.Add(row);
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

                    UpdateGridWithData(cells);
                    return _grid.SelectMany(row => row).ToList();
                }
                finally
                {
                    _semaphore.Release();
                }
            }

            private void UpdateGridWithData(List<CellEntity> cells)
            {
                if (!cells.Any()) return;

                int maxRow = Math.Max(cells.Max(c => c.RowIndex) + 1, _grid.Count);
                int maxCol = Math.Max(cells.Max(c => c.ColumnIndex) + 1, ColumnCount);

                // Ensure we have enough rows and columns
                while (_grid.Count < maxRow)
                {
                    var newRow = new List<Cell>();
                    for (int j = 0; j < maxCol; j++)
                    {
                        newRow.Add(new Cell(new CellPosition(_grid.Count, j)));
                    }
                    _grid.Add(newRow);
                }

                // Add columns if needed
                if (maxCol > ColumnCount)
                {
                    foreach (var row in _grid)
                    {
                        while (row.Count < maxCol)
                        {
                            row.Add(new Cell(new CellPosition(_grid.IndexOf(row), row.Count)));
                        }
                    }
                }

                // Update cell values
                foreach (var dbCell in cells)
                {
                    if (dbCell.RowIndex < _grid.Count && dbCell.ColumnIndex < _grid[dbCell.RowIndex].Count)
                    {
                        var cell = _grid[dbCell.RowIndex][dbCell.ColumnIndex];
                        cell.Value = dbCell.Formula != null
                            ? CellValue.FromFormula(dbCell.Formula, dbCell.DisplayValue ?? string.Empty)
                            : CellValue.FromText(dbCell.DisplayValue ?? string.Empty);
                    }
                }
            }

            private void InitializeGrid(List<CellEntity> cells)
            {
                int rows = Math.Max(cells.Any() ? cells.Max(c => c.RowIndex) + 1 : 0, DefaultRows);
                int cols = Math.Max(cells.Any() ? cells.Max(c => c.ColumnIndex) + 1 : 0, DefaultColumns);

                for (int i = 0; i < rows; i++)
                {
                    var row = new List<Cell>();
                    for (int j = 0; j < cols; j++)
                    {
                        var position = new CellPosition(i, j);
                        var cell = new Cell(position);

                        var dbCell = cells.FirstOrDefault(c => c.RowIndex == i && c.ColumnIndex == j);
                        if (dbCell != null)
                        {
                            cell.Value = dbCell.Formula != null
                                ? CellValue.FromFormula(dbCell.Formula, dbCell.DisplayValue ?? string.Empty)
                                : CellValue.FromText(dbCell.DisplayValue ?? string.Empty);
                        }

                        row.Add(cell);
                    }
                    _grid.Add(row);
                }
            }

            public async Task SetCellValueAsync(CellPosition position, string value)
            {
                await _semaphore.WaitAsync();
                try
                {
                    var cell = _grid[position.Row][position.Column];

                    if (string.IsNullOrWhiteSpace(value))
                    {
                        cell.Value = CellValue.Empty;
                    }
                    else if (value.StartsWith("="))
                    {
                        string formula = value[1..];
                        var result = EvaluateFormula(formula);
                        cell.Value = CellValue.FromFormula(value, result);
                    }
                    else
                    {
                        cell.Value = CellValue.FromText(value);
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

                    // Replace Excel-style column references (A1, B2, etc.) with values
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
                // Simple replacement for now - can be expanded for more complex formulas
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
                            Formula = cell.Value.Formula
                        })
                        .ToList();

                    await using var transaction = await _context.Database.BeginTransactionAsync();
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

            // Grid manipulation methods
            public async Task AddRowAsync()
            {
                await _semaphore.WaitAsync();
                try
                {
                    var newRow = new List<Cell>();
                    for (int j = 0; j < ColumnCount; j++)
                    {
                        newRow.Add(new Cell(new CellPosition(RowCount, j)));
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
                    int newColumnIndex = ColumnCount;
                    foreach (var row in _grid)
                    {
                        var cell = new Cell(new CellPosition(_grid.IndexOf(row), newColumnIndex));
                        row.Add(cell);
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
                    if (_grid.Count > 1)
                    {
                        _grid.RemoveAt(_grid.Count - 1);
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
                    if (ColumnCount > 1)
                    {
                        foreach (var row in _grid)
                        {
                            row.RemoveAt(row.Count - 1);
                        }
                    }
                }
                finally
                {
                    _semaphore.Release();
                }
            }
        }

        // ViewModel
        public partial class MainViewModel : ObservableObject
        {
            private readonly ISpreadsheetService _spreadsheetService;

            [ObservableProperty]
            private ObservableCollection<ObservableCollection<CellViewModel>> _cells = new();

            [ObservableProperty]
            private string _statusMessage = string.Empty;

            [ObservableProperty]
            private bool _isLoading;

            public MainViewModel(ISpreadsheetService spreadsheetService)
            {
                _spreadsheetService = spreadsheetService;
                _spreadsheetService.CellChanged += OnCellChanged;
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

            private void UpdateGrid(IReadOnlyList<Cell> cells)
            {
                var rowGroups = cells.GroupBy(c => c.Position.Row)
                                    .OrderBy(g => g.Key);

                Cells.Clear();

                foreach (var rowGroup in rowGroups)
                {
                    var cellRow = new ObservableCollection<CellViewModel>();
                    var row = rowGroup.OrderBy(c => c.Position.Column);

                    // Ensure all columns exist in the row
                    for (int i = 0; i < _spreadsheetService.ColumnCount; i++)
                    {
                        var cell = row.FirstOrDefault(c => c.Position.Column == i);
                        if (cell == null)
                        {
                            cell = new Cell(new CellPosition(rowGroup.Key, i));
                        }
                        cellRow.Add(new CellViewModel(cell));
                    }

                    Cells.Add(cellRow);
                }
            }

            private void OnCellChanged(object? sender, CellChangedEventArgs e)
            {
                if (e.Cell.Position.Row < Cells.Count &&
                    e.Cell.Position.Column < Cells[e.Cell.Position.Row].Count)
                {
                    var cellVm = Cells[e.Cell.Position.Row][e.Cell.Position.Column];
                    cellVm.UpdateFrom(e.Cell);
                }
            }

            [RelayCommand]
            private async Task AddRowAsync()
            {
                try
                {
                    StatusMessage = "Adding row...";
                    await _spreadsheetService.AddRowAsync();
                    await InitializeAsync();
                }
                catch (Exception ex)
                {
                    StatusMessage = $"Error adding row: {ex.Message}";
                    throw;
                }
            }

            [RelayCommand]
            private async Task AddColumnAsync()
            {
                try
                {
                    StatusMessage = "Adding column...";
                    await _spreadsheetService.AddColumnAsync();
                    await InitializeAsync();
                }
                catch (Exception ex)
                {
                    StatusMessage = $"Error adding column: {ex.Message}";
                    throw;
                }
            }

            [RelayCommand]
            private async Task RemoveRowAsync()
            {
                try
                {
                    StatusMessage = "Removing row...";
                    await _spreadsheetService.RemoveLastRowAsync();
                    await InitializeAsync();
                }
                catch (Exception ex)
                {
                    StatusMessage = $"Error removing row: {ex.Message}";
                    throw;
                }
            }

            [RelayCommand]
            private async Task RemoveColumnAsync()
            {
                try
                {
                    StatusMessage = "Removing column...";
                    await _spreadsheetService.RemoveLastColumnAsync();
                    await InitializeAsync();
                }
                catch (Exception ex)
                {
                    StatusMessage = $"Error removing column: {ex.Message}";
                    throw;
                }
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

            [RelayCommand]
            public async Task SaveAsync()
            {
                try
                {
                    StatusMessage = "Saving...";
                    await _spreadsheetService.SaveChangesAsync();
                    StatusMessage = "Saved successfully";
                }
                catch (Exception ex)
                {
                    StatusMessage = $"Error saving: {ex.Message}";
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

            public CellViewModel(Cell cell)
            {
                _cell = cell;
                UpdateFrom(cell);
            }

            public void UpdateFrom(Cell cell)
            {
                DisplayValue = cell.DisplayValue ?? string.Empty;
                EditValue = cell.Value.Formula ?? DisplayValue;
                OnPropertyChanged(nameof(DisplayValue));
                OnPropertyChanged(nameof(EditValue));
            }
        }
    }

    public partial class App : Application
    {
        public static IHost? Host { get; private set; }

        public App()
        {
            Host = CreateHostBuilder().Build();
            InitializeDatabase();
        }

        private void InitializeDatabase()
        {
            using var scope = Host?.Services.CreateScope();
            var db = scope?.ServiceProvider.GetRequiredService<FuturePortfolioDbContext>();

            if (db != null)
            {
                // This will create the database if it doesn't exist
                // and apply any pending migrations
                db.Database.Migrate();
            }
        }

        public static IHostBuilder CreateHostBuilder(string[]? args = null) =>
            Microsoft.Extensions.Hosting.Host.CreateDefaultBuilder(args)
                .ConfigureServices((hostContext, services) =>
                {
                    services.AddDbContextPool<FuturePortfolioDbContext>(options =>
                        options.UseSqlServer(
                            "Server=localhost;Database=FuturePortfolio;Integrated Security=True;TrustServerCertificate=True;"));

                    services.AddScoped<ISpreadsheetService, SpreadsheetService>();
                    services.AddTransient<MainViewModel>();
                });

        protected override async void OnStartup(StartupEventArgs e)
        {
            await Host!.StartAsync();
            base.OnStartup(e);
        }

        protected override async void OnExit(ExitEventArgs e)
        {
            if (Host != null)
            {
                await Host.StopAsync();
                Host.Dispose();
            }
            base.OnExit(e);
        }
    }
}