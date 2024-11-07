using Microsoft.EntityFrameworkCore;
using System.Collections.ObjectModel;
using System.Collections.Specialized;

namespace FuturePortfolio.Data
{
    public class SpreadSheetContext : DbContext
    {
        private const int MinRows = 1;
        private const int MinColumns = 1;
        private const string ConnectionString = "Server=localhost;Database=DoWell;Integrated Security=True;TrustServerCertificate=True;";

        public DbSet<Cell> Cells { get; set; } = null!;

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            if (!optionsBuilder.IsConfigured)
            {
                optionsBuilder.UseSqlServer(
                    ConnectionString,
                    options => options.EnableRetryOnFailure(
                        maxRetryCount: 3,
                        maxRetryDelay: TimeSpan.FromSeconds(5),
                        errorNumbersToAdd: null)
                );
            }
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Cell>(entity =>
            {
                entity.HasKey(e => e.Id);
                entity.HasIndex(e => new { e.RowIndex, e.ColumnIndex })
                      .HasDatabaseName("IX_CellPosition")
                      .IsUnique();
            });
        }

        public class SpreadsheetDataWithDb : ObservableCollection<ObservableCollection<WpfCell>>, IDisposable
        {
            private readonly SpreadSheetContext _context;
            private bool _isDisposed;
            private readonly SemaphoreSlim _semaphore = new(1, 1);
            private const int DefaultRowCount = 1;
            private const int DefaultColumnCount = 5;

            public SpreadsheetDataWithDb(SpreadSheetContext context)
            {
                _context = context ?? throw new ArgumentNullException(nameof(context));
                this.CollectionChanged += HandleCollectionChanged;
                LoadFromDatabase();
                EnsureMinimumGrid();
            }

            private void EnsureMinimumGrid()
            {
                if (Count == 0)
                {
                    for (int i = 0; i < DefaultRowCount; i++)
                    {
                        var row = new ObservableCollection<WpfCell>();
                        for (int j = 0; j < DefaultColumnCount; j++)
                        {
                            row.Add(CreateWpfCell(i, j));
                        }
                        Add(row);
                    }
                }
                else if (this[0].Count == 0)
                {
                    for (int i = 0; i < Count; i++)
                    {
                        for (int j = 0; j < DefaultColumnCount; j++)
                        {
                            this[i].Add(CreateWpfCell(i, j));
                        }
                    }
                }
            }

            private void LoadFromDatabase()
            {
                var cells = _context.Cells
                    .AsNoTracking()
                    .OrderBy(c => c.RowIndex)
                    .ThenBy(c => c.ColumnIndex)
                    .ToList();

                var dimensions = GetGridDimensions(cells);
                InitializeGrid(dimensions.rows, dimensions.cols, cells);
            }

            private (int rows, int cols) GetGridDimensions(List<Cell> cells)
            {
                if (!cells.Any())
                    return (MinRows, MinColumns);

                return (
                    rows: Math.Max(cells.Max(c => c.RowIndex) + 1, MinRows),
                    cols: Math.Max(cells.Max(c => c.ColumnIndex) + 1, MinColumns)
                );
            }

            private void InitializeGrid(int rows, int cols, List<Cell> cells)
            {
                for (int i = 0; i < rows; i++)
                {
                    var wpfRow = new ObservableCollection<WpfCell>();
                    wpfRow.CollectionChanged += HandleRowCollectionChanged;

                    for (int j = 0; j < cols; j++)
                    {
                        var dbCell = cells.FirstOrDefault(c => c.RowIndex == i && c.ColumnIndex == j);
                        wpfRow.Add(CreateWpfCell(i, j, dbCell));
                    }
                    Add(wpfRow);
                }
            }

            private WpfCell CreateWpfCell(int rowIndex, int colIndex, Cell? dbCell = null) =>
                new()
                {
                    Id = dbCell?.Id ?? 0,
                    RowIndex = rowIndex,
                    ColumnIndex = colIndex,
                    Value = dbCell?.Value,
                    Formula = dbCell?.Formula
                };

            public async Task SaveToDatabaseAsync()
            {
                if (_isDisposed)
                    throw new ObjectDisposedException(nameof(SpreadsheetDataWithDb));

                await _semaphore.WaitAsync();
                try
                {
                    var modifiedCells = this.SelectMany((row, i) =>
                        row.Select((cell, j) => cell))
                        .Where(cell => cell.IsModified &&
                               (!string.IsNullOrEmpty(cell.Value) || !string.IsNullOrEmpty(cell.Formula)))
                        .Select(cell => new Cell
                        {
                            Id = cell.Id,
                            RowIndex = cell.RowIndex,
                            ColumnIndex = cell.ColumnIndex,
                            Value = cell.Value,
                            Formula = cell.Formula
                        })
                        .ToList();

                    if (!modifiedCells.Any())
                        return;

                    using var transaction = await _context.Database.BeginTransactionAsync();
                    try
                    {
                        _context.Cells.RemoveRange(_context.Cells);
                        _context.Cells.AddRange(modifiedCells);
                        await _context.SaveChangesAsync();
                        await transaction.CommitAsync();

                        foreach (var row in this)
                            foreach (var cell in row)
                                cell.ResetModifiedFlag();
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

            public void AddRow()
            {
                var newRow = new ObservableCollection<WpfCell>();
                int columnCount = this.FirstOrDefault()?.Count ?? MinColumns;

                for (int j = 0; j < columnCount; j++)
                    newRow.Add(CreateWpfCell(Count, j));

                newRow.CollectionChanged += HandleRowCollectionChanged;
                Add(newRow);
            }

            public void AddColumn()
            {
                int newColumnIndex = this.FirstOrDefault()?.Count ?? 0;
                foreach (var row in this)
                    row.Add(CreateWpfCell(this.IndexOf(row), newColumnIndex));
            }

            public void RemoveLastRow()
            {
                if (Count > MinRows)
                {
                    var lastRow = this[Count - 1];
                    lastRow.CollectionChanged -= HandleRowCollectionChanged;
                    RemoveAt(Count - 1);
                }
            }

            public void RemoveLastColumn()
            {
                if (this.FirstOrDefault()?.Count > MinColumns)
                {
                    foreach (var row in this)
                        row.RemoveAt(row.Count - 1);
                }
            }

            private void HandleCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
            {
                if (e.NewItems != null)
                {
                    foreach (ObservableCollection<WpfCell> row in e.NewItems)
                        row.CollectionChanged += HandleRowCollectionChanged;
                }

                if (e.OldItems != null)
                {
                    foreach (ObservableCollection<WpfCell> row in e.OldItems)
                        row.CollectionChanged -= HandleRowCollectionChanged;
                }
            }

            private void HandleRowCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
            {
            }

            public void Dispose()
            {
                if (!_isDisposed)
                {
                    foreach (var row in this)
                        row.CollectionChanged -= HandleRowCollectionChanged;
                    this.CollectionChanged -= HandleCollectionChanged;
                    _context.Dispose();
                    _semaphore.Dispose();
                    _isDisposed = true;
                }
                GC.SuppressFinalize(this);
            }
        }
    }
}