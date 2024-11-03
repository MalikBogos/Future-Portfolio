using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.ObjectModel;
using System.Linq;

namespace FuturePortfolio.Data
{
    public class SpreadSheetContext : DbContext
    {
        public DbSet<Cell> Cells { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            if (!optionsBuilder.IsConfigured)
            {
                optionsBuilder.UseSqlServer(
                    "Server=localhost;Database=DoWell;Integrated Security=True;TrustServerCertificate=True;",
                    options => options.EnableRetryOnFailure()
                );
            }
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Cell>(entity =>
            {
                entity.HasKey(e => e.Id);
                entity.Property(e => e.RowIndex).IsRequired();
                entity.Property(e => e.ColumnIndex).IsRequired();
                entity.Property(e => e.Value).IsRequired(false);
                entity.Property(e => e.Formula).IsRequired(false);
                entity.HasIndex(e => new { e.RowIndex, e.ColumnIndex });
            });
        }

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
                var cells = _context.Cells
                    .AsNoTracking()
                    .OrderBy(c => c.RowIndex)
                    .ThenBy(c => c.ColumnIndex)
                    .ToList();

                int maxRow = cells.Any() ? cells.Max(c => c.RowIndex) : 0;
                int maxCol = cells.Any() ? cells.Max(c => c.ColumnIndex) : 0;

                // Ensure at least one row and column
                maxRow = Math.Max(maxRow, 0);
                maxCol = Math.Max(maxCol, 0);

                for (int i = 0; i <= maxRow; i++)
                {
                    var wpfRow = new ObservableCollection<WpfCell>();
                    for (int j = 0; j <= maxCol; j++)
                    {
                        var dbCell = cells.FirstOrDefault(c => c.RowIndex == i && c.ColumnIndex == j);
                        var wpfCell = new WpfCell
                        {
                            Id = dbCell?.Id ?? 0,
                            RowIndex = i,
                            ColumnIndex = j,
                            Value = dbCell?.Value ?? "",
                            Formula = dbCell?.Formula
                        };
                        wpfRow.Add(wpfCell);
                    }
                    Add(wpfRow);
                }
            }

            public void SaveToDatabase()
            {
                // Clear existing cells
                _context.Cells.RemoveRange(_context.Cells);
                _context.SaveChanges();

                // Create new cells from WpfCells
                var cellsToAdd = new List<Cell>();

                for (int rowIndex = 0; rowIndex < Count; rowIndex++)
                {
                    for (int colIndex = 0; colIndex < this[rowIndex].Count; colIndex++)
                    {
                        var wpfCell = this[rowIndex][colIndex];
                        if (!string.IsNullOrEmpty(wpfCell.Value) || !string.IsNullOrEmpty(wpfCell.Formula))
                        {
                            var cell = new Cell
                            {
                                RowIndex = rowIndex,
                                ColumnIndex = colIndex,
                                Value = wpfCell.Value,
                                Formula = wpfCell.Formula
                            };
                            cellsToAdd.Add(cell);
                        }
                    }
                }

                // Add all new cells at once
                _context.Cells.AddRange(cellsToAdd);
                _context.SaveChanges();
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
}