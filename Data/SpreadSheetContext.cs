using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace FuturePortfolio.Data
{
    public class SpreadSheetContext : DbContext
    {
        public DbSet<Cell> Cells { get; set; }
        public DbSet<CellFormat> CellFormats { get; set; }

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

                entity.Property(e => e.RowIndex)
                      .IsRequired();

                entity.Property(e => e.ColumnIndex)
                      .IsRequired();

                entity.Property(e => e.Value)
                      .IsRequired(false);

                entity.Property(e => e.Formula)
                      .IsRequired(false);

                entity.HasIndex(e => new { e.RowIndex, e.ColumnIndex });

                entity.HasOne(c => c.Format)
                      .WithOne(f => f.Cell)
                      .HasForeignKey<CellFormat>(f => f.CellId)
                      .OnDelete(DeleteBehavior.Cascade);
            });

            modelBuilder.Entity<CellFormat>(entity =>
            {
                entity.HasKey(e => e.Id);

                entity.Property(e => e.FontStyleString)
                      .HasDefaultValue("Normal")
                      .IsRequired()
                      .HasMaxLength(50);

                entity.Property(e => e.FontWeightValue)
                      .HasDefaultValue(4.0)
                      .IsRequired();

                entity.Property(e => e.ForegroundColorHex)
                      .HasDefaultValue("#000000")
                      .IsRequired()
                      .HasMaxLength(50);

                entity.Property(e => e.BackgroundColorHex)
                      .HasDefaultValue("#FFFFFF")
                      .IsRequired()
                      .HasMaxLength(50);
            });
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
                    .Include(c => c.Format)
                    .AsNoTracking()
                    .OrderBy(c => c.RowIndex)
                    .ThenBy(c => c.ColumnIndex)
                    .ToList();

                int maxRow = cells.Any() ? cells.Max(c => c.RowIndex) : 0;
                int maxCol = cells.Any() ? cells.Max(c => c.ColumnIndex) : 0;

                for (int i = 0; i <= maxRow; i++)
                {
                    var wpfRow = new ObservableCollection<WpfCell>();
                    for (int j = 0; j <= maxCol; j++)
                    {
                        var dbCell = cells.FirstOrDefault(c => c.RowIndex == i && c.ColumnIndex == j);
                        if (dbCell == null)
                        {
                            dbCell = new Cell
                            {
                                RowIndex = i,
                                ColumnIndex = j,
                                Format = new CellFormat()
                            };
                        }
                        wpfRow.Add(DataConverter.ToWpfCell(dbCell));
                    }
                    Add(wpfRow);
                }
            }

            public void SaveToDatabase()
            {
                _context.CellFormats.RemoveRange(_context.CellFormats);
                _context.Cells.RemoveRange(_context.Cells);
                _context.SaveChanges();
                try
                {
                    _context.CellFormats.RemoveRange(_context.CellFormats);
                    _context.Cells.RemoveRange(_context.Cells);
                    _context.SaveChanges();

                    for (int rowIndex = 0; rowIndex < Count; rowIndex++)
                    {
                        for (int colIndex = 0; colIndex < this[rowIndex].Count; colIndex++)
                        {
                            var wpfCell = this[rowIndex][colIndex];
                            if (!string.IsNullOrEmpty(wpfCell.Value) || !string.IsNullOrEmpty(wpfCell.Formula))
                            {
                                wpfCell.RowIndex = rowIndex;
                                wpfCell.ColumnIndex = colIndex;
                                var dbCell = DataConverter.ToDbCell(wpfCell);
                                _context.Cells.Add(dbCell);
                            }
                        }
                    }
                    _context.SaveChanges();
                }
                catch (Exception ex)
                {
                    throw;
                }
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
                        Format = new WpfCellFormat()
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
                        Format = new WpfCellFormat()
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
