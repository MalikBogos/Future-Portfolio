using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
    }
}
