﻿// <auto-generated />
using FuturePortfolio.Data;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Infrastructure;
using Microsoft.EntityFrameworkCore.Metadata;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion;

#nullable disable

namespace FuturePortfolio.Migrations
{
    [DbContext(typeof(SpreadSheetContext))]
    partial class SpreadSheetContextModelSnapshot : ModelSnapshot
    {
        protected override void BuildModel(ModelBuilder modelBuilder)
        {
#pragma warning disable 612, 618
            modelBuilder
                .HasAnnotation("ProductVersion", "8.0.10")
                .HasAnnotation("Relational:MaxIdentifierLength", 128);

            SqlServerModelBuilderExtensions.UseIdentityColumns(modelBuilder);

            modelBuilder.Entity("FuturePortfolio.Data.Cell", b =>
                {
                    b.Property<int>("Id")
                        .ValueGeneratedOnAdd()
                        .HasColumnType("int");

                    SqlServerPropertyBuilderExtensions.UseIdentityColumn(b.Property<int>("Id"));

                    b.Property<int>("ColumnIndex")
                        .HasColumnType("int");

                    b.Property<string>("Formula")
                        .HasColumnType("nvarchar(max)");

                    b.Property<int>("RowIndex")
                        .HasColumnType("int");

                    b.Property<string>("Value")
                        .HasColumnType("nvarchar(max)");

                    b.HasKey("Id");

                    b.HasIndex("RowIndex", "ColumnIndex");

                    b.ToTable("Cells");
                });

            modelBuilder.Entity("FuturePortfolio.Data.CellFormat", b =>
                {
                    b.Property<int>("Id")
                        .ValueGeneratedOnAdd()
                        .HasColumnType("int");

                    SqlServerPropertyBuilderExtensions.UseIdentityColumn(b.Property<int>("Id"));

                    b.Property<string>("BackgroundColorHex")
                        .IsRequired()
                        .ValueGeneratedOnAdd()
                        .HasMaxLength(50)
                        .HasColumnType("nvarchar(50)")
                        .HasDefaultValue("#FFFFFF");

                    b.Property<int>("CellId")
                        .HasColumnType("int");

                    b.Property<string>("FontStyleString")
                        .IsRequired()
                        .ValueGeneratedOnAdd()
                        .HasMaxLength(50)
                        .HasColumnType("nvarchar(50)")
                        .HasDefaultValue("Normal");

                    b.Property<double>("FontWeightValue")
                        .ValueGeneratedOnAdd()
                        .HasColumnType("float")
                        .HasDefaultValue(4.0);

                    b.Property<string>("ForegroundColorHex")
                        .IsRequired()
                        .ValueGeneratedOnAdd()
                        .HasMaxLength(50)
                        .HasColumnType("nvarchar(50)")
                        .HasDefaultValue("#000000");

                    b.HasKey("Id");

                    b.HasIndex("CellId")
                        .IsUnique();

                    b.ToTable("CellFormats");
                });

            modelBuilder.Entity("FuturePortfolio.Data.CellFormat", b =>
                {
                    b.HasOne("FuturePortfolio.Data.Cell", "Cell")
                        .WithOne("Format")
                        .HasForeignKey("FuturePortfolio.Data.CellFormat", "CellId")
                        .OnDelete(DeleteBehavior.Cascade)
                        .IsRequired();

                    b.Navigation("Cell");
                });

            modelBuilder.Entity("FuturePortfolio.Data.Cell", b =>
                {
                    b.Navigation("Format")
                        .IsRequired();
                });
#pragma warning restore 612, 618
        }
    }
}