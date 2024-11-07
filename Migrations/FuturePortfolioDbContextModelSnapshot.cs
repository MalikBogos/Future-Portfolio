﻿// <auto-generated />
using FuturePortfolio.Core;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Infrastructure;
using Microsoft.EntityFrameworkCore.Metadata;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion;

#nullable disable

namespace FuturePortfolio.Migrations
{
    [DbContext(typeof(FuturePortfolioDbContext))]
    partial class FuturePortfolioDbContextModelSnapshot : ModelSnapshot
    {
        protected override void BuildModel(ModelBuilder modelBuilder)
        {
#pragma warning disable 612, 618
            modelBuilder
                .HasAnnotation("ProductVersion", "8.0.10")
                .HasAnnotation("Relational:MaxIdentifierLength", 128);

            SqlServerModelBuilderExtensions.UseIdentityColumns(modelBuilder);

            modelBuilder.Entity("FuturePortfolio.Core.CellEntity", b =>
                {
                    b.Property<int>("Id")
                        .ValueGeneratedOnAdd()
                        .HasColumnType("int");

                    SqlServerPropertyBuilderExtensions.UseIdentityColumn(b.Property<int>("Id"));

                    b.Property<int>("ColumnIndex")
                        .HasColumnType("int");

                    b.Property<string>("DisplayValue")
                        .HasColumnType("nvarchar(max)");

                    b.Property<string>("Formula")
                        .HasColumnType("nvarchar(max)");

                    b.Property<int>("RowIndex")
                        .HasColumnType("int");

                    b.HasKey("Id");

                    b.HasIndex("RowIndex", "ColumnIndex")
                        .IsUnique();

                    b.ToTable("Cells");
                });
#pragma warning restore 612, 618
        }
    }
}