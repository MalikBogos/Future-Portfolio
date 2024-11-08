using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace FuturePortfolio.Migrations
{
    /// <inheritdoc />
    public partial class InitialMigration : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "Cells",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    RowIndex = table.Column<int>(type: "int", nullable: false),
                    ColumnIndex = table.Column<int>(type: "int", nullable: false),
                    DisplayValue = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Formula = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    IsBold = table.Column<bool>(type: "bit", nullable: false, defaultValue: false),
                    IsItalic = table.Column<bool>(type: "bit", nullable: false, defaultValue: false),
                    IsUnderlined = table.Column<bool>(type: "bit", nullable: false, defaultValue: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Cells", x => x.Id);
                });

            migrationBuilder.CreateIndex(
                name: "IX_Cells_RowIndex_ColumnIndex",
                table: "Cells",
                columns: new[] { "RowIndex", "ColumnIndex" },
                unique: true);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "Cells");
        }
    }
}
