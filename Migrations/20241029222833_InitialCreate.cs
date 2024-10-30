using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace FuturePortfolio.Migrations
{
    /// <inheritdoc />
    public partial class InitialCreate : Migration
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
                    Value = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Formula = table.Column<string>(type: "nvarchar(max)", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Cells", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "CellFormats",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    FontStyleString = table.Column<string>(type: "nvarchar(50)", maxLength: 50, nullable: false, defaultValue: "Normal"),
                    FontWeightValue = table.Column<double>(type: "float", nullable: false, defaultValue: 4.0),
                    ForegroundColorHex = table.Column<string>(type: "nvarchar(50)", maxLength: 50, nullable: false, defaultValue: "#000000"),
                    BackgroundColorHex = table.Column<string>(type: "nvarchar(50)", maxLength: 50, nullable: false, defaultValue: "#FFFFFF"),
                    CellId = table.Column<int>(type: "int", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_CellFormats", x => x.Id);
                    table.ForeignKey(
                        name: "FK_CellFormats_Cells_CellId",
                        column: x => x.CellId,
                        principalTable: "Cells",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateIndex(
                name: "IX_CellFormats_CellId",
                table: "CellFormats",
                column: "CellId",
                unique: true);

            migrationBuilder.CreateIndex(
                name: "IX_Cells_RowIndex_ColumnIndex",
                table: "Cells",
                columns: new[] { "RowIndex", "ColumnIndex" });
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "CellFormats");

            migrationBuilder.DropTable(
                name: "Cells");
        }
    }
}
