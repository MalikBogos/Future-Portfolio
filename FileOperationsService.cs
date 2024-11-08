using System.Text.Json;
using Microsoft.Win32;
using FuturePortfolio.Core;
using System.IO;

namespace FuturePortfolio
{
    public interface IFileOperationsService
    {
        Task SaveToFileAsync(string filePath, IEnumerable<CellEntity> cells);
        Task<IEnumerable<CellEntity>> LoadFromFileAsync(string filePath);
        string ShowSaveFileDialog();
        string ShowOpenFileDialog();
    }

    public class FileOperationsService : IFileOperationsService
    {
        public async Task SaveToFileAsync(string filePath, IEnumerable<CellEntity> cells)
        {
            var data = new
            {
                Version = 1,
                SaveDate = DateTime.UtcNow,
                Cells = cells.Select(c => new
                {
                    c.RowIndex,
                    c.ColumnIndex,
                    c.DisplayValue,
                    c.Formula,
                    c.IsBold,
                    c.IsItalic,
                    c.IsUnderlined
                })
            };

            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            };

            await File.WriteAllTextAsync(filePath, JsonSerializer.Serialize(data, options));
        }

        public async Task<IEnumerable<CellEntity>> LoadFromFileAsync(string filePath)
        {
            try
            {
                var json = await File.ReadAllTextAsync(filePath);
                using var document = JsonDocument.Parse(json);
                var root = document.RootElement;

                // Version check
                if (root.GetProperty("version").GetInt32() != 1)
                {
                    throw new InvalidOperationException("Unsupported file version");
                }

                var cellsArray = root.GetProperty("cells");
                var cells = new List<CellEntity>();

                foreach (var cell in cellsArray.EnumerateArray())
                {
                    cells.Add(new CellEntity
                    {
                        RowIndex = cell.GetProperty("rowIndex").GetInt32(),
                        ColumnIndex = cell.GetProperty("columnIndex").GetInt32(),
                        DisplayValue = cell.GetProperty("displayValue").GetString(),
                        Formula = cell.TryGetProperty("formula", out var formula) ? formula.GetString() : null,
                        IsBold = cell.GetProperty("isBold").GetBoolean(),
                        IsItalic = cell.GetProperty("isItalic").GetBoolean(),
                        IsUnderlined = cell.GetProperty("isUnderlined").GetBoolean()
                    });
                }

                return cells;
            }
            catch (JsonException ex)
            {
                throw new InvalidOperationException("Invalid file format", ex);
            }
        }

        public string ShowSaveFileDialog()
        {
            var dialog = new SaveFileDialog
            {
                Filter = "Spreadsheet files (*.json)|*.json|All files (*.*)|*.*",
                Title = "Save Spreadsheet",
                DefaultExt = ".json",
                AddExtension = true
            };

            return dialog.ShowDialog() == true ? dialog.FileName : string.Empty;
        }

        public string ShowOpenFileDialog()
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Spreadsheet files (*.json)|*.json|All files (*.*)|*.*",
                Title = "Open Spreadsheet"
            };

            return dialog.ShowDialog() == true ? dialog.FileName : string.Empty;
        }
    }
}