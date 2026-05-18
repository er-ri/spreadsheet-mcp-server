using System.Globalization;
using ClosedXML.Excel;
using SpreadsheetMcpServer.Core.Models;

namespace SpreadsheetMcpServer.Core.Services;

/// <summary>
/// Service for reading and updating cells in Excel worksheets.
/// </summary>
public class CellService : ICellService
{
    /// <summary>
    /// Reads all cells from the specified worksheet.
    /// </summary>
    public List<SerializableCell> ReadSpreadSheet(string spreadSheetPath, string spreadSheetName)
    {
        try
        {
            using var workbook = new XLWorkbook(spreadSheetPath);
            var worksheet = workbook.Worksheet(spreadSheetName);
            var cells = new List<SerializableCell>();

            foreach (var cell in worksheet.CellsUsed())
            {
                string cellRef = cell.Address.ToString()!;
                string? formula = cell.HasFormula ? cell.FormulaA1 : null;

                string? dataType = cell.DataType switch
                {
                    XLDataType.Text => "SharedString",
                    XLDataType.Number => "Number",
                    XLDataType.Boolean => "Boolean",
                    XLDataType.DateTime => "Date",
                    XLDataType.TimeSpan => "Number",
                    XLDataType.Error => "Error",
                    XLDataType.Blank => null,
                    _ => null,
                };

                string? value = cell.DataType == XLDataType.Blank ? null : cell.Value.ToString();

                cells.Add(
                    new SerializableCell
                    {
                        Address = $"{spreadSheetName}!{cellRef}",
                        DataType = dataType,
                        RawValue = value,
                        Value = value,
                        Formula = formula,
                    }
                );
            }

            return cells;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Error reading Excel file: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Updates the specified cells in an Excel file.
    /// Each <see cref="SerializableCell.Address"/> must be in the format "SheetName!CellRef" (e.g. "Sheet1!A1").
    /// </summary>
    public void UpdateCells(string spreadSheetPath, List<SerializableCell> cells)
    {
        if (cells == null || cells.Count == 0)
            return;

        try
        {
            using var workbook = new XLWorkbook(spreadSheetPath);

            var cellsBySheet = cells.Where(c => c.Address.Contains('!')).GroupBy(c => c.Address.Split('!')[0]);

            foreach (var sheetGroup in cellsBySheet)
            {
                var worksheet = workbook.Worksheet(sheetGroup.Key);

                foreach (var cellInfo in sheetGroup)
                {
                    string cellRef = cellInfo.Address.Split('!')[1];
                    var cell = worksheet.Cell(cellRef);

                    if (cellInfo.Value == null && cellInfo.Formula == null)
                    {
                        cell.Clear();
                    }
                    else if (cellInfo.Formula != null)
                    {
                        cell.FormulaA1 = cellInfo.Formula;
                    }
                    else if (
                        cellInfo.DataType == "Number"
                        && double.TryParse(
                            cellInfo.Value,
                            NumberStyles.Any,
                            CultureInfo.InvariantCulture,
                            out double numVal
                        )
                    )
                    {
                        cell.Value = numVal;
                    }
                    else if (cellInfo.DataType == "Boolean" && bool.TryParse(cellInfo.Value, out bool boolVal))
                    {
                        cell.Value = boolVal;
                    }
                    else if (cellInfo.DataType == "Date" && DateTime.TryParse(cellInfo.Value, out DateTime dateVal))
                    {
                        cell.Value = dateVal;
                    }
                    else
                    {
                        cell.Value = cellInfo.Value;
                    }
                }
            }

            workbook.Save();
        }
        catch (Exception ex) when (ex is not InvalidOperationException)
        {
            throw new InvalidOperationException($"Error updating Excel file: {ex.Message}", ex);
        }
    }
}
