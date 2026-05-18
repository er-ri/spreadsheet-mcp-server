using SpreadsheetMcpServer.Core.Models;

namespace SpreadsheetMcpServer.Core.Services;

/// <summary>
/// Service for reading and updating cells in Excel worksheets.
/// </summary>
public interface ICellService
{
    /// <summary>
    /// Reads all cells from the specified worksheet.
    /// </summary>
    List<SerializableCell> ReadSpreadSheet(string spreadSheetPath, string spreadSheetName);

    /// <summary>
    /// Updates the specified cells in an Excel file.
    /// Each <see cref="SerializableCell.Address"/> must be in the format "SheetName!CellRef" (e.g. "Sheet1!A1").
    /// String values are stored via the shared string table; numeric/boolean/other values are stored inline.
    /// </summary>
    void UpdateCells(string spreadSheetPath, List<SerializableCell> cells);
}
