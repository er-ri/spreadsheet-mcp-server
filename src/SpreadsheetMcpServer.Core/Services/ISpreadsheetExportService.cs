namespace SpreadsheetMcpServer.Core.Services;

/// <summary>
/// Service for exporting data to Excel files.
/// </summary>
public interface ISpreadsheetExportService
{
    /// <summary>
    /// Exports a JSON string to an Excel file.
    /// </summary>
    /// <param name="spreadSheetPath">The file path where the Excel file will be saved.</param>
    /// <param name="sheetName">The name of the worksheet to create.</param>
    /// <param name="json">A JSON string representing an array of objects (or a single object) to export as rows.</param>
    void ExportJsonArrayToSpreadSheet(string spreadSheetPath, string sheetName, string json);
}
