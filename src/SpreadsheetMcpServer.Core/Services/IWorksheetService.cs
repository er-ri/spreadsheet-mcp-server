namespace SpreadsheetMcpServer.Core.Services;

/// <summary>
/// Service for managing Excel worksheets.
/// </summary>
public interface IWorksheetService
{
    /// <summary>
    /// Gets a list of all worksheet names in the spreadsheet.
    /// </summary>
    List<string> GetAllWorksheets(string spreadSheetPath);
}
