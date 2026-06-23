using SpreadsheetMcpServer.Core.Models;

namespace SpreadsheetMcpServer.Core.Services;

/// <summary>
/// Service for managing Excel worksheets.
/// </summary>
public interface IWorksheetService
{
    /// <summary>
    /// Gets a list of all worksheets in the spreadsheet, each with its name and used range.
    /// </summary>
    List<SerializableSheet> GetAllWorksheets(string spreadSheetPath);

    /// <summary>
    /// Creates, deletes, or renames a worksheet. 'action' must be "create", "delete", or
    /// "rename"; 'newSheetName' is required only when renaming. Returns a confirmation message.
    /// </summary>
    string ManageWorksheet(string spreadSheetPath, string spreadSheetName, string action, string? newSheetName = null);
}
