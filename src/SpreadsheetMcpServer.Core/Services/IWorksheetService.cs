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

    /// <summary>
    /// Sets or clears frozen panes on a worksheet. 'action' must be "set" or "clear". For "set",
    /// 'freezeAtAddress' (e.g. "B2") is required and freezes every row above and every column left
    /// of that cell — passing "A1" freezes nothing, since there are no rows above or columns left of
    /// it. For "clear", any existing freeze is removed. Returns a confirmation message.
    /// </summary>
    string ManageFreezePanes(
        string spreadSheetPath,
        string spreadSheetName,
        string action,
        string? freezeAtAddress = null
    );

    /// <summary>
    /// Inserts or deletes whole rows or columns on a worksheet, shifting subsequent cells. 'action'
    /// must be "insertRows", "deleteRows", "insertColumns", or "deleteColumns". 'target' is a 1-based
    /// row number (for the Rows actions) or a column letter (for the Columns actions) marking where
    /// the operation applies — for insert, the new rows/columns are inserted starting at 'target',
    /// pushing existing content at and after 'target' forward; for delete, the rows/columns starting
    /// at 'target' are removed. 'count' (default 1) is how many rows/columns to insert or delete.
    /// Returns a confirmation message.
    /// </summary>
    string ManageRowsColumns(
        string spreadSheetPath,
        string spreadSheetName,
        string action,
        string target,
        int count = 1
    );
}
