using SpreadsheetMcpServer.Core.Models;

namespace SpreadsheetMcpServer.Core.Services;

/// <summary>
/// Service for managing Excel tables.
/// </summary>
public interface ITableService
{
    /// <summary>
    /// Gets, creates, or deletes a table on the specified worksheet. 'action' must be
    /// "get", "create", or "delete". 'table' is required for "create"; 'tableName' is
    /// required for "delete". Returns the tables remaining on the sheet after the operation.
    /// </summary>
    List<SerializableTable> ManageTables(
        string spreadSheetPath,
        string sheetName,
        string action,
        string? tableName = null,
        SerializableTable? table = null
    );
}
