using SpreadsheetMcpServer.Core.Models;

namespace SpreadsheetMcpServer.Core.Services;

/// <summary>
/// Service for managing Excel tables.
/// </summary>
public interface ITableService
{
    /// <summary>
    /// Returns all Excel tables defined on the specified worksheet.
    /// </summary>
    List<SerializableTable> GetTables(string spreadSheetPath, string sheetName);

    /// <summary>
    /// Creates an Excel table over the given range on the specified sheet.
    /// The first row of the range must already contain header values.
    /// </summary>
    void CreateTable(string spreadSheetPath, SerializableTable tableInfo, string sheetName);

    /// <summary>
    /// Removes the named table from the specified sheet.
    /// </summary>
    void DeleteTable(string spreadSheetPath, string sheetName, string tableName);
}
