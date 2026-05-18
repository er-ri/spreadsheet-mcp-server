using ClosedXML.Excel;
using SpreadsheetMcpServer.Core.Helpers;
using SpreadsheetMcpServer.Core.Models;

namespace SpreadsheetMcpServer.Core.Services;

/// <summary>
/// Service for managing Excel tables.
/// </summary>
public class TableService : ITableService
{
    /// <summary>
    /// Returns all Excel tables defined on the specified worksheet.
    /// </summary>
    public List<SerializableTable> GetTables(string spreadSheetPath, string sheetName)
    {
        try
        {
            using var workbook = new XLWorkbook(spreadSheetPath);
            var worksheet = workbook.Worksheet(sheetName);
            return worksheet.Tables.Select(ReadSerializableTable).ToList();
        }
        catch (Exception ex) when (ex is not InvalidOperationException)
        {
            throw new InvalidOperationException($"Error reading tables: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Creates an Excel table over the given range on the specified sheet.
    /// The first row of the range must already contain header values.
    /// </summary>
    public void CreateTable(string spreadSheetPath, SerializableTable tableInfo, string sheetName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(tableInfo.Name);
        ArgumentException.ThrowIfNullOrWhiteSpace(tableInfo.Reference);

        try
        {
            using var workbook = new XLWorkbook(spreadSheetPath);

            bool nameExists = workbook.Worksheets.Any(ws => ws.Tables.Contains(tableInfo.Name));
            if (nameExists)
                throw new InvalidOperationException($"A table named '{tableInfo.Name}' already exists.");

            var worksheet = workbook.Worksheet(sheetName);

            (string startCol, uint startRow, string endCol, _) = CellReferenceParser.ParseRange(tableInfo.Reference);
            var columnLetters = CellReferenceParser.ExpandColumns(startCol, endCol);

            List<string> columnNames =
                tableInfo.Columns.Count == columnLetters.Count
                    ? tableInfo.Columns
                    : columnLetters.Select((col, i) => $"Column{i + 1}").ToList();

            for (int i = 0; i < columnLetters.Count; i++)
                worksheet.Cell($"{columnLetters[i]}{startRow}").Value = columnNames[i];

            var table = worksheet.Range(tableInfo.Reference).CreateTable(tableInfo.Name);
            table.ShowRowStripes = tableInfo.ShowRowStripes;
            table.EmphasizeFirstColumn = tableInfo.ShowFirstColumn;
            table.EmphasizeLastColumn = tableInfo.ShowLastColumn;
            table.ShowColumnStripes = tableInfo.ShowColumnStripes;

            if (!string.IsNullOrWhiteSpace(tableInfo.StyleName))
                table.Theme = XLTableTheme.FromName(tableInfo.StyleName);

            workbook.Save();
        }
        catch (Exception ex) when (ex is not InvalidOperationException)
        {
            throw new InvalidOperationException($"Error creating table: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Removes the named table from the specified sheet.
    /// </summary>
    public void DeleteTable(string spreadSheetPath, string sheetName, string tableName)
    {
        try
        {
            using var workbook = new XLWorkbook(spreadSheetPath);
            var worksheet = workbook.Worksheet(sheetName);

            if (!worksheet.Tables.Contains(tableName))
                throw new InvalidOperationException($"Table '{tableName}' not found on sheet '{sheetName}'.");

            worksheet.Tables.Remove(tableName);
            workbook.Save();
        }
        catch (Exception ex) when (ex is not InvalidOperationException)
        {
            throw new InvalidOperationException($"Error deleting table: {ex.Message}", ex);
        }
    }

    // ── Private helpers ──────────────────────────────────────────────────────────

    private static SerializableTable ReadSerializableTable(IXLTable table)
    {
        return new SerializableTable
        {
            Name = table.Name,
            Reference = table.RangeAddress.ToStringRelative(false),
            StyleName = table.Theme.Name == "None" ? null : table.Theme.Name,
            ShowFirstColumn = table.EmphasizeFirstColumn,
            ShowLastColumn = table.EmphasizeLastColumn,
            ShowRowStripes = table.ShowRowStripes,
            ShowColumnStripes = table.ShowColumnStripes,
            Columns = table.Fields.Select(f => f.Name).ToList(),
        };
    }
}
