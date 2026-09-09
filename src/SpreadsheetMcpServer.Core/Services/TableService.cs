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
    /// Gets, creates, or deletes a table on the specified worksheet. 'action' must be
    /// "get", "create", or "delete". 'table' is required for "create"; 'tableName' is
    /// required for "delete". Returns the tables remaining on the sheet after the operation.
    /// </summary>
    public List<SerializableTable> ManageTables(
        string spreadSheetPath,
        string sheetName,
        string action,
        string? tableName = null,
        SerializableTable? table = null
    )
    {
        try
        {
            using var workbook = new XLWorkbook(spreadSheetPath);
            var worksheet = WorkbookReader.GetWorksheet(workbook, sheetName);
            bool dirty = false;

            switch (action.Trim().ToLowerInvariant())
            {
                case "get":
                    break;

                case "create":
                    if (table is null)
                        throw new ArgumentException("'table' is required when action is 'create'.");
                    CreateTableInternal(workbook, worksheet, table);
                    dirty = true;
                    break;

                case "delete":
                    ArgumentException.ThrowIfNullOrWhiteSpace(tableName);
                    if (!worksheet.Tables.Contains(tableName))
                        throw new InvalidOperationException($"Table '{tableName}' not found on sheet '{sheetName}'.");
                    worksheet.Tables.Remove(tableName);
                    dirty = true;
                    break;

                default:
                    throw new InvalidOperationException(
                        $"Unknown action '{action}'. Expected 'get', 'create', or 'delete'."
                    );
            }

            if (dirty)
                workbook.Save();

            return worksheet.Tables.Select(ReadSerializableTable).ToList();
        }
        catch (Exception ex) when (ex is not InvalidOperationException and not ArgumentException)
        {
            throw new InvalidOperationException($"Error updating Excel file: {ex.Message}", ex);
        }
    }

    // ── Private helpers ──────────────────────────────────────────────────────────

    private static void CreateTableInternal(XLWorkbook workbook, IXLWorksheet worksheet, SerializableTable tableInfo)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(tableInfo.Name);
        ArgumentException.ThrowIfNullOrWhiteSpace(tableInfo.Reference);

        bool nameExists = workbook.Worksheets.Any(ws => ws.Tables.Contains(tableInfo.Name));
        if (nameExists)
            throw new InvalidOperationException($"A table named '{tableInfo.Name}' already exists.");

        (string startCol, int startRow, string endCol, _) = CellReferenceParser.ParseRange(tableInfo.Reference);
        var columnLetters = CellReferenceParser.ExpandColumns(startCol, endCol);

        List<string> columnNames =
            tableInfo.Columns.Count == columnLetters.Count
                ? tableInfo.Columns
                : columnLetters.Select((col, i) => $"Column{i + 1}").ToList();

        for (int i = 0; i < columnLetters.Count; i++)
            worksheet.Cell($"{columnLetters[i]}{startRow}").Value = columnNames[i];

        var xlTable = worksheet.Range(tableInfo.Reference).CreateTable(tableInfo.Name);
        xlTable.ShowRowStripes = tableInfo.ShowRowStripes;
        xlTable.EmphasizeFirstColumn = tableInfo.ShowFirstColumn;
        xlTable.EmphasizeLastColumn = tableInfo.ShowLastColumn;
        xlTable.ShowColumnStripes = tableInfo.ShowColumnStripes;

        if (!string.IsNullOrWhiteSpace(tableInfo.StyleName))
            xlTable.Theme = XLTableTheme.FromName(tableInfo.StyleName);
    }

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
