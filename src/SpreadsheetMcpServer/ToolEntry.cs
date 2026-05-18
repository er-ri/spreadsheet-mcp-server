using System.ComponentModel;
using ModelContextProtocol.Server;
using SpreadsheetMcpServer.Core.Models;
using SpreadsheetMcpServer.Core.Services;

[McpServerToolType]
public static class ToolEntry
{
    private static readonly IWorksheetService _worksheetService = new WorksheetService();
    private static readonly ICellService _cellService = new CellService();
    private static readonly ITableService _tableService = new TableService();
    private static readonly ISpreadsheetExportService _exportService = new SpreadsheetExportService();

    [
        McpServerTool,
        Description(
            "Returns the base directory where spreadsheet files are located. "
                + "IMPORTANT: Call this first before using any other tool to resolve relative file names to absolute paths."
        )
    ]
    public static string GetWorkingDirectory() =>
        Environment.GetEnvironmentVariable("SPREADSHEET_BASE_PATH") ?? Directory.GetCurrentDirectory();

    [
        McpServerTool,
        Description(
            "Get all worksheet names from an Excel file. Prefix the spreadSheetPath with the path from GetWorkingDirectory."
        )
    ]
    public static List<string> GetAllWorksheets(string spreadSheetPath) =>
        _worksheetService.GetAllWorksheets(spreadSheetPath);

    [
        McpServerTool,
        Description("Read spread sheet. Prefix the spreadSheetPath with the path from GetWorkingDirectory.")
    ]
    public static List<SerializableCell> ReadSpreadSheet(string spreadSheetPath, string spreadSheetName) =>
        _cellService.ReadSpreadSheet(spreadSheetPath, spreadSheetName);

    [
        McpServerTool,
        Description(
            "Update the specified cells in an Excel file. "
                + "Prefix the spreadSheetPath with the path from GetWorkingDirectory. "
                + "Each cell's Address must be in the format 'SheetName!CellRef' (e.g. 'Sheet1!A1')."
        )
    ]
    public static void UpdateCells(string spreadSheetPath, List<SerializableCell> cells) =>
        _cellService.UpdateCells(spreadSheetPath, cells);

    [
        McpServerTool,
        Description(
            "Get all Excel tables defined on a worksheet. Prefix the filename with the path from GetWorkingDirectory."
        )
    ]
    public static List<SerializableTable> GetTables(string spreadSheetPath, string sheetName) =>
        _tableService.GetTables(spreadSheetPath, sheetName);

    [
        McpServerTool,
        Description(
            "Create an Excel table on the specified worksheet. "
                + "Prefix the spreadSheetPath with the path from GetWorkingDirectory. "
                + "The 'reference' field must be a range like 'A1:C5' where the first row contains headers. "
                + "Supply column names in 'columns' matching the header row. "
                + "Use 'styleName' for a built-in style (e.g. 'TableStyleMedium9'); omit for no style."
        )
    ]
    public static void CreateTable(string spreadSheetPath, string sheetName, SerializableTable table) =>
        _tableService.CreateTable(spreadSheetPath, table, sheetName);

    [
        McpServerTool,
        Description(
            "Delete a named Excel table from the specified worksheet. Prefix the spreadSheetPath with the path from GetWorkingDirectory."
        )
    ]
    public static void DeleteTable(string spreadSheetPath, string sheetName, string tableName) =>
        _tableService.DeleteTable(spreadSheetPath, sheetName, tableName);

    [
        McpServerTool,
        Description(
            "Export a JSON array (or single object) to a new worksheet in an Excel file. "
                + "Prefix the spreadSheetPath with the path from GetWorkingDirectory. "
                + "Creates the file if it does not exist. Throws if a sheet with the same name already exists."
        )
    ]
    public static void ExportJsonToSpreadSheet(string spreadSheetPath, string sheetName, string json) =>
        _exportService.ExportJsonArrayToSpreadSheet(spreadSheetPath, sheetName, json);
}
