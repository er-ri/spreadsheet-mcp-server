using System.ComponentModel;
using ModelContextProtocol.Protocol;
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
    private static readonly IPictureService _pictureService = new PictureService();
    private static readonly ICellFormatService _cellFormatService = new CellFormatService();

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
            "Get all worksheets from an Excel file, each with its name, used range, and the cell addresses of any pictures it contains."
        )
    ]
    public static List<SerializableSheet> GetAllWorksheets(string spreadSheetPath) =>
        _worksheetService.GetAllWorksheets(spreadSheetPath);

    [
        McpServerTool,
        Description(
            "Create, delete, or rename a worksheet in an Excel file. "
                + "'action' must be 'create', 'delete', or 'rename'. "
                + "For 'create', a new empty worksheet named spreadSheetName is added. "
                + "For 'delete', the worksheet named spreadSheetName is removed. "
                + "For 'rename', the worksheet named spreadSheetName is renamed to newSheetName (newSheetName is required). "
                + "Returns a confirmation message."
        )
    ]
    public static string ManageWorksheet(
        string spreadSheetPath,
        string spreadSheetName,
        string action,
        string? newSheetName = null
    ) => _worksheetService.ManageWorksheet(spreadSheetPath, spreadSheetName, action, newSheetName);

    [
        McpServerTool,
        Description(
            "Read cells from a worksheet within the specified range. "
                + "Prefer this tool over LoadRangeInMarkdownTable when you only need cell values and do not care about spatial layout — it omits empty cells entirely, making it more token-efficient for sparse sheets. "
                + "The range parameter is required (e.g. 'A1:Q60'). "
                + "If the sheet's used range is smaller than the requested range, only the used range is read. "
                + "'Contents' is a two-column Markdown table ('| Address | Contents |') with one row per non-empty cell: the Address column is a cell address (or, for merged regions, the full range address such as 'AD629:AL638') and the Contents column is the cell text rendered as Markdown (bold → **text**, italic → *text*, strikethrough → ~~text~~). Rows are ordered by spreadsheet position (row first, then column); any literal '|' in a value is escaped as '\\|'. "
                + "Empty cells are excluded from the result (no blank rows). "
                + "Use 'truncate' to cap the output table size (in characters); when set, the last rows in spreadsheet order are dropped to fit within the limit."
        )
    ]
    public static MarkdownSheet LoadRange(
        string spreadSheetPath,
        string spreadSheetName,
        string range,
        int truncate = 10000
    ) => _cellService.LoadRange(spreadSheetPath, spreadSheetName, range, truncate);

    [
        McpServerTool,
        Description(
            "Read cells from a worksheet within the specified range, laid out as a visual grid Markdown table (the spatial counterpart to LoadRange). "
                + "Use this tool instead of LoadRange when spatial layout matters — for example, when reading a form, understanding column alignment, or diagnosing merged cell structure."
                + "The range parameter is required (e.g. 'A1:Q60'). "
                + "If the sheet's used range is smaller than the requested range, only the used range is read. "
                + "'Contents' is a grid Markdown table: the first header column is empty and the remaining headers are the column letters (A, B, C…); each body row is prefixed with its spreadsheet row number (1, 2, 3…), and each grid cell holds the cell value rendered as Markdown (bold → **text**, italic → *text*, strikethrough → ~~text~~). A literal '|' is escaped as '\\|' and newlines are collapsed to spaces. "
                + "Interior empty cells are kept as blank grid cells to preserve the layout; trailing fully-empty rows and columns are dropped. A merged region's value appears only in its top-left anchor cell, with the covered cells left blank. "
                + "Use 'truncate' to cap the output table size (in characters); when set, the last rows are dropped to fit within the limit. "
                + "'TargetRange' reports the tight bounding box actually emitted."
        )
    ]
    public static MarkdownSheet LoadRangeInMarkdownTable(
        string spreadSheetPath,
        string spreadSheetName,
        string range,
        int truncate = 10000
    ) => _cellService.LoadRangeInMarkdownTable(spreadSheetPath, spreadSheetName, range, truncate);

    [
        McpServerTool,
        Description(
            "Search worksheets of a workbook for cells whose text contains any of the given strings. "
                + "'searchStrings' is a list of substrings; a cell matches if its text contains any of them (case-insensitive). "
                + "'spreadSheetList' restricts the search to the named worksheets; leave it empty to search every worksheet. "
                + "Returns one entry per worksheet that has at least one match; sheets with no match are omitted. "
                + "Each entry's 'Contents' is the same two-column Markdown table as LoadRange ('| Address | Contents |'), "
                + "with 'SheetName' set to the worksheet name and 'TargetRange' left empty."
        )
    ]
    public static List<MarkdownSheet> FindStringsInSheets(
        List<string> searchStrings,
        string spreadSheetPath,
        List<string>? spreadSheetList = null
    ) => _cellService.FindStringsInSheets(searchStrings, spreadSheetPath, spreadSheetList);

    [
        McpServerTool,
        Description(
            "Write Markdown-formatted text into a worksheet — the inverse of LoadRange. "
                + "Pass a sheet object with 'SheetName' and 'Contents', where 'Contents' is a two-column "
                + "Markdown table ('| Address | Contents |', with a '| --- | --- |' separator row) and one "
                + "row per cell. An Address is either a single cell ('A1') or a merged range address "
                + "('AD629:AL638'); range addresses merge the region and write the value into its "
                + "top-left cell. Escape any literal '|' in a value as '\\|'. Inline Markdown is converted "
                + "to cell formatting: '**text**' → bold, '*text*' → italic, '~~text~~' → strikethrough."
        )
    ]
    public static void UpdateRange(string spreadSheetPath, MarkdownSheet sheet) =>
        _cellService.UpdateRange(spreadSheetPath, sheet);

    [
        McpServerTool,
        Description(
            "Get, create, or delete an Excel table on a worksheet. "
                + "'action' must be 'get', 'create', or 'delete'. "
                + "For 'get', all tables defined on sheetName are returned; tableName and table are ignored. "
                + "For 'create', 'table' is required — it converts a cell range into a named Excel Table with auto-filter dropdowns, sortable headers, and optional row banding. "
                + "Use this whenever you are writing structured tabular data (a header row plus one or more data rows) to a spreadsheet — prefer it over writing plain cells with UpdateRange for any grid-shaped dataset. "
                + "Workflow: (1) write the data rows (not the header row) into the range first with UpdateRange, then (2) call ManageTables with action 'create' — it writes the column names from 'table.columns' into the first row of 'table.reference' automatically. "
                + "'table.reference' must be a full range covering headers and data, e.g. 'A1:C5'. Use 'table.styleName' for a built-in table style (e.g. 'TableStyleMedium9'); omit for no style. "
                + "For 'delete', 'tableName' is required — the named table is removed from sheetName. "
                + "Returns the list of tables remaining on sheetName after the operation."
        )
    ]
    public static List<SerializableTable> ManageTables(
        string spreadSheetPath,
        string sheetName,
        string action,
        string? tableName = null,
        SerializableTable? table = null
    ) => _tableService.ManageTables(spreadSheetPath, sheetName, action, tableName, table);

    [
        McpServerTool,
        Description(
            "Export a JSON array (or single object) to a new worksheet in an Excel file. "
                + "Creates the file if it does not exist. Throws if a sheet with the same name already exists."
        )
    ]
    public static void ExportJsonToSpreadSheet(string spreadSheetPath, string sheetName, string json) =>
        _exportService.ExportJsonArrayToSpreadSheet(spreadSheetPath, sheetName, json);

    [
        McpServerTool,
        Description(
            "Return the first picture whose anchor (top-left) cell falls within the given range, as an image the model can view directly. "
                + "The range parameter is required (e.g. 'A1:C5'). "
                + "'First' follows the worksheet's picture order. "
                + "Returns null when no picture is anchored within the range."
        )
    ]
    public static ImageContentBlock? GetPictures(string spreadSheetPath, string spreadSheetName, string range)
    {
        var picture = _pictureService.GetFirstPictureInRange(spreadSheetPath, spreadSheetName, range);
        return picture is null
            ? null
            : new ImageContentBlock { Data = Convert.FromBase64String(picture.Base64), MimeType = picture.MimeType };
    }

    [
        McpServerTool,
        Description(
            "Insert one or more pictures from image files into a worksheet. The inverse of GetPictures. "
                + "'pictures' is a list; each item has 'picturePath' (the source image file, resolved relative to GetWorkingDirectory like spreadSheetPath), "
                + "'targetSheet' (the worksheet to insert into), 'topLeftAddress' (the anchor cell, e.g. 'B2'), "
                + "and optional 'width'/'height' in pixels (the original image size is used when a dimension is omitted). "
                + "All pictures are written to the single workbook at spreadSheetPath. "
                + "Returns a confirmation message."
        )
    ]
    public static string PastePictures(string spreadSheetPath, List<PicturePasteSpec> pictures) =>
        _pictureService.PastePictures(spreadSheetPath, pictures);

    [
        McpServerTool,
        Description(
            "Apply formatting to cells or ranges in a worksheet. "
                + "'formats' is a list of CellFormatSpec objects, each specifying an address and formatting options. "
                + "Address can be a single cell ('A1'), a range ('A1:B5'), or include sheet name ('Sheet1!A1:B5'). "
                + "Supported formatting: "
                + "- Font: bold, italic, underline, strikethrough, fontColor (hex), fontName, fontSize; "
                + "- Fill: backgroundColor (hex); "
                + "- Alignment: horizontalAlignment ('Left', 'Center', 'Right', 'Justify'), verticalAlignment ('Top', 'Center', 'Bottom'), wrapText; "
                + "- Borders: topBorder, bottomBorder, leftBorder, rightBorder (each with style and color); "
                + "- Number format: numberFormat (e.g., '0.00', 'yyyy-mm-dd', '#,##0'); "
                + "- Dimensions: height (row height in points), width (column width in characters). "
                + "Returns a confirmation message."
        )
    ]
    public static string ApplyCellFormatting(string spreadSheetPath, List<CellFormatSpec> formats)
    {
        _cellFormatService.ApplyCellFormatting(spreadSheetPath, formats);
        return $"Applied formatting to {formats.Count} cell(s)/range(s).";
    }
}
