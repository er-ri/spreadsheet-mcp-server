# Spreadsheet MCP Server

An MCP server that lets AI assistants read and edit Excel (`.xlsx`) files on your machine.

## Getting Started

Add the configuration snippet to your MCP client's settings file (e.g. `.mcp.json` for Claude Code).

```json
{
  "mcpServers": {
    "spreadsheet-mcp-server": {
      "type": "stdio",
      "command": "docker",
      "args": [
        "run", "--rm", "-i",
        "-v", "/path/to/your/folder:/data",
        "-e", "SPREADSHEET_BASE_PATH=/data",
        "erlizs/spreadsheet-mcp-server"
      ]
    }
  }
}
```

> Replace `/path/to/your/folder` with the folder containing your `.xlsx` files.

## Available Tools

| Tool | What it does |
|------|-------------|
| `GetWorkingDirectory` | Returns the base folder the server uses to resolve file paths. Call this first if you're unsure of the path. |
| `GetAllWorksheets` | Lists every worksheet in a workbook with its name, used range, and the anchor cells of any pictures it contains. |
| `ManageWorksheet` | Creates, deletes, or renames a worksheet. |
| `LoadRange` | Reads non-empty cells within a range as a two-column `Address \| Contents` Markdown table. Use it when you need values but not layout — empty cells are skipped entirely. |
| `LoadRangeInMarkdownTable` | Reads a range as a visual grid Markdown table (row numbers down the side, column letters across the top). Use it when spatial layout matters. |
| `FindStringsInSheets` | Searches worksheets for cells containing any of the given substrings and returns the matches per sheet. |
| `UpdateRange` | Writes a Markdown `Address \| Contents` table back into a worksheet — the inverse of `LoadRange`. Range addresses merge the region; `**bold**`, `*italic*` and `~~strike~~` become cell formatting. |
| `ManageTables` | Gets, creates, or deletes an Excel table on a worksheet. For 'create', converts a range into a named table with auto-filter and optional banding — write the data rows with `UpdateRange` first. |
| `ExportJsonToSpreadSheet` | Exports a JSON array (or single object) to a new worksheet. Creates the file if it does not exist; throws if a sheet with that name already exists. |
| `GetPictures` | Returns the first picture anchored within a range as an image the model can view directly. |
| `PastePictures` | Inserts one or more image files into a worksheet at given anchor cells, with optional pixel dimensions. |
| `ApplyCellFormatting` | Applies fonts, colors, alignment, borders, number formats, and row/column sizes to cells or ranges. |
| `ReadCellFormatting` | Reads a worksheet's styling back out in the same shape `ApplyCellFormatting` accepts, so styling can be inspected or copied. Default-styled cells are omitted and identically-styled neighbours are reported as one range. |

## License

Distributed under the MIT License. See `LICENSE` for more information.