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
| `GetAllWorksheets` | Lists all sheet names in a workbook, each with its used range. |
| `LoadRange` | Reads non-empty cells from a worksheet within a range (default `A1:Q60`). Returns address, value, data type, formula, formatting, and merged-cell info. Accepts a `threshold` to cap JSON output size. |
| `UpdateRange` | Writes values, formulas, or formatting to one or more cells (bare cell refs, e.g. `A1`). Supports typed values (Number, Boolean, Date, text), formulas, and full cell formatting. Set both Value and Formula to null to clear a cell. |
| `ManageTables` | Gets, creates, or deletes an Excel table on a worksheet. For 'create', converts a cell range into a named Excel table with auto-filter, sortable headers, and optional row banding — write data rows with `UpdateRange` first, then call this to add headers and table formatting. |
| `ExportJsonToSpreadSheet` | Exports a JSON array (or single object) to a new worksheet in an Excel file. Creates the file if it does not exist; throws if a sheet with that name already exists. |

## License

Distributed under the MIT License. See `LICENSE` for more information.