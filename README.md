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
| `GetAllWorksheets` | Lists all sheet names in a workbook. |
| `ReadSpreadSheet` | Reads every cell in a sheet — returns address, value, data type, and formula. |
| `UpdateCells` | Writes values to one or more cells. Addresses must be in `SheetName!CellRef` format (e.g. `Sheet1!A1`). |
| `GetTables` | Lists all Excel tables on a sheet. |
| `CreateTable` | Creates a formatted Excel table from a cell range. The first row must contain headers. |
| `DeleteTable` | Removes a named Excel table from a sheet. |

## License

Distributed under the MIT License. See `LICENSE` for more information.