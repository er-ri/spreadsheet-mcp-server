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
>
> The container runs as UID 1000 (non-root), so the mounted folder must be readable and
> writable by that user. If your host files are owned by a different UID, run `id -u` and
> `id -g` to find yours and pass them to `--user` as literal numbers. MCP clients launch
> `docker` directly rather than through a shell, so `$(id -u)` would **not** be expanded —
> substitute the actual values:

```json
{
  "mcpServers": {
    "spreadsheet-mcp-server": {
      "type": "stdio",
      "command": "docker",
      "args": [
        "run", "--rm", "-i",
        "-u", "1000:1000",
        "-v", "/path/to/your/folder:/data",
        "-e", "SPREADSHEET_BASE_PATH=/data",
        "erlizs/spreadsheet-mcp-server"
      ]
    }
  }
}
```

## Available Tools

| Tool | What it does |
|------|-------------|
| `GetWorkingDirectory` | Returns the base folder the server uses to resolve file paths. Call this first if you're unsure of the path. |
| `GetAllWorksheets` | Lists every worksheet in a workbook with its name, used range, and the anchor cells of any pictures it contains. |
| `ManageWorksheet` | Creates, deletes, or renames a worksheet. |
| `LoadRange` | Reads non-empty cells within a range as a two-column `Address \| Contents` Markdown table. Use it when you need values but not layout — empty cells are skipped entirely. |
| `LoadRangeInMarkdownTable` | Reads a range as a visual grid Markdown table (row numbers down the side, column letters across the top). Use it when spatial layout matters. |
| `FindStringsInSheets` | Searches worksheets for cells containing any of the given substrings and returns the matches per sheet. |
| `UpdateRange` | Writes a Markdown `Address \| Contents` table back into a worksheet — the inverse of `LoadRange`. Range addresses merge the region; `**bold**`, `*italic*` and `~~strike~~` become cell formatting. Unformatted values are written with their natural type: `=…` becomes a formula, numbers stay numeric, `true`/`false` become booleans, and ISO dates become date cells. |
| `ClearRange` | Removes the values of a range of cells — the inverse of `UpdateRange` for deletions. With `clearFormats`, wipes the cells completely instead: styling and number formats, but also merged regions, data validation and comments. |
| `ManageTables` | Gets, creates, or deletes an Excel table on a worksheet. For 'create', converts a range into a named table with auto-filter, sortable headers, and optional row/column banding (`showRowStripes`, `showFirstColumn`, `showLastColumn`, `showColumnStripes`) — write the data rows with `UpdateRange` first. |
| `GetPictures` | Returns the first picture anchored within a range as an image the model can view directly. |
| `PastePictures` | Inserts one or more image files into a worksheet at given anchor cells, with optional pixel dimensions. |
| `ApplyCellFormatting` | Applies fonts, colors, alignment, borders, number formats, and row/column sizes to cells or ranges. |
| `ReadCellFormatting` | Reads a worksheet's styling back out in the same shape `ApplyCellFormatting` accepts, so styling can be inspected or copied. Default-styled cells are omitted and identically-styled neighbours are reported as one range. |
| `ManageFreezePanes` | Sets or clears frozen panes on a worksheet. For 'set', freezes every row above and column left of the given anchor cell. |
| `ManageRowsColumns` | Inserts or deletes whole rows or columns on a worksheet, shifting subsequent cells to make room or fill the gap. |
| `ManageMerge` | Merges or unmerges a range of cells on a worksheet. |
| `AutofitRange` | Auto-sizes column widths and/or row heights within a range to fit their contents. |

## License

Distributed under the MIT License. See `LICENSE` for more information.
