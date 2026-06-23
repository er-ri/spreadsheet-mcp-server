# Overview

An MCP server built on C sharp.

## Architecture

This is a **Model Context Protocol (MCP) server** that exposes Excel/XLSX manipulation as MCP tools.

```
src/
  SpreadsheetMcpServer/         # Host: MCP server entry point
    Program.cs                  # Sets up MCP server with stdio transport
    ToolEntry.cs                # 9 MCP tools, discovered via reflection + attributes
  SpreadsheetMcpServer.Core/    # Library: business logic
    Services/                   # IWorksheetService, ICellService, ITableService,
                                # ISpreadsheetExportService + impls
    Models/                     # SerializableCell, SerializableCellFormat,
                                # SerializableSheet, SerializableTable (JSON DTOs)
    Helpers/                    # CellReferenceParser (expands ranges like A1:C5)
tests/
  SpreadsheetMcpServer.Tests/   # xUnit tests against Core only (not the Host)
```

**Data flow:** MCP client → stdio → `ToolEntry` → Services → ClosedXML → `.xlsx` files

**Key design points:**
- `ToolEntry` is a static class; MCP tools are discovered via `[McpServerTool]` attributes
- The server reads the `SPREADSHEET_BASE_PATH` environment variable to resolve relative file paths
- Tests use temporary files with cleanup (see `TestFixtureFactory`)
- `Directory.Build.props` centralizes `TargetFramework`, `ImplicitUsings`, and nullable settings for all projects
- `LoadRange` strips phonetic run (`<rPh>`) elements from shared strings before loading to work around a ClosedXML parsing bug
- `SerializableCellFormat` carries full visual formatting (font, colors, alignment, borders); read via `LoadRange`, written via `UpdateRange`
- `SerializableSheet` includes `MergedRanges` (all merged regions intersecting the loaded range) and per-cell `MergeAddress` on anchor cells

**MCP tools exposed (13):**
`GetWorkingDirectory`, `GetAllWorksheets`, `ManageWorksheet`, `LoadRange`, `LoadRangeInMarkdownTable`, `FindStringsInSheets`, `UpdateRange`, `GetTables`, `CreateTable`, `DeleteTable`, `ExportJsonToSpreadSheet`, `GetPictures`, `PastePictures`
