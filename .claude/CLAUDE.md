# Overview

An MCP server built on C sharp.

## Architecture

This is a **Model Context Protocol (MCP) server** that exposes Excel/XLSX manipulation as MCP tools.

```
src/
  SpreadsheetMcpServer/         # Host: MCP server entry point
    Program.cs                  # Sets up MCP server with stdio transport
    ToolEntry.cs                # 13 MCP tools, discovered via reflection + attributes
  SpreadsheetMcpServer.Core/    # Library: business logic
    Services/                   # IWorksheetService, ICellService, ICellFormatService,
                                # ITableService, IPictureService,
                                # ISpreadsheetExportService + impls
    Models/                     # MarkdownSheet, CellFormatSpec, SerializableSheet,
                                # SerializableTable, PictureData, PicturePasteSpec (JSON DTOs)
    Helpers/                    # CellReferenceParser (expands ranges like A1:C5),
                                # WorkbookReader (shared load + effective-range logic)
tests/
  SpreadsheetMcpServer.Tests/   # xUnit tests against Core only (not the Host)
```

**Data flow:** MCP client → stdio → `ToolEntry` → Services → ClosedXML → `.xlsx` files

**Key design points:**
- `ToolEntry` is a static class; MCP tools are discovered via `[McpServerTool]` attributes
- The server reads the `SPREADSHEET_BASE_PATH` environment variable to resolve relative file paths
- Tests use temporary files with cleanup (see `TestFixtureFactory`)
- `Directory.Build.props` centralizes `TargetFramework`, `ImplicitUsings`, and nullable settings for all projects
- Every read path goes through `WorkbookReader.SanitizePhoneticRuns`, which strips phonetic run (`<rPh>`) elements from shared strings before loading to work around a ClosedXML parsing bug
- `WorkbookReader.ComputeEffectiveRange` intersects a requested range with the used range; pass `XLCellsUsedOptions.AllFormats` to also count cells that carry only styling
- Cell values move as Markdown: `LoadRange` / `LoadRangeInMarkdownTable` read a `MarkdownSheet`, `UpdateRange` writes one back
- Cell *styling* moves as `CellFormatSpec`: `ReadCellFormatting` reads it, `ApplyCellFormatting` writes it, and the read output is valid input to the write — keep the two sides in sync when adding a formatting property

**MCP tools exposed (13):**
`GetWorkingDirectory`, `GetAllWorksheets`, `ManageWorksheet`, `LoadRange`, `LoadRangeInMarkdownTable`, `FindStringsInSheets`, `UpdateRange`, `ManageTables`, `ExportJsonToSpreadSheet`, `GetPictures`, `PastePictures`, `ApplyCellFormatting`, `ReadCellFormatting`
