# Overview

An MCP server built on C sharp.

## Architecture

This is a **Model Context Protocol (MCP) server** that exposes Excel/XLSX manipulation as MCP tools.

```
src/
  SpreadsheetMcpServer/         # Host: MCP server entry point
    Program.cs                  # Sets up MCP server with stdio transport
    ToolEntry.cs                # 17 MCP tools, discovered via reflection + attributes
  SpreadsheetMcpServer.Core/    # Library: business logic
    Services/                   # IWorksheetService, ICellService, ICellFormatService,
                                # ITableService, IPictureService + impls
    Models/                     # MarkdownSheet, CellFormatSpec, SerializableSheet,
                                # SerializableTable, PictureData, PicturePasteSpec (JSON DTOs)
    Helpers/                    # CellReferenceParser (expands ranges like A1:C5),
                                # WorkbookReader (shared load + effective-range logic),
                                # PathResolver (resolve + sandbox every file path)
tests/
  SpreadsheetMcpServer.Tests/   # xUnit tests against Core, plus ToolEntryPathSandboxTests
                                # (references the Host to assert the sandbox at the tool boundary)
```

**Data flow:** MCP client → stdio → `ToolEntry` → Services → ClosedXML → `.xlsx` files

**MCP tools exposed (17):**
`GetWorkingDirectory`, `GetAllWorksheets`, `ManageWorksheet`, `LoadRange`, `LoadRangeInMarkdownTable`, `FindStringsInSheets`, `UpdateRange`, `ClearRange`, `ManageTables`, `GetPictures`, `PastePictures`, `ApplyCellFormatting`, `ReadCellFormatting`, `ManageFreezePanes`, `ManageRowsColumns`, `ManageMerge`, `AutofitRange`
