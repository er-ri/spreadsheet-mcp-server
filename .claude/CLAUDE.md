# Overview

An MCP server built on C sharp.

## Architecture

This is a **Model Context Protocol (MCP) server** that exposes Excel/XLSX manipulation as MCP tools.

```
src/
  SpreadsheetMcpServer/         # Host: MCP server entry point
    Program.cs                  # Sets up MCP server with stdio transport
    ToolEntry.cs                # 14 MCP tools, discovered via reflection + attributes
  SpreadsheetMcpServer.Core/    # Library: business logic
    Services/                   # IWorksheetService, ICellService, ICellFormatService,
                                # ITableService, IPictureService,
                                # ISpreadsheetExportService + impls
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

**Key design points:**
- `ToolEntry` is a static class; MCP tools are discovered via `[McpServerTool]` attributes
- The server reads the `SPREADSHEET_BASE_PATH` environment variable to resolve relative file paths
- Every `spreadSheetPath` (and the picture path in `PastePictures`) is routed through `PathResolver.Resolve`, which canonicalizes the path against `SPREADSHEET_BASE_PATH` and rejects anything that would escape it with `../` or an outside absolute path. Keep all file entry points inside this resolver. The check is lexical and does not follow symlinks — the base directory's own contents are trusted, since the operator chooses what to expose; the resolver guards the *path argument*, not the directory
- The sandbox is applied at the tool boundary, not in the services, so `ToolEntryPathSandboxTests` enumerates `[McpServerTool]` methods by reflection and asserts each one with a `spreadSheetPath` rejects an escaping path. A new tool that forgets `PathResolver.Resolve` fails there; keep its expected-tool-count assertion in step with the tool list
- Tests use temporary files with cleanup (see `TestFixtureFactory`). `SPREADSHEET_BASE_PATH` is process-wide state, so `BasePathFixture` owns it and every test class touching a spreadsheet path joins `[Collection(BasePathCollectionDefinition.Name)]` — that is what serializes them, rather than it being an accident of living in one class
- `Directory.Build.props` centralizes `TargetFramework`, `ImplicitUsings`, nullable settings, and (since the analyzer gate) `TreatWarningsAsErrors` for all projects
- Every read path goes through `WorkbookReader.SanitizePhoneticRuns`, which strips phonetic run (`<rPh>`) elements from shared strings before loading to work around a ClosedXML parsing bug
- `WorkbookReader.ComputeEffectiveRange` intersects a requested range with the used range; pass `XLCellsUsedOptions.AllFormats` to also count cells that carry only styling (`ClearRange` passes `All` when clearing formats, so a styling-only cell counts as used and is actually cleared)
- Look a worksheet up with `WorkbookReader.GetWorksheet`, never ClosedXML's `workbook.Worksheet(name)` — the latter throws `ArgumentException`, which the catch filters below deliberately let through, so a missing sheet would escape unwrapped. `TryGetWorksheet` stays for existence checks (create/delete/rename, duplicate detection)
- Cell values move as Markdown: `LoadRange` / `LoadRangeInMarkdownTable` read a `MarkdownSheet`, `UpdateRange` writes one back. An unformatted value is written as a **typed** cell (`=…` → formula, numeric → number, boolean, ISO date); Markdown styling switches a cell to rich text (`SetTypedValue` in `CellService`)
- The read/write pair must stay inverse: `CellService.RenderValue` emits dates in an ISO format from `SupportedDateFormats` and prefixes a *text* cell with `'` whenever `WouldBeTyped` says `SetTypedValue` would re-type it, so `LoadRange` → `UpdateRange` leaves `007`, `true` and `2026-09-07` untouched. `WouldBeTyped` and `SetTypedValue`'s branches have to be kept in step
- Cell *styling* moves as `CellFormatSpec`: `ReadCellFormatting` reads it, `ApplyCellFormatting` writes it, and the read output is valid input to the write — keep the two sides in sync when adding a formatting property
- Exception handling convention: services wrap unexpected exceptions in `InvalidOperationException` via `catch (Exception ex) when (ex is not InvalidOperationException and not ArgumentException)`; expected conditions ("worksheet not found", "table already exists") use explicit `InvalidOperationException`/`ArgumentException` throws
- `ToolEntry` is the composition root: it names the concrete services but holds them through Core's interfaces, so those interfaces stay the seam Core is consumed through. CA1859 is suppressed for that file in `.editorconfig` (as it is under `tests/**`) rather than typing the fields concretely
- `Directory.Build.props` pins `<AnalysisLevel>` to an explicit version rather than `latest`: `global.json` rolls the SDK forward, and with `TreatWarningsAsErrors` on, `latest` lets a newer runner fail the build on an unmodified commit. Bump it deliberately

**MCP tools exposed (14):**
`GetWorkingDirectory`, `GetAllWorksheets`, `ManageWorksheet`, `LoadRange`, `LoadRangeInMarkdownTable`, `FindStringsInSheets`, `UpdateRange`, `ClearRange`, `ManageTables`, `ExportJsonToSpreadSheet`, `GetPictures`, `PastePictures`, `ApplyCellFormatting`, `ReadCellFormatting`
