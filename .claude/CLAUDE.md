# Overview

An MCP server built on C sharp.

## Architecture

This is a **Model Context Protocol (MCP) server** that exposes Excel/XLSX manipulation as MCP tools.

```
src/
  SpreadsheetMcpServer/         # Host: MCP server entry point
    Program.cs                  # Sets up MCP server with stdio transport
    ToolEntry.cs                # 8 MCP tools, discovered via reflection + attributes
  SpreadsheetMcpServer.Core/    # Library: business logic
    Services/                   # IWorksheetService, ICellService, ITableService,
                                # ISpreadsheetExportService + impls
    Models/                     # SerializableCell, SerializableTable (JSON DTOs)
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

**MCP tools exposed (8):**
`GetWorkingDirectory`, `ReadSpreadSheet`, `GetAllWorksheets`, `UpdateCells`, `GetTables`, `CreateTable`, `DeleteTable`, `ExportJsonToSpreadSheet`

## Local MCP Configuration

`.mcp.json` registers two server instances for development:
- `spreadsheet-mcp-server` — runs via `dotnet run`
- `spreadsheet-mcp-server-docker` — runs via Docker with a volume mount to `/data`
