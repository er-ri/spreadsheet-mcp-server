# Contributing

## Prerequisites

- [.NET 10 SDK](https://dotnet.microsoft.com/download)
- Docker (optional, for container builds)

## Getting started

```bash
git clone https://github.com/er-ri/spreadsheet-mcp-server
cd spreadsheet-mcp-server
dotnet restore
dotnet build
dotnet test
```

## Project layout

```
src/
  SpreadsheetMcpServer/         # Host — MCP server entry point (stdio transport)
  SpreadsheetMcpServer.Core/    # Library — all business logic (ClosedXML, MiniExcel)
tests/
  SpreadsheetMcpServer.Tests/   # xUnit tests against Core only
```

`Directory.Build.props` at the root sets `TargetFramework`, `ImplicitUsings`, and `Nullable` for every project so you don't repeat them in individual `.csproj` files.

## Development workflow

```bash
dotnet build                    # compile everything
dotnet test                     # run all tests
dotnet test --logger "console;verbosity=detailed"   # verbose output
```

Run the server locally (used by the `spreadsheet-mcp-server` entry in `.mcp.json`):

```bash
dotnet run --project src/SpreadsheetMcpServer/SpreadsheetMcpServer.csproj
```

Set `SPREADSHEET_BASE_PATH` to control which directory the server resolves relative file paths against. Without it, the working directory is used.

## Adding a new MCP tool

1. Add the interface method and implementation in `SpreadsheetMcpServer.Core/Services/`.
2. Inject the service in `ToolEntry.cs` (static field at the top of the class).
3. Add a public static method to `ToolEntry` decorated with `[McpServerTool]` and `[Description("...")]`.
4. Write xUnit tests in `SpreadsheetMcpServer.Tests` using `TestFixtureFactory` to create temp workbooks.
5. Update the tool count and list in `CLAUDE.md` and `README.md`.

## Code style

The project uses [CSharpier](https://csharpier.com/) for formatting (configured in `.csharpierrc.json`). Run it before committing:

```bash
dotnet tool run csharpier format .
```

Nullable reference types are enabled project-wide — keep all warnings clean.

## Testing approach

Tests live in `SpreadsheetMcpServer.Tests` and cover `Core` services only (not the Host). Each test creates a temporary `.xlsx` file via `TestFixtureFactory`, exercises the service, then deletes the file in a `finally` block. No mocks — tests hit real ClosedXML workbooks.

## Building and publishing the Docker image

```bash
# Build image locally
dotnet publish src/SpreadsheetMcpServer/SpreadsheetMcpServer.csproj \
  -c Release -r linux-x64 --self-contained false \
  -t:PublishContainer \
  -p:ContainerRepository=erlizs/spreadsheet-mcp-server

# Push to Docker Hub
dotnet publish src/SpreadsheetMcpServer/SpreadsheetMcpServer.csproj \
  -c Release -r linux-x64 --self-contained false \
  -t:PublishContainer \
  -p:ContainerRegistry=docker.io \
  -p:ContainerRepository=erlizs/spreadsheet-mcp-server
```
