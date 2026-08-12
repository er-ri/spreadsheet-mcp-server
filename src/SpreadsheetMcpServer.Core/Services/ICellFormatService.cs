using SpreadsheetMcpServer.Core.Models;

namespace SpreadsheetMcpServer.Core.Services;

/// <summary>
/// Service for applying cell formatting.
/// </summary>
public interface ICellFormatService
{
    void ApplyCellFormatting(string spreadSheetPath, List<CellFormatSpec> formats);
}
