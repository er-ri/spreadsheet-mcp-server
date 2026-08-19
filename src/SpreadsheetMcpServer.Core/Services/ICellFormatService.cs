using SpreadsheetMcpServer.Core.Models;

namespace SpreadsheetMcpServer.Core.Services;

/// <summary>
/// Service for reading and applying cell formatting.
/// </summary>
public interface ICellFormatService
{
    void ApplyCellFormatting(string spreadSheetPath, List<CellFormatSpec> formats);

    /// <summary>
    /// Reads the visual styling of <paramref name="spreadSheetName"/> within <paramref name="range"/> and
    /// returns it as <see cref="CellFormatSpec"/> entries — the inverse of
    /// <see cref="ApplyCellFormatting"/>, whose <c>formats</c> argument accepts the result unchanged.
    /// The requested range is intersected with the used range (counting cells that carry only styling,
    /// not just contents). Cells whose style matches the workbook default are omitted entirely, and each
    /// spec carries only the properties that differ from that default. Adjacent cells sharing identical
    /// styling are coalesced into a single rectangular range entry (e.g. "A1:D1"). Row heights and column
    /// widths that differ from the sheet default are emitted first, as dimension-only entries addressed to
    /// a cell in the affected row/column. Addresses have no sheet prefix. When <paramref name="truncate"/>
    /// is positive the returned list is capped at that many entries.
    /// </summary>
    /// <remarks>
    /// Known fidelity limits, all bounded by what <see cref="CellFormatSpec"/> can express: a double or
    /// accounting underline reads back as <c>true</c> and re-applies as a single underline; theme-based
    /// colors have no literal RGB and are omitted; number formats stored as a built-in format id rather
    /// than a format string are omitted.
    /// </remarks>
    List<CellFormatSpec> ReadCellFormatting(
        string spreadSheetPath,
        string spreadSheetName,
        string range,
        int truncate = 200
    );
}
