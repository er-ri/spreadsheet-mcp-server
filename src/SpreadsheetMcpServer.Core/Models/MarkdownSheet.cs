namespace SpreadsheetMcpServer.Core.Models;

/// <summary>
/// A flattened, Markdown-oriented view of a worksheet range produced by
/// <see cref="Services.ICellService.LoadRange"/>.
/// </summary>
public class MarkdownSheet
{
    /// <summary>Name of the worksheet that was read.</summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// The tight bounding box of the cells actually present in <see cref="Contents"/> (e.g. "A1:B3").
    /// This is computed from the emitted rows after empty-cell filtering and truncation, so it reflects
    /// exactly what the table holds rather than the requested range. On early-return paths (no used
    /// range, an empty intersection, or no surviving rows) this echoes the requested range instead.
    /// </summary>
    public string TargetRange { get; set; } = string.Empty;

    /// <summary>
    /// A two-column Markdown table mapping each cell address to its text content, e.g.:
    /// <code>
    /// | Address | Contents |
    /// | --- | --- |
    /// | A1 | Name |
    /// | A2 | Gender |
    /// </code>
    /// The Address column holds a single cell address ("A1") or, for a merged region, its
    /// full range address such as "AD629:AL638". The Contents column holds the cell text
    /// rendered as Markdown, with inline run formatting converted to Markdown (bold →
    /// "**text**", italic → "*text*", strikethrough → "~~text~~"). Rows are ordered by
    /// spreadsheet position (row first, then column). Cells whose resolved value is an
    /// empty string are omitted entirely — no blank rows are emitted. Any literal "|" in a
    /// cell's content is escaped as "\|" so it does not break the table.
    /// </summary>
    public string Contents { get; set; } = string.Empty;
}
