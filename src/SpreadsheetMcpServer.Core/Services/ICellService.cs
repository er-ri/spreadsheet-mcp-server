using SpreadsheetMcpServer.Core.Models;

namespace SpreadsheetMcpServer.Core.Services;

/// <summary>
/// Service for reading and updating cells in Excel worksheets.
/// </summary>
public interface ICellService
{
    /// <summary>
    /// Reads non-empty cells from the specified worksheet within the given range and returns a
    /// flattened map of cell address → Markdown text. If the used range is smaller than
    /// <paramref name="range"/>, only the used range is read. Empty cells are excluded, inline run
    /// formatting is converted to Markdown, and merged regions are keyed by their full range address.
    /// Entries are ordered by spreadsheet position (row first, then column). When
    /// <paramref name="truncate"/> is positive the ordered entry list is truncated so its serialized
    /// JSON does not exceed that many characters. The returned <see cref="MarkdownSheet.TargetRange"/>
    /// is the tight bounding box of the entries that actually remain after filtering and truncation.
    /// </summary>
    MarkdownSheet LoadRange(string spreadSheetPath, string spreadSheetName, string range, int truncate = 0);

    /// <summary>
    /// Reads the cells of <paramref name="spreadSheetName"/> within <paramref name="range"/> and returns
    /// them as a visual grid Markdown table (the spatial counterpart to <see cref="LoadRange"/>): an
    /// empty first header column followed by the column letters, each body row prefixed with its
    /// spreadsheet row number, and each grid cell holding the value rendered as Markdown. Interior empty
    /// cells stay blank to preserve the layout; a merged region's value appears only in its anchor cell.
    /// The extent is the requested range intersected with the used range, with trailing fully-empty rows
    /// and columns dropped. When <paramref name="truncate"/> is positive, trailing rows are dropped so
    /// the rendered table stays within that many characters. <see cref="MarkdownSheet.TargetRange"/> is
    /// the tight bounding box actually emitted.
    /// </summary>
    MarkdownSheet LoadRangeInMarkdownTable(
        string spreadSheetPath,
        string spreadSheetName,
        string range,
        int truncate = 0
    );

    /// <summary>
    /// Searches worksheets in the workbook for cells whose text contains any element of
    /// <paramref name="searchStrings"/> (case-insensitive substring match against the plain cell value) and
    /// returns one <see cref="MarkdownSheet"/> per worksheet that has at least one match. When
    /// <paramref name="spreadSheetList"/> is non-empty only those named worksheets are searched;
    /// when it is empty or null every worksheet is searched. Each entry's
    /// <see cref="MarkdownSheet.SheetName"/> is the worksheet name and its
    /// <see cref="MarkdownSheet.Contents"/> is the same two-column Address/Contents Markdown table as
    /// <see cref="LoadRange"/> (matching cells only, Markdown-rendered, merged regions keyed by their
    /// full range address, rows ordered row-then-column). <see cref="MarkdownSheet.TargetRange"/> is left
    /// empty. Worksheets with no match are omitted; an empty or all-empty <paramref name="searchStrings"/> list
    /// yields an empty result.
    /// </summary>
    List<MarkdownSheet> FindStringsInSheets(
        List<string> searchStrings,
        string spreadSheetPath,
        List<string>? spreadSheetList = null
    );

    /// <summary>
    /// Writes a Markdown-formatted <see cref="MarkdownSheet"/> into the worksheet named by
    /// <see cref="MarkdownSheet.SheetName"/> — the inverse of <see cref="LoadRange"/>. Each
    /// <see cref="MarkdownSheet.Contents"/> key is a cell address (e.g. "A1") or a merged range
    /// address (e.g. "AD629:AL638"); range keys are merged and the value written to the anchor cell.
    /// Inline Markdown is converted to rich text: "**text**" → bold, "*text*" → italic,
    /// "~~text~~" → strikethrough. A value without Markdown styling is written as a typed cell —
    /// "=…" becomes a formula, and numbers, booleans and ISO dates keep their natural types.
    /// </summary>
    void UpdateRange(string spreadSheetPath, MarkdownSheet sheet);

    /// <summary>
    /// Removes the contents of every cell in <paramref name="range"/> of
    /// <paramref name="spreadSheetName"/>. If the requested range is larger than the sheet's used range,
    /// only the used range is cleared. When <paramref name="clearFormats"/> is true the cells are wiped
    /// completely instead — styling, borders and number formats, but also merged regions, data
    /// validation, conditional formats and comments — and cells carrying only styling count as used, so
    /// a format-only banner row inside the range is cleared too. Returns a confirmation message.
    /// </summary>
    string ClearRange(string spreadSheetPath, string spreadSheetName, string range, bool clearFormats = false);
}
