using System.Text;
using ClosedXML.Excel;
using SpreadsheetMcpServer.Core.Helpers;
using SpreadsheetMcpServer.Core.Models;

namespace SpreadsheetMcpServer.Core.Services;

/// <summary>
/// Service for reading and updating cells in Excel worksheets.
/// </summary>
public class CellService : ICellService
{
    /// <summary>
    /// Reads non-empty cells from the specified worksheet within the given range and returns them as a
    /// two-column Address/Contents Markdown table (see <see cref="MarkdownSheet.Contents"/>). If the
    /// used range is smaller than <paramref name="range"/>, only the used range is read. Cells whose
    /// value is an empty string are omitted — no blank rows are emitted. Inline run formatting is
    /// converted to Markdown (bold → "**text**", italic → "*text*", strikethrough → "~~text~~").
    /// Merged regions are emitted as a single row keyed by the full merged range address
    /// (e.g. "AD629:AL638") using the anchor cell's value. Rows are ordered by spreadsheet position
    /// (row first, then column), with merged-range keys ordered by their anchor (top-left) address.
    /// When <paramref name="truncate"/> is positive the table is truncated so its rendered length does
    /// not exceed that many characters (dropping the last rows in spreadsheet order).
    /// <see cref="MarkdownSheet.TargetRange"/> reports the tight bounding box of the rows that actually
    /// remain in the table (after empty-cell filtering and truncation), not the requested range.
    /// </summary>
    public MarkdownSheet LoadRange(string spreadSheetPath, string spreadSheetName, string range, int truncate = 0)
    {
        try
        {
            using var sanitized = WorkbookReader.SanitizePhoneticRuns(spreadSheetPath);
            using var workbook = new XLWorkbook(sanitized);
            var worksheet = workbook.Worksheet(spreadSheetName);
            var contents = new Dictionary<string, string>();

            // Determine the effective read range: intersect the requested range with the used range.
            var bounds = WorkbookReader.ComputeEffectiveRange(worksheet, range);
            if (bounds == null)
                return new MarkdownSheet
                {
                    SheetName = spreadSheetName,
                    TargetRange = range,
                    Contents = string.Empty,
                };

            var (firstRow, firstCol, lastRow, lastCol) = bounds.Value;
            var effectiveRange = worksheet.Range(firstRow, firstCol, lastRow, lastCol);

            // Build a lookup from anchor address → merged range address for all merges that
            // intersect the effective range. The value of a merged region is taken from its
            // anchor (top-left) cell and emitted once under the full range address; the member
            // cells are skipped.
            var (mergedRangeByAnchor, mergedMemberCells) = BuildMergeMaps(
                worksheet,
                firstRow,
                firstCol,
                lastRow,
                lastCol
            );

            foreach (var cell in effectiveRange.CellsUsed())
            {
                // Skip blank cells to improve efficiency.
                if (cell.DataType == XLDataType.Blank)
                    continue;

                string? value = cell.Value.ToString();
                if (string.IsNullOrEmpty(value))
                    continue;

                string cellRef = cell.Address.ToString()!;

                // For a merged region, emit a single entry under the full range address using the
                // anchor cell; skip non-anchor members entirely.
                if (mergedMemberCells.Contains(cellRef) && !mergedRangeByAnchor.ContainsKey(cellRef))
                    continue;

                string key = mergedRangeByAnchor.TryGetValue(cellRef, out string? mergeAddr) ? mergeAddr : cellRef;
                contents[key] = ToMarkdown(cell);
            }

            // Order entries by spreadsheet position — row first, then column — using a natural
            // cell-address ordering (so A2 sorts before A10, and A1 before B1). A merged-range key
            // (e.g. "AD629:AL638") sorts by its anchor (top-left) address.
            var orderedKeys = contents
                .Keys.OrderBy(k => CellReferenceParser.GetRowIndex(AnchorAddress(k)))
                .ThenBy(k =>
                    CellReferenceParser.ColumnLetterToIndex(CellReferenceParser.GetColumnName(AnchorAddress(k)))
                )
                .ToList();

            // Select the rows that survive truncation, then derive both the rendered table and the
            // reported range from that same set so TargetRange reflects exactly what Contents holds.
            var selectedRows = SelectRows(orderedKeys.Select(k => (k, contents[k])), truncate);

            return new MarkdownSheet
            {
                SheetName = spreadSheetName,
                TargetRange = BoundingBox(selectedRows, range),
                Contents = RenderTable(selectedRows),
            };
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Error reading Excel file: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Searches worksheets in the workbook for cells whose plain text contains any element of
    /// <paramref name="searchStrings"/> (case-insensitive substring) and returns one <see cref="MarkdownSheet"/>
    /// per worksheet that has at least one match. When <paramref name="spreadSheetList"/> is non-empty only
    /// those named worksheets are searched; when it is empty or null every worksheet is searched. Each
    /// entry reuses the same Address/Contents rendering as <see cref="LoadRange"/> — matching cells only,
    /// ordered by spreadsheet position, merged regions keyed by their full range address — with
    /// <see cref="MarkdownSheet.SheetName"/> set to the worksheet name and
    /// <see cref="MarkdownSheet.TargetRange"/> left empty. Worksheets with no match are omitted; an empty
    /// or all-empty <paramref name="searchStrings"/> list yields an empty result. Matching is done against the
    /// plain cell value so Markdown formatting markers never interfere with the search.
    /// </summary>
    public List<MarkdownSheet> FindStringsInSheets(
        List<string> searchStrings,
        string spreadSheetPath,
        List<string>? spreadSheetList = null
    )
    {
        // Normalize the search needles once; no usable needles means nothing can match.
        var needles = (searchStrings ?? new List<string>()).Where(s => !string.IsNullOrEmpty(s)).ToList();
        if (needles.Count == 0)
            return new List<MarkdownSheet>();

        // An empty/null sheet filter means "search every worksheet"; otherwise restrict to the named set.
        var sheetFilter = (spreadSheetList ?? new List<string>())
            .Where(s => !string.IsNullOrEmpty(s))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        try
        {
            using var sanitized = WorkbookReader.SanitizePhoneticRuns(spreadSheetPath);
            using var workbook = new XLWorkbook(sanitized);

            var results = new List<MarkdownSheet>();
            foreach (var worksheet in workbook.Worksheets)
            {
                if (sheetFilter.Count > 0 && !sheetFilter.Contains(worksheet.Name))
                    continue;

                var usedRangeAddress = worksheet.RangeUsed()?.RangeAddress;
                if (usedRangeAddress == null)
                    continue;

                int firstRow = usedRangeAddress.FirstAddress.RowNumber;
                int firstCol = usedRangeAddress.FirstAddress.ColumnNumber;
                int lastRow = usedRangeAddress.LastAddress.RowNumber;
                int lastCol = usedRangeAddress.LastAddress.ColumnNumber;
                var usedRange = worksheet.Range(firstRow, firstCol, lastRow, lastCol);

                // Same merge handling as LoadRange: a merged region is emitted once under its full
                // range address using the anchor cell; non-anchor members are skipped.
                var (mergedRangeByAnchor, mergedMemberCells) = BuildMergeMaps(
                    worksheet,
                    firstRow,
                    firstCol,
                    lastRow,
                    lastCol
                );

                var contents = new Dictionary<string, string>();
                foreach (var cell in usedRange.CellsUsed())
                {
                    if (cell.DataType == XLDataType.Blank)
                        continue;

                    string? value = cell.Value.ToString();
                    if (string.IsNullOrEmpty(value))
                        continue;

                    // Match against the plain value, not the Markdown-wrapped text.
                    if (!needles.Any(n => value.Contains(n, StringComparison.OrdinalIgnoreCase)))
                        continue;

                    string cellRef = cell.Address.ToString()!;
                    if (mergedMemberCells.Contains(cellRef) && !mergedRangeByAnchor.ContainsKey(cellRef))
                        continue;

                    string key = mergedRangeByAnchor.TryGetValue(cellRef, out string? mergeAddr) ? mergeAddr : cellRef;
                    contents[key] = ToMarkdown(cell);
                }

                if (contents.Count == 0)
                    continue;

                var orderedKeys = contents
                    .Keys.OrderBy(k => CellReferenceParser.GetRowIndex(AnchorAddress(k)))
                    .ThenBy(k =>
                        CellReferenceParser.ColumnLetterToIndex(CellReferenceParser.GetColumnName(AnchorAddress(k)))
                    )
                    .ToList();

                var selectedRows = SelectRows(orderedKeys.Select(k => (k, contents[k])), truncate: 0);

                results.Add(
                    new MarkdownSheet
                    {
                        SheetName = worksheet.Name,
                        TargetRange = string.Empty,
                        Contents = RenderTable(selectedRows),
                    }
                );
            }

            return results;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Error reading Excel file: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Renders the cells of <paramref name="spreadSheetName"/> within <paramref name="range"/> as a
    /// visual grid Markdown table — the spatial counterpart to <see cref="LoadRange"/>'s two-column
    /// Address/Contents table. The first header column is empty; the remaining headers are the column
    /// letters (A, B, C…) of the emitted columns, and each body row is prefixed with its spreadsheet
    /// row number. Each grid cell holds the cell value rendered as Markdown using the same per-cell
    /// rendering as <see cref="LoadRange"/> (bold → "**text**", italic → "*text*", strikethrough →
    /// "~~text~~"; literal "|" escaped as "\|"; newlines collapsed to spaces). Interior empty cells are
    /// kept as blank grid cells so the layout is preserved; a merged region's value appears only in its
    /// top-left anchor cell, with the covered cells left blank. The extent is the requested range
    /// intersected with the used range, with trailing fully-empty rows and columns dropped. When
    /// <paramref name="truncate"/> is positive, trailing rows are dropped so the rendered table does not
    /// exceed that many characters. <see cref="MarkdownSheet.TargetRange"/> reports the tight bounding
    /// box actually emitted.
    /// </summary>
    public MarkdownSheet LoadRangeInMarkdownTable(
        string spreadSheetPath,
        string spreadSheetName,
        string range,
        int truncate = 0
    )
    {
        try
        {
            using var sanitized = WorkbookReader.SanitizePhoneticRuns(spreadSheetPath);
            using var workbook = new XLWorkbook(sanitized);
            var worksheet = workbook.Worksheet(spreadSheetName);

            var bounds = WorkbookReader.ComputeEffectiveRange(worksheet, range);
            if (bounds == null)
                return new MarkdownSheet
                {
                    SheetName = spreadSheetName,
                    TargetRange = range,
                    Contents = string.Empty,
                };

            var (firstRow, firstCol, lastRow, lastCol) = bounds.Value;

            // Map each merged region's covered cells so only the anchor carries the value; the rest
            // render as blank grid cells.
            var (mergedRangeByAnchor, mergedMemberCells) = BuildMergeMaps(
                worksheet,
                firstRow,
                firstCol,
                lastRow,
                lastCol
            );

            // Render every cell in the box to its Markdown value (empty string when blank or merge-
            // covered), indexed by absolute (row, col) so we can later trim and lay out the grid.
            var values = new Dictionary<(int Row, int Col), string>();
            for (int r = firstRow; r <= lastRow; r++)
            for (int c = firstCol; c <= lastCol; c++)
            {
                var cell = worksheet.Cell(r, c);
                string cellRef = cell.Address.ToString()!;

                // A merge-covered cell that is not the anchor stays blank.
                if (mergedMemberCells.Contains(cellRef) && !mergedRangeByAnchor.ContainsKey(cellRef))
                {
                    values[(r, c)] = string.Empty;
                    continue;
                }

                if (cell.DataType == XLDataType.Blank)
                {
                    values[(r, c)] = string.Empty;
                    continue;
                }

                string raw = cell.Value.ToString() ?? string.Empty;
                values[(r, c)] = raw.Length == 0 ? string.Empty : ToMarkdown(cell);
            }

            // Drop trailing fully-empty rows and columns so the grid hugs its content. Interior empty
            // rows/columns are preserved to keep the layout intact.
            int tightLastRow = firstRow - 1;
            for (int r = firstRow; r <= lastRow; r++)
                if (Enumerable.Range(firstCol, lastCol - firstCol + 1).Any(c => values[(r, c)].Length > 0))
                    tightLastRow = r;

            int tightLastCol = firstCol - 1;
            for (int c = firstCol; c <= lastCol; c++)
                if (Enumerable.Range(firstRow, lastRow - firstRow + 1).Any(r => values[(r, c)].Length > 0))
                    tightLastCol = c;

            if (tightLastRow < firstRow || tightLastCol < firstCol)
                return new MarkdownSheet
                {
                    SheetName = spreadSheetName,
                    TargetRange = range,
                    Contents = string.Empty,
                };

            var selected = SelectGridRows(values, firstRow, firstCol, tightLastRow, tightLastCol, truncate);

            return new MarkdownSheet
            {
                SheetName = spreadSheetName,
                TargetRange =
                    selected.LastRow < firstRow
                        ? range
                        : $"{CellReferenceParser.IndexToColumnLetter(firstCol)}{firstRow}:"
                            + $"{CellReferenceParser.IndexToColumnLetter(tightLastCol)}{selected.LastRow}",
                Contents = selected.Table,
            };
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Error reading Excel file: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Builds, for every merged region intersecting the box (firstRow..lastRow, firstCol..lastCol),
    /// a lookup from anchor (top-left) address → full range address and a set of every covered cell
    /// address (so non-anchor members can be skipped). Shared by <see cref="LoadRange"/> and
    /// <see cref="LoadRangeInMarkdownTable"/>.
    /// </summary>
    private static (Dictionary<string, string> ByAnchor, HashSet<string> MemberCells) BuildMergeMaps(
        IXLWorksheet worksheet,
        int firstRow,
        int firstCol,
        int lastRow,
        int lastCol
    )
    {
        var mergedRangeByAnchor = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var mergedMemberCells = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var mergedRange in worksheet.MergedRanges)
        {
            var mr = mergedRange.RangeAddress;
            int mrFirstRow = mr.FirstAddress.RowNumber;
            int mrFirstCol = mr.FirstAddress.ColumnNumber;
            int mrLastRow = mr.LastAddress.RowNumber;
            int mrLastCol = mr.LastAddress.ColumnNumber;

            bool overlaps =
                mrFirstRow <= lastRow && mrLastRow >= firstRow && mrFirstCol <= lastCol && mrLastCol >= firstCol;
            if (!overlaps)
                continue;

            mergedRangeByAnchor[mr.FirstAddress.ToStringRelative(false)] = mr.ToStringRelative(false);
            for (int r = mrFirstRow; r <= mrLastRow; r++)
            for (int c = mrFirstCol; c <= mrLastCol; c++)
                mergedMemberCells.Add(worksheet.Cell(r, c).Address.ToString()!);
        }
        return (mergedRangeByAnchor, mergedMemberCells);
    }

    /// <summary>
    /// Renders the grid (firstRow..tightLastRow, firstCol..tightLastCol) as a visual Markdown table and
    /// returns it together with the last row actually emitted. The header row is "| | A | B | …" (an
    /// empty first column followed by the column letters) and each body row is "| {rowNumber} | … |".
    /// When <paramref name="truncate"/> is positive, trailing body rows are dropped so the rendered
    /// length does not exceed that many characters (the header and separator are always counted).
    /// <see cref="GridResult.LastRow"/> is one less than firstRow when no body row fits.
    /// </summary>
    private static GridResult SelectGridRows(
        Dictionary<(int Row, int Col), string> values,
        int firstRow,
        int firstCol,
        int tightLastRow,
        int tightLastCol,
        int truncate
    )
    {
        var headerCells = new StringBuilder("|");
        var separator = new StringBuilder("|");
        // Leading empty corner cell, then one column per spreadsheet column.
        headerCells.Append(" |");
        separator.Append(" --- |");
        for (int c = firstCol; c <= tightLastCol; c++)
        {
            headerCells.Append($" {CellReferenceParser.IndexToColumnLetter(c)} |");
            separator.Append(" --- |");
        }

        var sb = new StringBuilder();
        sb.Append(headerCells);
        sb.Append('\n').Append(separator);
        int length = sb.Length;
        int lastEmitted = firstRow - 1;

        for (int r = firstRow; r <= tightLastRow; r++)
        {
            var rowBuilder = new StringBuilder($"\n| {r} |");
            for (int c = firstCol; c <= tightLastCol; c++)
                rowBuilder.Append($" {EscapeTableCell(values[(r, c)])} |");

            if (truncate > 0 && length + rowBuilder.Length > truncate)
                break;

            length += rowBuilder.Length;
            sb.Append(rowBuilder);
            lastEmitted = r;
        }

        return new GridResult(sb.ToString(), lastEmitted);
    }

    /// <summary>The rendered grid table and the last spreadsheet row that survived truncation.</summary>
    private readonly record struct GridResult(string Table, int LastRow);

    /// <summary>
    /// Returns the anchor (top-left) address of a Contents key. A merged-range key such as
    /// "AD629:AL638" yields its first address ("AD629"); a single-cell key is returned unchanged.
    /// </summary>
    private static string AnchorAddress(string key)
    {
        int colon = key.IndexOf(':');
        return colon < 0 ? key : key[..colon];
    }

    // The header and separator rows of the two-column Address/Contents Markdown table.
    private const string TableHeader = "| Address | Contents |\n| --- | --- |";

    /// <summary>
    /// Selects the ordered (address, markdown) rows that fit within a truncation budget. When
    /// <paramref name="truncate"/> is positive, rows are taken in order and selection stops before the
    /// rendered table length would exceed that many characters (dropping the last rows in spreadsheet
    /// order); the header is always accounted for. When <paramref name="truncate"/> is not positive all
    /// rows are returned. This is the single source of truth for which rows end up in the table, so the
    /// reported <see cref="MarkdownSheet.TargetRange"/> can be derived from the same set.
    /// </summary>
    private static List<(string Address, string Markdown)> SelectRows(
        IEnumerable<(string Address, string Markdown)> rows,
        int truncate
    )
    {
        var selected = new List<(string, string)>();
        int length = TableHeader.Length;
        foreach (var (address, markdown) in rows)
        {
            string row = $"\n| {address} | {EscapeTableCell(markdown)} |";
            if (truncate > 0 && length + row.Length > truncate)
                break;
            length += row.Length;
            selected.Add((address, markdown));
        }
        return selected;
    }

    /// <summary>
    /// Renders pre-selected ordered (address, markdown) pairs as a two-column Markdown table. A literal
    /// "|" in a cell value is escaped as "\|" so it does not break the table. When there are no rows,
    /// only the header and separator are returned.
    /// </summary>
    private static string RenderTable(IEnumerable<(string Address, string Markdown)> rows)
    {
        var sb = new StringBuilder(TableHeader);
        foreach (var (address, markdown) in rows)
            sb.Append($"\n| {address} | {EscapeTableCell(markdown)} |");
        return sb.ToString();
    }

    /// <summary>
    /// Computes the tight bounding box (e.g. "A1:B3") that encloses every address in
    /// <paramref name="rows"/>, so the reported range reflects exactly the cells present in the table.
    /// Each address is a single cell ("B3") or a merged range ("AD629:AL638"); both endpoints of a
    /// range are considered. Returns <paramref name="fallback"/> when there are no rows.
    /// </summary>
    private static string BoundingBox(IReadOnlyList<(string Address, string Markdown)> rows, string fallback)
    {
        if (rows.Count == 0)
            return fallback;

        int minRow = int.MaxValue,
            minCol = int.MaxValue,
            maxRow = int.MinValue,
            maxCol = int.MinValue;

        foreach (var (address, _) in rows)
        foreach (string endpoint in address.Split(':'))
        {
            int row = (int)CellReferenceParser.GetRowIndex(endpoint);
            int col = CellReferenceParser.ColumnLetterToIndex(CellReferenceParser.GetColumnName(endpoint));
            minRow = Math.Min(minRow, row);
            maxRow = Math.Max(maxRow, row);
            minCol = Math.Min(minCol, col);
            maxCol = Math.Max(maxCol, col);
        }

        return $"{CellReferenceParser.IndexToColumnLetter(minCol)}{minRow}:"
            + $"{CellReferenceParser.IndexToColumnLetter(maxCol)}{maxRow}";
    }

    /// <summary>
    /// Escapes a value for use inside a Markdown table cell: literal "|" becomes "\|" and newlines
    /// become spaces so each entry stays on a single table row. Inverse of <see cref="UnescapeTableCell"/>.
    /// </summary>
    private static string EscapeTableCell(string value) =>
        value.Replace("\r\n", " ").Replace('\n', ' ').Replace('\r', ' ').Replace("|", "\\|");

    /// <summary>
    /// Reverses <see cref="EscapeTableCell"/>: "\|" becomes a literal "|".
    /// </summary>
    private static string UnescapeTableCell(string value) => value.Replace("\\|", "|");

    /// <summary>
    /// Parses a two-column Address/Contents Markdown table back into ordered (address, markdown)
    /// pairs — the inverse of <see cref="RenderTable"/>. The header row, the "| --- | --- |" separator
    /// and blank lines are skipped; each remaining row is split on its unescaped "|" delimiters and its
    /// cells are trimmed and unescaped. Rows that do not have exactly two columns, or whose address is
    /// empty, are ignored.
    /// </summary>
    private static List<(string Address, string Markdown)> ParseTable(string table)
    {
        var rows = new List<(string, string)>();
        foreach (var rawLine in table.Split('\n'))
        {
            string line = rawLine.Trim();
            if (line.Length == 0)
                continue;

            var cells = SplitTableRow(line);
            if (cells.Count != 2)
                continue;

            string address = cells[0].Trim();
            // Skip the header row and the "--- | ---" separator row.
            if (
                address.Length == 0
                || address.Equals("Address", StringComparison.OrdinalIgnoreCase)
                || address.Trim('-', ' ').Length == 0
            )
                continue;

            rows.Add((address, UnescapeTableCell(cells[1].Trim())));
        }
        return rows;
    }

    /// <summary>
    /// Splits a single Markdown table row into its cell values, honoring "\|" as an escaped literal
    /// pipe (not a delimiter) and dropping the leading/trailing empty cells produced by the outer
    /// "|" borders. Escapes are left intact for the caller to unescape.
    /// </summary>
    private static List<string> SplitTableRow(string line)
    {
        var cells = new List<string>();
        var buffer = new StringBuilder();
        for (int i = 0; i < line.Length; i++)
        {
            char ch = line[i];
            if (ch == '\\' && i + 1 < line.Length && line[i + 1] == '|')
            {
                buffer.Append("\\|");
                i++;
            }
            else if (ch == '|')
            {
                cells.Add(buffer.ToString());
                buffer.Clear();
            }
            else
            {
                buffer.Append(ch);
            }
        }
        cells.Add(buffer.ToString());

        // Drop the empty leading/trailing cells created by the outer "|" borders.
        if (cells.Count > 0 && cells[0].Trim().Length == 0)
            cells.RemoveAt(0);
        if (cells.Count > 0 && cells[^1].Trim().Length == 0)
            cells.RemoveAt(cells.Count - 1);
        return cells;
    }

    /// <summary>
    /// Renders a cell's text as Markdown. When the cell carries rich text, each run is wrapped
    /// individually so partial formatting is preserved (e.g. "sample **string**"). Otherwise the
    /// whole-cell font style is applied to the entire value.
    /// </summary>
    private static string ToMarkdown(IXLCell cell)
    {
        if (cell.HasRichText)
        {
            var sb = new StringBuilder();
            foreach (var run in cell.GetRichText())
                sb.Append(WrapMarkdown(run.Text, run.Bold, run.Italic, run.Strikethrough));
            return sb.ToString();
        }

        string value = cell.Value.ToString() ?? string.Empty;
        var font = cell.Style.Font;
        return WrapMarkdown(value, font.Bold, font.Italic, font.Strikethrough);
    }

    /// <summary>
    /// Wraps <paramref name="text"/> with Markdown markers for the given styles. Strikethrough is
    /// applied outermost, then bold, then italic, so combined styles nest predictably
    /// (e.g. bold+italic → "***text***"). Whitespace-only text is returned unwrapped.
    /// </summary>
    private static string WrapMarkdown(string text, bool bold, bool italic, bool strikethrough)
    {
        if (string.IsNullOrWhiteSpace(text))
            return text;

        if (italic)
            text = $"*{text}*";
        if (bold)
            text = $"**{text}**";
        if (strikethrough)
            text = $"~~{text}~~";
        return text;
    }

    /// <summary>
    /// Writes a Markdown-formatted <see cref="MarkdownSheet"/> into the worksheet named by
    /// <see cref="MarkdownSheet.SheetName"/> — the inverse of <see cref="LoadRange"/>. The
    /// <see cref="MarkdownSheet.Contents"/> two-column table is parsed row by row; each Address is a
    /// single cell address (e.g. "A1") or a merged range address (e.g. "AD629:AL638"); range keys are
    /// merged and the value written to the anchor (top-left) cell. Inline Markdown is converted to
    /// rich text: "**text**" → bold, "*text*" → italic, "~~text~~" → strikethrough.
    /// </summary>
    public void UpdateRange(string spreadSheetPath, MarkdownSheet sheet)
    {
        if (sheet == null || string.IsNullOrWhiteSpace(sheet.Contents))
            return;

        var rows = ParseTable(sheet.Contents);
        if (rows.Count == 0)
            return;

        try
        {
            using var workbook = new XLWorkbook(spreadSheetPath);
            var worksheet = workbook.Worksheet(sheet.SheetName);

            foreach (var (address, markdown) in rows)
            {
                IXLCell cell;

                // A key containing ':' is a merged range address (e.g. "AD629:AL638"): merge the
                // region and write into its anchor (top-left) cell. Otherwise it is a single cell.
                if (address.Contains(':'))
                {
                    var range = worksheet.Range(address);
                    range.Merge();
                    cell = range.FirstCell();
                }
                else
                {
                    cell = worksheet.Cell(address);
                }

                WriteMarkdown(cell, markdown);
            }

            workbook.Save();
        }
        catch (Exception ex) when (ex is not InvalidOperationException)
        {
            throw new InvalidOperationException($"Error updating Excel file: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Clears <paramref name="cell"/> and writes <paramref name="markdown"/> as rich text, parsing
    /// inline Markdown into runs. This is the inverse of <see cref="WrapMarkdown"/>: "**text**" →
    /// bold, "*text*" → italic, "~~text~~" → strikethrough, with combined markers composing onto a
    /// single run (e.g. "***text***" → bold + italic). Text outside any marker is written plain.
    /// </summary>
    private static void WriteMarkdown(IXLCell cell, string markdown)
    {
        cell.Clear(XLClearOptions.Contents);

        var richText = cell.GetRichText();
        richText.ClearText();

        foreach (var (text, bold, italic, strikethrough) in ParseMarkdown(markdown))
        {
            var run = richText.AddText(text);
            if (bold)
                run.SetBold(true);
            if (italic)
                run.SetItalic(true);
            if (strikethrough)
                run.SetStrikethrough(true);
        }
    }

    /// <summary>
    /// Splits Markdown text into runs, each tagged with the bold/italic/strikethrough styles active
    /// over it. Recognized markers are "~~" (strikethrough), "**" (bold) and "*" (italic); they may
    /// nest in any order. Unmatched markers are treated as literal text.
    /// </summary>
    private static List<(string Text, bool Bold, bool Italic, bool Strikethrough)> ParseMarkdown(string markdown)
    {
        var runs = new List<(string, bool, bool, bool)>();
        var buffer = new StringBuilder();
        bool bold = false,
            italic = false,
            strikethrough = false;
        int i = 0;

        void Flush()
        {
            if (buffer.Length > 0)
            {
                runs.Add((buffer.ToString(), bold, italic, strikethrough));
                buffer.Clear();
            }
        }

        while (i < markdown.Length)
        {
            if (markdown[i] == '~' && i + 1 < markdown.Length && markdown[i + 1] == '~')
            {
                Flush();
                strikethrough = !strikethrough;
                i += 2;
            }
            else if (markdown[i] == '*' && i + 1 < markdown.Length && markdown[i + 1] == '*')
            {
                Flush();
                bold = !bold;
                i += 2;
            }
            else if (markdown[i] == '*')
            {
                Flush();
                italic = !italic;
                i += 1;
            }
            else
            {
                buffer.Append(markdown[i]);
                i += 1;
            }
        }

        Flush();
        return runs;
    }
}
