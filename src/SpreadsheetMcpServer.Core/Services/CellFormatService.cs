using System.Globalization;
using System.Text;
using ClosedXML.Excel;
using SpreadsheetMcpServer.Core.Helpers;
using SpreadsheetMcpServer.Core.Models;

namespace SpreadsheetMcpServer.Core.Services;

/// <summary>
/// Service implementation for reading and applying cell formatting on Excel worksheets.
/// </summary>
public class CellFormatService : ICellFormatService
{
    /// <summary>
    /// Applies formatting to cells in an Excel worksheet based on user specifications.
    /// Supports individual cells or ranges with comprehensive formatting options.
    /// </summary>
    public void ApplyCellFormatting(string spreadSheetPath, List<CellFormatSpec> formats)
    {
        if (formats == null || formats.Count == 0)
            return;

        try
        {
            using var workbook = new XLWorkbook(spreadSheetPath);

            foreach (var format in formats)
            {
                if (string.IsNullOrWhiteSpace(format.Address))
                    continue;

                // Determine if this is a range or single cell
                IXLRange? range = null;
                IXLCell? cell = null;

                if (format.Address.Contains(':'))
                {
                    // It's a range like "A1:B5"
                    range = GetRangeFromAnySheet(workbook, format.Address);
                }
                else
                {
                    // It's a single cell like "A1"
                    cell = GetCellFromAnySheet(workbook, format.Address);
                }

                if (range != null)
                {
                    ApplyFormatToRange(range, format);
                }
                else if (cell != null)
                {
                    ApplyFormatToCell(cell, format);
                }
            }

            workbook.Save();
        }
        catch (Exception ex) when (ex is not InvalidOperationException and not ArgumentException)
        {
            throw new InvalidOperationException($"Error applying cell formatting: {ex.Message}", ex);
        }
    }

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
    public List<CellFormatSpec> ReadCellFormatting(
        string spreadSheetPath,
        string spreadSheetName,
        string range,
        int truncate = 200
    )
    {
        try
        {
            using var sanitized = WorkbookReader.SanitizePhoneticRuns(spreadSheetPath);
            using var workbook = new XLWorkbook(sanitized);
            var worksheet = WorkbookReader.GetWorksheet(workbook, spreadSheetName);

            // Include cells that carry only formatting — a colored banner row or a bordered box has no
            // contents but is exactly what this tool exists to report.
            var bounds = WorkbookReader.ComputeEffectiveRange(
                worksheet,
                range,
                XLCellsUsedOptions.AllContents | XLCellsUsedOptions.AllFormats
            );
            if (bounds == null)
                return [];

            var (firstRow, firstCol, lastRow, lastCol) = bounds.Value;
            var defaults = workbook.Style;

            // Row heights and column widths first: they describe the grid rather than any one cell.
            var results = ReadDimensions(worksheet, firstRow, firstCol, lastRow, lastCol);

            // Build a per-cell spec for every non-default cell, keyed by position, then coalesce
            // identically-styled neighbours into rectangles.
            var specs = new Dictionary<(int Row, int Col), (string Signature, CellFormatSpec Spec)>();
            for (int row = firstRow; row <= lastRow; row++)
            for (int col = firstCol; col <= lastCol; col++)
            {
                var spec = BuildSpec(worksheet.Cell(row, col), defaults);
                if (spec != null)
                    specs[(row, col)] = (Signature(spec), spec);
            }

            results.AddRange(CoalesceIntoRectangles(specs, firstRow, firstCol, lastRow, lastCol));

            if (truncate > 0 && results.Count > truncate)
                results.RemoveRange(truncate, results.Count - truncate);

            return results;
        }
        catch (Exception ex) when (ex is not InvalidOperationException and not ArgumentException)
        {
            throw new InvalidOperationException($"Error reading cell formatting: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Returns dimension-only specs for every row whose height and every column whose width differs from
    /// the worksheet default. Each is addressed to a cell inside the affected row/column, because
    /// <see cref="ApplyFormatToCell"/> applies <see cref="CellFormatSpec.Height"/> to that cell's row and
    /// <see cref="CellFormatSpec.Width"/> to its column — so the entries re-apply as written.
    /// </summary>
    private static List<CellFormatSpec> ReadDimensions(
        IXLWorksheet worksheet,
        int firstRow,
        int firstCol,
        int lastRow,
        int lastCol
    )
    {
        // Keyed by address so a row height and a column width that land on the same cell — which they do
        // whenever the first row and first column of the range are both resized — share one entry.
        var dimensions = new Dictionary<string, CellFormatSpec>();
        string firstColLetter = CellReferenceParser.IndexToColumnLetter(firstCol);

        CellFormatSpec At(string address)
        {
            if (!dimensions.TryGetValue(address, out var spec))
                dimensions[address] = spec = new CellFormatSpec { Address = address };
            return spec;
        }

        for (int row = firstRow; row <= lastRow; row++)
        {
            double height = worksheet.Row(row).Height;
            if (Differs(height, worksheet.RowHeight))
                At($"{firstColLetter}{row}").Height = height;
        }

        for (int col = firstCol; col <= lastCol; col++)
        {
            double width = worksheet.Column(col).Width;
            if (Differs(width, worksheet.ColumnWidth))
                At($"{CellReferenceParser.IndexToColumnLetter(col)}{firstRow}").Width = width;
        }

        return [.. dimensions.Values];
    }

    /// <summary>
    /// Maps a cell's ClosedXML style to a <see cref="CellFormatSpec"/> — the inverse of
    /// <see cref="ApplyFormatToCell"/> — setting only the properties that differ from
    /// <paramref name="defaults"/> (the workbook default style). Returns <c>null</c> when nothing differs,
    /// so entirely unstyled cells cost nothing in the output. <see cref="CellFormatSpec.Address"/> is left
    /// blank; the caller assigns it once neighbouring cells have been coalesced.
    /// </summary>
    private static CellFormatSpec? BuildSpec(IXLCell cell, IXLStyle defaults)
    {
        var style = cell.Style;
        var font = style.Font;
        var fill = style.Fill;
        var alignment = style.Alignment;
        var border = style.Border;
        var spec = new CellFormatSpec();
        bool any = false;

        if (font.Bold != defaults.Font.Bold)
            (spec.Bold, any) = (font.Bold, true);

        if (font.Italic != defaults.Font.Italic)
            (spec.Italic, any) = (font.Italic, true);

        if (font.Strikethrough != defaults.Font.Strikethrough)
            (spec.Strikethrough, any) = (font.Strikethrough, true);

        // CellFormatSpec models underline as a bool, so Double/Accounting variants collapse to true.
        if (font.Underline != defaults.Font.Underline)
            (spec.Underline, any) = (font.Underline != XLFontUnderlineValues.None, true);

        if (!string.Equals(font.FontName, defaults.Font.FontName, StringComparison.Ordinal))
            (spec.FontName, any) = (font.FontName, true);

        if (Differs(font.FontSize, defaults.Font.FontSize))
            (spec.FontSize, any) = (font.FontSize, true);

        string? fontColor = ToHex(font.FontColor);
        if (fontColor != null && fontColor != ToHex(defaults.Font.FontColor))
            (spec.FontColor, any) = (fontColor, true);

        // Only a solid fill maps back onto BackgroundColor; other pattern types have no representation.
        if (fill.PatternType == XLFillPatternValues.Solid && ToHex(fill.BackgroundColor) is { } background)
            (spec.BackgroundColor, any) = (background, true);

        if (alignment.Horizontal != defaults.Alignment.Horizontal)
            (spec.HorizontalAlignment, any) = (alignment.Horizontal.ToString(), true);

        if (alignment.Vertical != defaults.Alignment.Vertical)
            (spec.VerticalAlignment, any) = (alignment.Vertical.ToString(), true);

        if (alignment.WrapText != defaults.Alignment.WrapText)
            (spec.WrapText, any) = (alignment.WrapText, true);

        if (ReadBorder(border.TopBorder, border.TopBorderColor) is { } top)
            (spec.TopBorder, any) = (top, true);

        if (ReadBorder(border.BottomBorder, border.BottomBorderColor) is { } bottom)
            (spec.BottomBorder, any) = (bottom, true);

        if (ReadBorder(border.LeftBorder, border.LeftBorderColor) is { } left)
            (spec.LeftBorder, any) = (left, true);

        if (ReadBorder(border.RightBorder, border.RightBorderColor) is { } right)
            (spec.RightBorder, any) = (right, true);

        // A format stored as a built-in id rather than a format string leaves Format empty; there is no
        // string to hand back, so it is skipped.
        string numberFormat = style.NumberFormat.Format;
        if (
            !string.IsNullOrEmpty(numberFormat)
            && !string.Equals(numberFormat, defaults.NumberFormat.Format, StringComparison.Ordinal)
        )
            (spec.NumberFormat, any) = (numberFormat, true);

        return any ? spec : null;
    }

    /// <summary>
    /// Returns a <see cref="BorderSpec"/> for an edge that has a border, or <c>null</c> when it has none.
    /// </summary>
    private static BorderSpec? ReadBorder(XLBorderStyleValues style, XLColor color) =>
        style == XLBorderStyleValues.None ? null : new BorderSpec { Style = style.ToString(), Color = ToHex(color) };

    /// <summary>
    /// Formats a ClosedXML color as "#RRGGBB", or returns <c>null</c> when it has no literal RGB value.
    /// Theme colors are excluded deliberately: <see cref="XLColor.Color"/> throws for them, and a theme
    /// slot cannot be expressed as a hex string that <see cref="ApplyCellFormatting"/> could apply back.
    /// </summary>
    private static string? ToHex(XLColor color) =>
        color is { HasValue: true, ColorType: XLColorType.Color or XLColorType.Indexed }
            ? $"#{color.Color.R:X2}{color.Color.G:X2}{color.Color.B:X2}"
            : null;

    /// <summary>Compares two dimensions with a tolerance, so float noise is not reported as a change.</summary>
    private static bool Differs(double left, double right) => Math.Abs(left - right) > 0.0001;

    /// <summary>
    /// Builds a canonical string identifying a spec's formatting (everything but its address), so two
    /// cells that look identical compare equal when coalescing. Each value is length-prefixed so no
    /// combination of values can produce a colliding signature.
    /// </summary>
    private static string Signature(CellFormatSpec spec)
    {
        var sb = new StringBuilder();
        void Append(object? value)
        {
            string? text = Convert.ToString(value, CultureInfo.InvariantCulture);
            sb.Append(text?.Length ?? -1).Append(':').Append(text);
        }

        Append(spec.Bold);
        Append(spec.Italic);
        Append(spec.Underline);
        Append(spec.Strikethrough);
        Append(spec.FontColor);
        Append(spec.FontName);
        Append(spec.FontSize);
        Append(spec.BackgroundColor);
        Append(spec.HorizontalAlignment);
        Append(spec.VerticalAlignment);
        Append(spec.WrapText);
        Append(spec.NumberFormat);
        foreach (var edge in new[] { spec.TopBorder, spec.BottomBorder, spec.LeftBorder, spec.RightBorder })
        {
            Append(edge?.Style);
            Append(edge?.Color);
        }

        return sb.ToString();
    }

    /// <summary>
    /// Walks the box in row-major order and greedily grows each unclaimed cell into the largest rectangle
    /// of identically-styled neighbours — right as far as the signature holds, then down as far as every
    /// cell of that full-width strip matches. Each rectangle becomes one entry, addressed as a single cell
    /// ("A1") when it is 1x1 and as a range ("A1:D1", "A2:A10") otherwise.
    /// </summary>
    private static List<CellFormatSpec> CoalesceIntoRectangles(
        Dictionary<(int Row, int Col), (string Signature, CellFormatSpec Spec)> specs,
        int firstRow,
        int firstCol,
        int lastRow,
        int lastCol
    )
    {
        var results = new List<CellFormatSpec>();
        var claimed = new HashSet<(int Row, int Col)>();

        for (int row = firstRow; row <= lastRow; row++)
        for (int col = firstCol; col <= lastCol; col++)
        {
            if (claimed.Contains((row, col)) || !specs.TryGetValue((row, col), out var origin))
                continue;

            bool Matches(int r, int c) =>
                !claimed.Contains((r, c))
                && specs.TryGetValue((r, c), out var other)
                && other.Signature == origin.Signature;

            int blockLastCol = col;
            while (blockLastCol + 1 <= lastCol && Matches(row, blockLastCol + 1))
                blockLastCol++;

            int blockLastRow = row;
            while (blockLastRow + 1 <= lastRow && RowSpanMatches(blockLastRow + 1, col, blockLastCol, Matches))
                blockLastRow++;

            for (int r = row; r <= blockLastRow; r++)
            for (int c = col; c <= blockLastCol; c++)
                claimed.Add((r, c));

            string topLeft = $"{CellReferenceParser.IndexToColumnLetter(col)}{row}";
            string bottomRight = $"{CellReferenceParser.IndexToColumnLetter(blockLastCol)}{blockLastRow}";
            origin.Spec.Address = topLeft == bottomRight ? topLeft : $"{topLeft}:{bottomRight}";
            results.Add(origin.Spec);
        }

        return results;
    }

    /// <summary>
    /// Returns true when every cell of <paramref name="row"/> from <paramref name="fromCol"/> to
    /// <paramref name="toCol"/> matches the block being grown.
    /// </summary>
    private static bool RowSpanMatches(int row, int fromCol, int toCol, Func<int, int, bool> matches)
    {
        for (int col = fromCol; col <= toCol; col++)
        {
            if (!matches(row, col))
                return false;
        }
        return true;
    }

    /// <summary>
    /// Gets a range from the specified address. Supports both "Sheet!A1:B5" and "A1:B5" formats.
    /// If no sheet is specified, uses the first worksheet.
    /// </summary>
    private static IXLRange? GetRangeFromAnySheet(XLWorkbook workbook, string address)
    {
        try
        {
            // Check if address includes sheet name (e.g., "Sheet1!A1:B5")
            var parts = address.Split('!', 2);

            if (parts.Length == 2)
            {
                var sheetName = parts[0].Trim('\'', '"');
                var rangeAddress = parts[1];
                var worksheet = workbook.Worksheet(sheetName);
                return worksheet.Range(rangeAddress);
            }
            else
            {
                // No sheet specified, use the first worksheet
                var worksheet = workbook.Worksheet(1);
                return worksheet.Range(address);
            }
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Gets a cell from the specified address. Supports both "Sheet!A1" and "A1" formats.
    /// If no sheet is specified, uses the first worksheet.
    /// </summary>
    private static IXLCell? GetCellFromAnySheet(XLWorkbook workbook, string address)
    {
        try
        {
            // Check if address includes sheet name (e.g., "Sheet1!A1")
            var parts = address.Split('!', 2);

            if (parts.Length == 2)
            {
                var sheetName = parts[0].Trim('\'', '"');
                var cellAddress = parts[1];
                var worksheet = workbook.Worksheet(sheetName);
                return worksheet.Cell(cellAddress);
            }
            else
            {
                // No sheet specified, use the first worksheet
                var worksheet = workbook.Worksheet(1);
                return worksheet.Cell(address);
            }
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Applies all specified formatting to a range of cells.
    /// </summary>
    private static void ApplyFormatToRange(IXLRange range, CellFormatSpec format)
    {
        foreach (var cell in range.Cells())
        {
            ApplyFormatToCell(cell, format);
        }

        // Apply row height and column width to the range
        if (format.Height.HasValue)
        {
            var worksheet = range.Worksheet;
            foreach (var row in range.Rows())
            {
                worksheet.Row(row.RowNumber()).Height = format.Height.Value;
            }
        }

        if (format.Width.HasValue)
        {
            var worksheet = range.Worksheet;
            foreach (var column in range.Columns())
            {
                worksheet.Column(column.ColumnNumber()).Width = format.Width.Value;
            }
        }
    }

    /// <summary>
    /// Applies all specified formatting to a single cell.
    /// </summary>
    private static void ApplyFormatToCell(IXLCell cell, CellFormatSpec format)
    {
        var style = cell.Style;

        // Font formatting
        if (format.Bold.HasValue)
            style.Font.Bold = format.Bold.Value;

        if (format.Italic.HasValue)
            style.Font.Italic = format.Italic.Value;

        if (format.Underline.HasValue)
            style.Font.Underline = format.Underline.Value ? XLFontUnderlineValues.Single : XLFontUnderlineValues.None;

        if (format.Strikethrough.HasValue)
            style.Font.Strikethrough = format.Strikethrough.Value;

        if (!string.IsNullOrEmpty(format.FontColor))
            style.Font.FontColor = XLColor.FromHtml(format.FontColor);

        if (!string.IsNullOrEmpty(format.FontName))
            style.Font.FontName = format.FontName;

        if (format.FontSize.HasValue)
            style.Font.FontSize = format.FontSize.Value;

        // Fill/Background
        if (!string.IsNullOrEmpty(format.BackgroundColor))
            style.Fill.BackgroundColor = XLColor.FromHtml(format.BackgroundColor);

        // Alignment
        if (!string.IsNullOrEmpty(format.HorizontalAlignment))
            style.Alignment.Horizontal = ParseHorizontalAlignment(format.HorizontalAlignment);

        if (!string.IsNullOrEmpty(format.VerticalAlignment))
            style.Alignment.Vertical = ParseVerticalAlignment(format.VerticalAlignment);

        if (format.WrapText.HasValue)
            style.Alignment.WrapText = format.WrapText.Value;

        // Borders
        if (format.TopBorder != null)
            ApplyBorder(style.Border.SetTopBorder, style.Border.SetTopBorderColor, format.TopBorder);

        if (format.BottomBorder != null)
            ApplyBorder(style.Border.SetBottomBorder, style.Border.SetBottomBorderColor, format.BottomBorder);

        if (format.LeftBorder != null)
            ApplyBorder(style.Border.SetLeftBorder, style.Border.SetLeftBorderColor, format.LeftBorder);

        if (format.RightBorder != null)
            ApplyBorder(style.Border.SetRightBorder, style.Border.SetRightBorderColor, format.RightBorder);

        // Number format
        if (!string.IsNullOrEmpty(format.NumberFormat))
            cell.Style.NumberFormat.Format = format.NumberFormat;

        // Row height and column width
        if (format.Height.HasValue)
        {
            cell.Worksheet.Row(cell.Address.RowNumber).Height = format.Height.Value;
        }

        if (format.Width.HasValue)
        {
            cell.Worksheet.Column(cell.Address.ColumnNumber).Width = format.Width.Value;
        }
    }

    /// <summary>
    /// Parses horizontal alignment string to ClosedXML enum.
    /// </summary>
    private static XLAlignmentHorizontalValues ParseHorizontalAlignment(string? alignment)
    {
        return alignment?.ToLowerInvariant() switch
        {
            "left" => XLAlignmentHorizontalValues.Left,
            "center" => XLAlignmentHorizontalValues.Center,
            "right" => XLAlignmentHorizontalValues.Right,
            "justify" => XLAlignmentHorizontalValues.Justify,
            "fill" => XLAlignmentHorizontalValues.Fill,
            "centercontinuous" => XLAlignmentHorizontalValues.CenterContinuous,
            "distributed" => XLAlignmentHorizontalValues.Distributed,
            _ => XLAlignmentHorizontalValues.General,
        };
    }

    /// <summary>
    /// Parses vertical alignment string to ClosedXML enum.
    /// </summary>
    private static XLAlignmentVerticalValues ParseVerticalAlignment(string? alignment)
    {
        return alignment?.ToLowerInvariant() switch
        {
            "top" => XLAlignmentVerticalValues.Top,
            "center" or "middle" => XLAlignmentVerticalValues.Center,
            "bottom" => XLAlignmentVerticalValues.Bottom,
            "justify" => XLAlignmentVerticalValues.Justify,
            "distributed" => XLAlignmentVerticalValues.Distributed,
            _ => XLAlignmentVerticalValues.Bottom,
        };
    }

    /// <summary>
    /// Applies border style and color using ClosedXML border setter methods.
    /// </summary>
    private static void ApplyBorder(
        Func<XLBorderStyleValues, IXLStyle> setBorderStyle,
        Func<XLColor, IXLStyle> setBorderColor,
        BorderSpec spec
    )
    {
        if (!string.IsNullOrEmpty(spec.Style))
        {
            setBorderStyle(ParseBorderStyle(spec.Style));
        }

        if (!string.IsNullOrEmpty(spec.Color))
        {
            setBorderColor(XLColor.FromHtml(spec.Color));
        }
    }

    /// <summary>
    /// Parses border style string to ClosedXML enum.
    /// </summary>
    private static XLBorderStyleValues ParseBorderStyle(string? style)
    {
        return style?.ToLowerInvariant() switch
        {
            "thin" => XLBorderStyleValues.Thin,
            "medium" => XLBorderStyleValues.Medium,
            "thick" => XLBorderStyleValues.Thick,
            "dashed" => XLBorderStyleValues.Dashed,
            "dotted" => XLBorderStyleValues.Dotted,
            "double" => XLBorderStyleValues.Double,
            "hair" => XLBorderStyleValues.Hair,
            "mediumdashdot" => XLBorderStyleValues.MediumDashDot,
            "dashdot" => XLBorderStyleValues.DashDot,
            "dashdotdot" => XLBorderStyleValues.DashDotDot,
            "slantdashdot" => XLBorderStyleValues.SlantDashDot,
            // Any remaining XLBorderStyleValues name (MediumDashDotDot, MediumDashed, None) is accepted
            // verbatim, so a style read by ReadCellFormatting survives being applied back.
            _ => Enum.TryParse<XLBorderStyleValues>(style, ignoreCase: true, out var parsed)
                ? parsed
                : XLBorderStyleValues.None,
        };
    }
}
