using ClosedXML.Excel;
using SpreadsheetMcpServer.Core.Models;

namespace SpreadsheetMcpServer.Core.Services;

/// <summary>
/// Service implementation for applying cell formatting to Excel worksheets.
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
        catch (Exception ex) when (ex is not InvalidOperationException)
        {
            throw new InvalidOperationException($"Error applying cell formatting: {ex.Message}", ex);
        }
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
            _ => XLBorderStyleValues.None,
        };
    }
}
