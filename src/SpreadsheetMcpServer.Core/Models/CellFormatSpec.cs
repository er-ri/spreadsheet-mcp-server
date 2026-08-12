namespace SpreadsheetMcpServer.Core.Models;

/// <summary>
/// Specifies formatting options for a cell or range.
/// </summary>
public class CellFormatSpec
{
    /// <summary>
    /// Cell address or range (e.g., "A1" or "A1:B5").
    /// </summary>
    public string Address { get; set; } = string.Empty;

    // Font formatting
    public bool? Bold { get; set; }
    public bool? Italic { get; set; }
    public bool? Underline { get; set; }
    public bool? Strikethrough { get; set; }
    public string? FontColor { get; set; } // Hex color (e.g., "#FF0000")
    public string? FontName { get; set; }
    public double? FontSize { get; set; }

    // Fill/Background
    public string? BackgroundColor { get; set; } // Hex color (e.g., "#FFFF00")

    // Alignment
    public string? HorizontalAlignment { get; set; } // "Left", "Center", "Right", "Justify", "Fill"
    public string? VerticalAlignment { get; set; } // "Top", "Center", "Bottom"
    public bool? WrapText { get; set; }

    // Borders
    public BorderSpec? TopBorder { get; set; }
    public BorderSpec? BottomBorder { get; set; }
    public BorderSpec? LeftBorder { get; set; }
    public BorderSpec? RightBorder { get; set; }

    // Number format
    public string? NumberFormat { get; set; } // e.g., "0.00", "yyyy-mm-dd", "#,##0"

    // Row/Column dimensions (in points/pixels)
    public double? Height { get; set; } // Row height
    public double? Width { get; set; } // Column width
}

/// <summary>
/// Specifies border style and color.
/// </summary>
public class BorderSpec
{
    public string? Style { get; set; } // "Thin", "Medium", "Thick", "Dashed", "Dotted"
    public string? Color { get; set; } // Hex color
}
