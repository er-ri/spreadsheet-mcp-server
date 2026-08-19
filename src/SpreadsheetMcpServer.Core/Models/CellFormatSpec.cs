using System.Text.Json.Serialization;

namespace SpreadsheetMcpServer.Core.Models;

/// <summary>
/// Specifies formatting options for a cell or range. Unset (null) properties are left untouched when
/// applied and are omitted from the JSON when serialized, so a spec read back from a worksheet only
/// carries the properties that actually differ from the workbook default.
/// </summary>
public class CellFormatSpec
{
    /// <summary>
    /// Cell address or range (e.g., "A1" or "A1:B5").
    /// </summary>
    public string Address { get; set; } = string.Empty;

    // Font formatting
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? Bold { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? Italic { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? Underline { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? Strikethrough { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FontColor { get; set; } // Hex color (e.g., "#FF0000")

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FontName { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public double? FontSize { get; set; }

    // Fill/Background
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? BackgroundColor { get; set; } // Hex color (e.g., "#FFFF00")

    // Alignment
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? HorizontalAlignment { get; set; } // "Left", "Center", "Right", "Justify", "Fill"

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? VerticalAlignment { get; set; } // "Top", "Center", "Bottom"

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? WrapText { get; set; }

    // Borders
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public BorderSpec? TopBorder { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public BorderSpec? BottomBorder { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public BorderSpec? LeftBorder { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public BorderSpec? RightBorder { get; set; }

    // Number format
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? NumberFormat { get; set; } // e.g., "0.00", "yyyy-mm-dd", "#,##0"

    // Row/Column dimensions (in points/pixels)
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public double? Height { get; set; } // Row height

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public double? Width { get; set; } // Column width
}

/// <summary>
/// Specifies border style and color.
/// </summary>
public class BorderSpec
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Style { get; set; } // "Thin", "Medium", "Thick", "Dashed", "Dotted"

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Color { get; set; } // Hex color
}
