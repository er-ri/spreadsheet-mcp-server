namespace SpreadsheetMcpServer.Core.Models;

public class SerializableCell
{
    /// <summary>Sheet-qualified cell address, e.g. "Sheet1!A1".</summary>
    public string Address { get; set; } = string.Empty;

    /// <summary>
    /// Cell data type: "SharedString", "Number", "Boolean", "Error",
    /// "String" (formula result), "InlineString", or "Date".
    /// Null means a plain numeric cell.
    /// </summary>
    public string? DataType { get; set; }

    /// <summary>Raw value as stored in the XML (shared string index, number string, etc.).</summary>
    public string? RawValue { get; set; }

    /// <summary>
    /// Resolved value: shared string indices are expanded to their actual text;
    /// all other types are identical to RawValue.
    /// </summary>
    public string? Value { get; set; }

    /// <summary>Formula text (e.g. "SUM(A1:A10)"), or null if the cell has no formula.</summary>
    public string? Formula { get; set; }
}
