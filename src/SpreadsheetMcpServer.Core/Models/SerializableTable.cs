namespace SpreadsheetMcpServer.Core.Models;

public class SerializableTable
{
    /// <summary>Internal table name, e.g. "Table1".</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>Cell range reference, e.g. "A1:C5".</summary>
    public string Reference { get; set; } = string.Empty;

    /// <summary>
    /// Built-in Excel table style name, e.g. "TableStyleMedium9".
    /// Null or empty means no style is applied.
    /// </summary>
    public string? StyleName { get; set; }

    public bool ShowFirstColumn { get; set; }
    public bool ShowLastColumn { get; set; }
    public bool ShowRowStripes { get; set; } = true;
    public bool ShowColumnStripes { get; set; }

    /// <summary>Header names for each column, in left-to-right order.</summary>
    public List<string> Columns { get; set; } = [];
}
