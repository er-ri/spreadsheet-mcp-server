using System.Globalization;

namespace SpreadsheetMcpServer.Core.Helpers;

/// <summary>
/// Provides utilities for parsing and manipulating Excel cell references and ranges.
/// </summary>
public static class CellReferenceParser
{
    /// <summary>Returns the column letters portion of a cell reference (e.g. "AB" from "AB12").</summary>
    public static string GetColumnName(string cellReference)
    {
        int i = 0;
        while (i < cellReference.Length && char.IsLetter(cellReference[i]))
            i++;
        return cellReference[..i];
    }

    /// <summary>Returns the row number portion of a cell reference (e.g. 12 from "AB12").</summary>
    public static int GetRowIndex(string cellReference)
    {
        int i = 0;
        while (i < cellReference.Length && char.IsLetter(cellReference[i]))
            i++;
        return int.Parse(cellReference[i..], CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Parses a range like "A1:C5" into (startCol, startRow, endCol, endRow).
    /// </summary>
    public static (string startCol, int startRow, string endCol, int endRow) ParseRange(string reference)
    {
        string[] parts = reference.Split(':');
        if (parts.Length != 2)
            throw new InvalidOperationException($"Invalid range reference: '{reference}'.");

        string startCol = GetColumnName(parts[0]);
        int startRow = GetRowIndex(parts[0]);
        string endCol = GetColumnName(parts[1]);
        int endRow = GetRowIndex(parts[1]);
        return (startCol, startRow, endCol, endRow);
    }

    /// <summary>Converts a column letter (e.g. "A") to a 1-based column index.</summary>
    public static int ColumnLetterToIndex(string col)
    {
        int result = 0;
        foreach (char c in col.ToUpperInvariant())
            result = result * 26 + (c - 'A' + 1);
        return result;
    }

    /// <summary>Converts a 1-based column index to a column letter (e.g. 1 → "A").</summary>
    public static string IndexToColumnLetter(int index)
    {
        string result = string.Empty;
        while (index > 0)
        {
            int rem = (index - 1) % 26;
            result = (char)('A' + rem) + result;
            index = (index - 1) / 26;
        }
        return result;
    }

    /// <summary>Returns all column letters between <paramref name="start"/> and <paramref name="end"/> inclusive.</summary>
    public static List<string> ExpandColumns(string start, string end)
    {
        int s = ColumnLetterToIndex(start);
        int e = ColumnLetterToIndex(end);
        return Enumerable.Range(s, e - s + 1).Select(IndexToColumnLetter).ToList();
    }
}
