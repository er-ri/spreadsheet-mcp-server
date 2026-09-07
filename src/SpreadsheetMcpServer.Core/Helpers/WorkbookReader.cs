using System.IO.Compression;
using System.Xml.Linq;
using ClosedXML.Excel;

namespace SpreadsheetMcpServer.Core.Helpers;

/// <summary>
/// Shared plumbing for the read paths: loading a workbook safely and resolving the portion of a
/// worksheet a request actually covers.
/// </summary>
public static class WorkbookReader
{
    // ClosedXML throws when an xlsx shared-strings part contains phonetic run (<rPh>)
    // elements that are out-of-order or overlapping. Strip all <rPh> nodes before loading.
    public static Stream SanitizePhoneticRuns(string path)
    {
        var ms = new MemoryStream();
        using (var src = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            src.CopyTo(ms);

        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            var sharedStrings = zip.GetEntry("xl/sharedStrings.xml");
            if (sharedStrings != null)
            {
                XDocument doc;
                using (var entryStream = sharedStrings.Open())
                    doc = XDocument.Load(entryStream);

                XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                bool removed = false;
                foreach (var rPh in doc.Descendants(ns + "rPh").ToList())
                {
                    rPh.Remove();
                    removed = true;
                }

                if (removed)
                {
                    sharedStrings.Delete();
                    var newEntry = zip.CreateEntry("xl/sharedStrings.xml");
                    using var writer = new StreamWriter(newEntry.Open());
                    writer.Write(doc.ToString());
                }
            }
        }

        ms.Position = 0;
        return ms;
    }

    /// <summary>
    /// Returns the worksheet named <paramref name="name"/>, throwing a uniform
    /// <see cref="InvalidOperationException"/> when it does not exist. ClosedXML's
    /// <c>Worksheet(string)</c> throws <see cref="ArgumentException"/> instead, which the services'
    /// catch filters deliberately let through — so every lookup goes through here to keep a missing
    /// sheet reported as an expected condition with one message.
    /// </summary>
    public static IXLWorksheet GetWorksheet(IXLWorkbook workbook, string name)
    {
        if (!workbook.TryGetWorksheet(name, out var worksheet))
            throw new InvalidOperationException($"No worksheet named '{name}' was found.");

        return worksheet;
    }

    /// <summary>
    /// Intersects the requested <paramref name="range"/> with the worksheet's used range and returns the
    /// resulting (firstRow, firstCol, lastRow, lastCol) bounds, or <c>null</c> when the sheet has no used
    /// range or the intersection is empty. <paramref name="options"/> selects what counts as "used":
    /// leave it null for ClosedXML's default (cell contents only), or pass
    /// <see cref="XLCellsUsedOptions.AllFormats"/> as well so cells carrying only styling are included.
    /// </summary>
    public static (int FirstRow, int FirstCol, int LastRow, int LastCol)? ComputeEffectiveRange(
        IXLWorksheet worksheet,
        string range,
        XLCellsUsedOptions? options = null
    )
    {
        var requestedRange = worksheet.Range(range);
        var usedRangeAddress = (
            options is null ? worksheet.RangeUsed() : worksheet.RangeUsed(options.Value)
        )?.RangeAddress;
        if (usedRangeAddress == null)
            return null;

        var usedRange = worksheet.Range(usedRangeAddress.ToStringRelative(false));
        int firstRow = Math.Max(requestedRange.FirstRow().RowNumber(), usedRange.FirstRow().RowNumber());
        int firstCol = Math.Max(requestedRange.FirstColumn().ColumnNumber(), usedRange.FirstColumn().ColumnNumber());
        int lastRow = Math.Min(requestedRange.LastRow().RowNumber(), usedRange.LastRow().RowNumber());
        int lastCol = Math.Min(requestedRange.LastColumn().ColumnNumber(), usedRange.LastColumn().ColumnNumber());

        if (firstRow > lastRow || firstCol > lastCol)
            return null;

        return (firstRow, firstCol, lastRow, lastCol);
    }
}
