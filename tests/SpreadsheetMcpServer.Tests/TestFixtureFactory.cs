using System.IO.Compression;
using System.Xml.Linq;
using ClosedXML.Excel;

namespace SpreadsheetMcpServer.Tests;

/// <summary>
/// Creates temporary .xlsx files for use in tests.
/// The caller is responsible for deleting the returned file path.
/// </summary>
internal static class TestFixtureFactory
{
    /// <summary>
    /// Creates a temporary workbook with one sheet containing two text cells
    /// ("Hello" in A1, "World" in B1) and one numeric cell (42 in C1).
    /// </summary>
    public static string CreateSimpleWorkbook(string sheetName = "Sheet1")
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_{Guid.NewGuid():N}.xlsx");

        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add(sheetName);
        worksheet.Cell("A1").Value = "Hello";
        worksheet.Cell("B1").Value = "World";
        worksheet.Cell("C1").Value = 42;
        workbook.SaveAs(tempFile);

        return tempFile;
    }

    /// <summary>
    /// Creates a workbook whose sharedStrings.xml contains out-of-order &lt;rPh&gt; elements,
    /// which triggers the ClosedXML "Phonetic runs must be in ascending order" bug.
    /// The sheet has one cell (A1) with the value "テスト" (katakana for "test").
    /// </summary>
    /// <summary>
    /// Creates a workbook with one sheet containing a formatted cell:
    /// A1 = "Styled" — bold, italic, strikethrough, single-underline, superscript, red font, yellow fill, center-aligned, wrap-text.
    /// </summary>
    public static string CreateFormattedWorkbook(string sheetName = "Sheet1")
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_fmt_{Guid.NewGuid():N}.xlsx");

        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add(sheetName);
        var cell = worksheet.Cell("A1");
        cell.Value = "Styled";
        cell.Style.Font.Bold = true;
        cell.Style.Font.Italic = true;
        cell.Style.Font.Strikethrough = true;
        cell.Style.Font.Underline = XLFontUnderlineValues.Single;
        cell.Style.Font.VerticalAlignment = XLFontVerticalTextAlignmentValues.Superscript;
        cell.Style.Font.FontColor = XLColor.Red;
        cell.Style.Fill.BackgroundColor = XLColor.Yellow;
        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        cell.Style.Alignment.WrapText = true;
        workbook.SaveAs(tempFile);

        return tempFile;
    }

    /// <summary>
    /// Creates a workbook with one merged region: A1:C2 (value "Merged"), plus D1 = "Solo".
    /// </summary>
    public static string CreateWorkbookWithMergedCells(string sheetName = "Sheet1")
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_merged_{Guid.NewGuid():N}.xlsx");

        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add(sheetName);
        worksheet.Cell("A1").Value = "Merged";
        worksheet.Range("A1:C2").Merge();
        worksheet.Cell("D1").Value = "Solo";
        workbook.SaveAs(tempFile);

        return tempFile;
    }

    /// <summary>
    /// Creates a workbook exercising the Markdown LoadRange output:
    /// A1 = plain "Plain"; A2 = rich text "sample " + bold "string" (partial run);
    /// A3 = whole-cell italic "ItalicAll"; A4 = whole-cell strikethrough "StruckAll";
    /// A5 = empty string (should be excluded); B1:C1 merged with value "MergedVal".
    /// </summary>
    public static string CreateWorkbookForMarkdown(string sheetName = "Sheet1")
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_md_{Guid.NewGuid():N}.xlsx");

        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add(sheetName);

        ws.Cell("A1").Value = "Plain";

        // Partial rich text: only the second run is bold.
        var a2 = ws.Cell("A2");
        a2.GetRichText().AddText("sample ");
        a2.GetRichText().AddText("string").SetBold(true);

        var a3 = ws.Cell("A3");
        a3.Value = "ItalicAll";
        a3.Style.Font.Italic = true;

        var a4 = ws.Cell("A4");
        a4.Value = "StruckAll";
        a4.Style.Font.Strikethrough = true;

        ws.Cell("A5").Value = ""; // empty — excluded

        ws.Cell("B1").Value = "MergedVal";
        ws.Range("B1:C1").Merge();

        workbook.SaveAs(tempFile);

        return tempFile;
    }

    /// <summary>
    /// Creates a workbook with two sheets for exercising FindStringsInSheets:
    /// "Sheet1" — A1 = "Hello", B1 = "World", C1 = "apple";
    /// "Sheet2" — A1 = "Goodbye", B1 = "Hello again", C1 = 42 (numeric).
    /// "Hello" appears on both sheets (A1 on Sheet1, B1 as substring on Sheet2).
    /// </summary>
    public static string CreateMultiSheetWorkbook()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_multi_{Guid.NewGuid():N}.xlsx");

        using var workbook = new XLWorkbook();

        var sheet1 = workbook.Worksheets.Add("Sheet1");
        sheet1.Cell("A1").Value = "Hello";
        sheet1.Cell("B1").Value = "World";
        sheet1.Cell("C1").Value = "apple";

        var sheet2 = workbook.Worksheets.Add("Sheet2");
        sheet2.Cell("A1").Value = "Goodbye";
        sheet2.Cell("B1").Value = "Hello again";
        sheet2.Cell("C1").Value = 42;

        workbook.SaveAs(tempFile);

        return tempFile;
    }

    /// <summary>
    /// A minimal valid 1x1 PNG, used as picture content in fixtures.
    /// </summary>
    public static readonly byte[] OnePixelPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M8AAAMCAQAU8m8AAAAASUVORK5CYII="
    );

    /// <summary>
    /// Creates a workbook with one picture anchored to cell B2 plus a plain cell (A1 = "Hello").
    /// </summary>
    public static string CreateWorkbookWithPicture(string sheetName = "Sheet1", string anchor = "B2")
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_pic_{Guid.NewGuid():N}.xlsx");

        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add(sheetName);
        worksheet.Cell("A1").Value = "Hello";

        using (var stream = new MemoryStream(OnePixelPng))
            worksheet.AddPicture(stream, "Pic1").MoveTo(worksheet.Cell(anchor));

        workbook.SaveAs(tempFile);

        return tempFile;
    }

    /// <summary>
    /// Writes <see cref="OnePixelPng"/> to a temporary .png file and returns its path,
    /// for use as the source image in PastePictures tests.
    /// The caller is responsible for deleting the returned file path.
    /// </summary>
    public static string CreateTempPngFile()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_src_{Guid.NewGuid():N}.png");
        File.WriteAllBytes(tempFile, OnePixelPng);
        return tempFile;
    }

    public static string CreateWorkbookWithOutOfOrderPhoneticRuns(string sheetName = "Sheet1")
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_phonetic_{Guid.NewGuid():N}.xlsx");

        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add(sheetName);
            worksheet.Cell("A1").Value = "テスト";
            workbook.SaveAs(tempFile);
        }

        // Reopen the ZIP and inject out-of-order <rPh> elements into sharedStrings.xml.
        // sb/eb are 0-based character positions; placing sb=2 before sb=0 is intentionally wrong.
        using (
            var zip = new ZipArchive(
                new FileStream(tempFile, FileMode.Open, FileAccess.ReadWrite),
                ZipArchiveMode.Update
            )
        )
        {
            var entry = zip.GetEntry("xl/sharedStrings.xml")!;

            XDocument doc;
            using (var s = entry.Open())
                doc = XDocument.Load(s);

            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var si = doc.Descendants(ns + "si").First();

            // Insert phonetic runs in reverse order (sb=2 before sb=0) to trigger the bug.
            si.Add(
                new XElement(
                    ns + "rPh",
                    new XAttribute("sb", "2"),
                    new XAttribute("eb", "3"),
                    new XElement(ns + "t", "ト")
                )
            );
            si.Add(
                new XElement(
                    ns + "rPh",
                    new XAttribute("sb", "0"),
                    new XAttribute("eb", "2"),
                    new XElement(ns + "t", "テス")
                )
            );

            entry.Delete();
            var newEntry = zip.CreateEntry("xl/sharedStrings.xml");
            using var writer = new StreamWriter(newEntry.Open());
            writer.Write(doc.ToString());
        }

        return tempFile;
    }
}
