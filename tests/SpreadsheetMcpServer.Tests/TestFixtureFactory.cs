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
}
