using System.Globalization;
using ClosedXML.Excel;
using SpreadsheetMcpServer.Core.Models;
using SpreadsheetMcpServer.Core.Services;

namespace SpreadsheetMcpServer.Tests;

[Collection(BasePathCollectionDefinition.Name)]
public class SpreadsheetMcpServerUnitTest
{
    // The base directory is set by BasePathFixture, which the collection creates once; joining the
    // collection is what keeps that process-wide env var from being raced by another test class.
    private readonly IWorksheetService _worksheetService = new WorksheetService();
    private readonly ICellService _cellService = new CellService();
    private readonly ITableService _tableService = new TableService();
    private readonly IPictureService _pictureService = new PictureService();
    private readonly ICellFormatService _cellFormatService = new CellFormatService();

    /// <summary>
    /// Parses a two-column Address/Contents Markdown table (the shape of
    /// <see cref="MarkdownSheet.Contents"/>) into an insertion-ordered map of address → contents,
    /// so tests can assert on individual entries and on ordering. Header/separator rows are skipped
    /// and "\|" is unescaped to a literal "|".
    /// </summary>
    private static Dictionary<string, string> ParseContents(string table)
    {
        var map = new Dictionary<string, string>();
        foreach (var rawLine in table.Split('\n'))
        {
            string line = rawLine.Trim();
            if (line.Length == 0 || !line.StartsWith('|'))
                continue;

            // Split on unescaped "|" only, so an escaped "\|" stays within a cell.
            var cells = System.Text.RegularExpressions.Regex.Split(line.Trim('|'), @"(?<!\\)\|");
            if (cells.Length != 2)
                continue;

            string address = cells[0].Trim();
            if (
                address.Length == 0
                || address.Equals("Address", StringComparison.OrdinalIgnoreCase)
                || address.Trim('-', ' ').Length == 0
            )
                continue;

            map[address] = cells[1].Trim().Replace("\\|", "|");
        }
        return map;
    }

    /// <summary>Builds a two-column Address/Contents Markdown table from ordered pairs.</summary>
    private static string BuildContents(params (string Address, string Markdown)[] rows)
    {
        var sb = new System.Text.StringBuilder("| Address | Contents |\n| --- | --- |");
        foreach (var (address, markdown) in rows)
            sb.Append(CultureInfo.InvariantCulture, $"\n| {address} | {markdown.Replace("|", "\\|")} |");
        return sb.ToString();
    }

    [Fact]
    public void GetAllWorksheets_WithValidFile_ReturnsSheetNames()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            var result = _worksheetService.GetAllWorksheets(tempFile);

            Assert.Contains("Sheet1", result.Select(s => s.Name));
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ManageWorksheet_Create_AddsWorksheet()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            _worksheetService.ManageWorksheet(tempFile, "Sheet2", "create");

            var names = _worksheetService.GetAllWorksheets(tempFile).Select(s => s.Name);
            Assert.Contains("Sheet2", names);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ManageWorksheet_Create_DuplicateName_Throws()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            Assert.Throws<InvalidOperationException>(() =>
                _worksheetService.ManageWorksheet(tempFile, "Sheet1", "create")
            );
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ManageWorksheet_Delete_RemovesWorksheet()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            _worksheetService.ManageWorksheet(tempFile, "Sheet2", "create");
            _worksheetService.ManageWorksheet(tempFile, "Sheet2", "delete");

            var names = _worksheetService.GetAllWorksheets(tempFile).Select(s => s.Name);
            Assert.DoesNotContain("Sheet2", names);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ManageWorksheet_Delete_NonExistent_Throws()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            Assert.Throws<InvalidOperationException>(() =>
                _worksheetService.ManageWorksheet(tempFile, "Missing", "delete")
            );
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ManageWorksheet_Create_BrandNewFileNamedSheet1_Succeeds()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"create_new_{Guid.NewGuid():N}.xlsx");
        try
        {
            _worksheetService.ManageWorksheet(tempFile, "Sheet1", "create");

            var names = _worksheetService.GetAllWorksheets(tempFile).Select(s => s.Name);
            Assert.Equal(["Sheet1"], names);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ManageWorksheet_Rename_ChangesName()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            _worksheetService.ManageWorksheet(tempFile, "Sheet1", "rename", "Renamed");

            var names = _worksheetService.GetAllWorksheets(tempFile).Select(s => s.Name).ToList();
            Assert.Contains("Renamed", names);
            Assert.DoesNotContain("Sheet1", names);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ManageWorksheet_Rename_NullNewName_Throws()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            Assert.ThrowsAny<ArgumentException>(() => _worksheetService.ManageWorksheet(tempFile, "Sheet1", "rename"));
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ManageWorksheet_UnknownAction_Throws()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            Assert.Throws<InvalidOperationException>(() =>
                _worksheetService.ManageWorksheet(tempFile, "Sheet1", "frobnicate")
            );
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void GetAllWorksheets_WithPicture_ReturnsPictureAnchorAddress()
    {
        string tempFile = TestFixtureFactory.CreateWorkbookWithPicture("Sheet1", "B2");
        try
        {
            var result = _worksheetService.GetAllWorksheets(tempFile);

            var sheet = result.Single(s => s.Name == "Sheet1");
            Assert.Equal(["B2"], sheet.PictureAddresses);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void GetAllWorksheets_WithoutPictures_ReturnsEmptyPictureAddresses()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            var result = _worksheetService.GetAllWorksheets(tempFile);

            var sheet = result.Single(s => s.Name == "Sheet1");
            Assert.Empty(sheet.PictureAddresses);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void GetFirstPictureInRange_AnchorInRange_ReturnsPicture()
    {
        string tempFile = TestFixtureFactory.CreateWorkbookWithPicture("Sheet1", "B2");
        try
        {
            var result = _pictureService.GetFirstPictureInRange(tempFile, "Sheet1", "A1:C5");

            Assert.NotNull(result);
            Assert.Equal("image/png", result.MimeType);
            Assert.Equal("B2", result.AnchorAddress);
            Assert.Equal(TestFixtureFactory.OnePixelPng, Convert.FromBase64String(result.Base64));
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void GetFirstPictureInRange_AnchorOutsideRange_ReturnsNull()
    {
        string tempFile = TestFixtureFactory.CreateWorkbookWithPicture("Sheet1", "B2");
        try
        {
            var result = _pictureService.GetFirstPictureInRange(tempFile, "Sheet1", "D1:F5");

            Assert.Null(result);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void PastePictures_InsertsPicture_RoundTripsViaGetFirstPictureInRange()
    {
        string workbook = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        string image = TestFixtureFactory.CreateTempPngFile();
        try
        {
            var specs = new List<PicturePasteSpec>
            {
                new()
                {
                    PicturePath = image,
                    TargetSheet = "Sheet1",
                    TopLeftAddress = "B2",
                },
            };

            _pictureService.PastePictures(workbook, specs);

            var result = _pictureService.GetFirstPictureInRange(workbook, "Sheet1", "A1:C5");
            Assert.NotNull(result);
            Assert.Equal("B2", result.AnchorAddress);
            Assert.Equal("image/png", result.MimeType);
            Assert.Equal(TestFixtureFactory.OnePixelPng, Convert.FromBase64String(result.Base64));
        }
        finally
        {
            File.Delete(workbook);
            File.Delete(image);
        }
    }

    [Fact]
    public void PastePictures_InsertsMultiplePictures_InOneCall()
    {
        string workbook = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        string image = TestFixtureFactory.CreateTempPngFile();
        try
        {
            var specs = new List<PicturePasteSpec>
            {
                new()
                {
                    PicturePath = image,
                    TargetSheet = "Sheet1",
                    TopLeftAddress = "B2",
                },
                new()
                {
                    PicturePath = image,
                    TargetSheet = "Sheet1",
                    TopLeftAddress = "E5",
                },
            };

            _pictureService.PastePictures(workbook, specs);

            using var wb = new XLWorkbook(workbook);
            Assert.Equal(2, wb.Worksheet("Sheet1").Pictures.Count);
        }
        finally
        {
            File.Delete(workbook);
            File.Delete(image);
        }
    }

    [Fact]
    public void PastePictures_WithExplicitDimensions_AppliesPixelSize()
    {
        string workbook = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        string image = TestFixtureFactory.CreateTempPngFile();
        try
        {
            var specs = new List<PicturePasteSpec>
            {
                new()
                {
                    PicturePath = image,
                    TargetSheet = "Sheet1",
                    TopLeftAddress = "B2",
                    Width = 120,
                    Height = 90,
                },
            };

            _pictureService.PastePictures(workbook, specs);

            using var wb = new XLWorkbook(workbook);
            var picture = wb.Worksheet("Sheet1").Pictures.Single();
            Assert.Equal(120, picture.Width);
            Assert.Equal(90, picture.Height);
        }
        finally
        {
            File.Delete(workbook);
            File.Delete(image);
        }
    }

    [Fact]
    public void PastePictures_WithoutDimensions_KeepsOriginalSize()
    {
        string workbook = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        string image = TestFixtureFactory.CreateTempPngFile();
        try
        {
            var specs = new List<PicturePasteSpec>
            {
                new()
                {
                    PicturePath = image,
                    TargetSheet = "Sheet1",
                    TopLeftAddress = "B2",
                },
            };

            _pictureService.PastePictures(workbook, specs);

            using var wb = new XLWorkbook(workbook);
            var picture = wb.Worksheet("Sheet1").Pictures.Single();
            // The source is a 1x1 PNG, so the original size is preserved.
            Assert.Equal(1, picture.Width);
            Assert.Equal(1, picture.Height);
        }
        finally
        {
            File.Delete(workbook);
            File.Delete(image);
        }
    }

    [Fact]
    public void PastePictures_RelativePicturePath_ResolvesAgainstBasePath()
    {
        string workbook = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        string image = TestFixtureFactory.CreateTempPngFile();
        string? previousBasePath = Environment.GetEnvironmentVariable("SPREADSHEET_BASE_PATH");
        try
        {
            Environment.SetEnvironmentVariable("SPREADSHEET_BASE_PATH", Path.GetDirectoryName(image));

            var specs = new List<PicturePasteSpec>
            {
                new()
                {
                    PicturePath = Path.GetFileName(image),
                    TargetSheet = "Sheet1",
                    TopLeftAddress = "B2",
                },
            };

            _pictureService.PastePictures(workbook, specs);

            var result = _pictureService.GetFirstPictureInRange(workbook, "Sheet1", "A1:C5");
            Assert.NotNull(result);
            Assert.Equal("B2", result.AnchorAddress);
        }
        finally
        {
            Environment.SetEnvironmentVariable("SPREADSHEET_BASE_PATH", previousBasePath);
            File.Delete(workbook);
            File.Delete(image);
        }
    }

    [Fact]
    public void PastePictures_MissingImageFile_Throws()
    {
        string workbook = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            var specs = new List<PicturePasteSpec>
            {
                new()
                {
                    PicturePath = Path.Combine(Path.GetTempPath(), $"missing_{Guid.NewGuid():N}.png"),
                    TargetSheet = "Sheet1",
                    TopLeftAddress = "B2",
                },
            };

            Assert.Throws<InvalidOperationException>(() => _pictureService.PastePictures(workbook, specs));
        }
        finally
        {
            File.Delete(workbook);
        }
    }

    [Fact]
    public void PastePictures_UnknownTargetSheet_Throws()
    {
        string workbook = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        string image = TestFixtureFactory.CreateTempPngFile();
        try
        {
            var specs = new List<PicturePasteSpec>
            {
                new()
                {
                    PicturePath = image,
                    TargetSheet = "NoSuchSheet",
                    TopLeftAddress = "B2",
                },
            };

            Assert.Throws<InvalidOperationException>(() => _pictureService.PastePictures(workbook, specs));
        }
        finally
        {
            File.Delete(workbook);
            File.Delete(image);
        }
    }

    [Fact]
    public void LoadRange_WithValidFile_ReturnsCells()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            MarkdownSheet result = _cellService.LoadRange(tempFile, "Sheet1", "A1:Q60");
            var contents = ParseContents(result.Contents);

            Assert.Equal("Sheet1", result.SheetName);
            Assert.True(contents.Count > 0);
            Assert.Equal("Hello", contents["A1"]);
            Assert.Equal("42", contents["C1"]);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ManageTables_Create_CreatesTableWithExpectedProperties()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            var table = new SerializableTable
            {
                Name = "SalesTable",
                Reference = "A1:C3",
                StyleName = "TableStyleMedium9",
                ShowRowStripes = true,
                Columns = ["Hello", "World", "Value"],
            };

            _tableService.ManageTables(tempFile, "Sheet1", "create", table: table);

            List<SerializableTable> result = _tableService.ManageTables(tempFile, "Sheet1", "get");

            Assert.Single(result);
            Assert.Equal("SalesTable", result[0].Name);
            Assert.Equal("A1:C3", result[0].Reference);
            Assert.Equal("TableStyleMedium9", result[0].StyleName);
            Assert.True(result[0].ShowRowStripes);
            Assert.Equal(["Hello", "World", "Value"], result[0].Columns);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ManageTables_Delete_RemovesTable()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            _tableService.ManageTables(
                tempFile,
                "Sheet1",
                "create",
                table: new SerializableTable
                {
                    Name = "ToDelete",
                    Reference = "A1:C1",
                    Columns = ["Hello", "World", "Value"],
                }
            );

            _tableService.ManageTables(tempFile, "Sheet1", "delete", tableName: "ToDelete");

            List<SerializableTable> result = _tableService.ManageTables(tempFile, "Sheet1", "get");
            Assert.Empty(result);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ManageTables_Get_ReturnsEmptyWhenNoTables()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            List<SerializableTable> result = _tableService.ManageTables(tempFile, "Sheet1", "get");
            Assert.Empty(result);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ManageTables_UnknownAction_Throws()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            Assert.Throws<InvalidOperationException>(() =>
                _tableService.ManageTables(tempFile, "Sheet1", "frobnicate")
            );
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ManageTables_Delete_NonExistentTable_Throws()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            Assert.Throws<InvalidOperationException>(() =>
                _tableService.ManageTables(tempFile, "Sheet1", "delete", tableName: "Missing")
            );
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void UpdateRange_WritesMarkdownAndRoundTripsCorrectly()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            const string sheetName = "Sheet1";

            var sheet = new MarkdownSheet
            {
                SheetName = sheetName,
                Contents = BuildContents(
                    ("A1", "plain text"),
                    ("A2", "sample **bold**"),
                    ("A3", "an *italic* word"),
                    ("A4", "a ~~strike~~ word"),
                    ("D5:E6", "Merged string")
                ),
            };

            _cellService.UpdateRange(tempFile, sheet);

            MarkdownSheet result = _cellService.LoadRange(tempFile, sheetName, "A1:Q60");
            var contents = ParseContents(result.Contents);

            // Inline Markdown round-trips through rich text back to the same Markdown.
            Assert.Equal("plain text", contents["A1"]);
            Assert.Equal("sample **bold**", contents["A2"]);
            Assert.Equal("an *italic* word", contents["A3"]);
            Assert.Equal("a ~~strike~~ word", contents["A4"]);

            // The merged-range key is merged and emitted under its full range address.
            Assert.Equal("Merged string", contents["D5:E6"]);
            Assert.False(contents.ContainsKey("D5"));
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void LoadRange_WholeCellFormatting_ConvertsToMarkdown()
    {
        // CreateFormattedWorkbook applies bold + italic + strikethrough to the whole cell A1 = "Styled".
        string tempFile = TestFixtureFactory.CreateFormattedWorkbook("Sheet1");
        try
        {
            MarkdownSheet result = _cellService.LoadRange(tempFile, "Sheet1", "A1:Q60");
            var contents = ParseContents(result.Contents);

            // Strikethrough outermost, then bold, then italic.
            Assert.Equal("~~***Styled***~~", contents["A1"]);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void LoadRange_WithOutOfOrderPhoneticRuns_ReturnsCorrectCells()
    {
        string tempFile = TestFixtureFactory.CreateWorkbookWithOutOfOrderPhoneticRuns("Sheet1");
        try
        {
            // Would throw "Phonetic runs must be in ascending order" without sanitization.
            MarkdownSheet result = _cellService.LoadRange(tempFile, "Sheet1", "A1:Q60");
            var contents = ParseContents(result.Contents);

            Assert.Single(contents);
            Assert.Equal("テスト", contents["A1"]);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void LoadRange_WithCustomRange_ReturnsOnlyCellsInRange()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.Cell("A1").Value = "InRange";
                ws.Cell("B2").Value = "InRange2";
                ws.Cell("D4").Value = "OutOfRange";
                workbook.SaveAs(tempFile);
            }

            MarkdownSheet result = _cellService.LoadRange(tempFile, "Sheet1", "A1:C3");
            var contents = ParseContents(result.Contents);

            Assert.Equal("InRange", contents["A1"]);
            Assert.Equal("InRange2", contents["B2"]);
            Assert.False(contents.ContainsKey("D4"));
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void LoadRange_UsedRangeSmallerThanDefault_ReadsOnlyUsedRange()
    {
        // SimpleWorkbook has data only in A1:C1, well within the default A1:Q60.
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            MarkdownSheet result = _cellService.LoadRange(tempFile, "Sheet1", "A1:Q60");

            // Should contain the three cells in A1:C1 — no extras.
            Assert.Equal(3, ParseContents(result.Contents).Count);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void LoadRange_EmptyCellsAreExcluded()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.Cell("A1").Value = "Value";
                ws.Cell("B1").Value = ""; // empty string — should be excluded
                ws.Cell("C1").Value = 0; // zero is a valid value — should be included
                workbook.SaveAs(tempFile);
            }

            MarkdownSheet result = _cellService.LoadRange(tempFile, "Sheet1", "A1:Q60");
            var contents = ParseContents(result.Contents);

            Assert.True(contents.ContainsKey("A1"));
            Assert.False(contents.ContainsKey("B1"));
            Assert.Equal("0", contents["C1"]);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void LoadRange_BlankSheet_ReturnsEmptyCells()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                workbook.Worksheets.Add("Sheet1");
                workbook.SaveAs(tempFile);
            }

            MarkdownSheet result = _cellService.LoadRange(tempFile, "Sheet1", "A1:Q60");

            Assert.Equal("Sheet1", result.SheetName);
            Assert.Empty(ParseContents(result.Contents));
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void LoadRange_WithThreshold_SplitsIntoMultipleChunks()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                for (int i = 1; i <= 20; i++)
                    ws.Cell(i, 1).Value = $"Value{i}";
                workbook.SaveAs(tempFile);
            }

            // Use a very small truncate limit to drop trailing cells.
            MarkdownSheet result = _cellService.LoadRange(tempFile, "Sheet1", "A1:A20", truncate: 100);
            var contents = ParseContents(result.Contents);

            Assert.True(contents.Count > 0, "Expected some cells within threshold.");
            Assert.True(contents.Count < 20, "Expected truncation with a small threshold.");
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void LoadRange_WithThreshold_NoChunkingWhenResultFits()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            // Large truncate limit — all cells fit without truncation.
            MarkdownSheet result = _cellService.LoadRange(tempFile, "Sheet1", "A1:Q60", truncate: 1_000_000);

            Assert.NotEmpty(ParseContents(result.Contents));
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void LoadRange_WithThreshold_LargeThresholdReturnsAllCells()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                for (int i = 1; i <= 10; i++)
                    ws.Cell(i, 1).Value = $"Row{i}";
                workbook.SaveAs(tempFile);
            }

            MarkdownSheet result = _cellService.LoadRange(tempFile, "Sheet1", "A1:A10", truncate: 1_000_000);

            Assert.Equal(10, ParseContents(result.Contents).Count);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void LoadRange_WithMergedCells_KeysEntryByFullRangeAddress()
    {
        string tempFile = TestFixtureFactory.CreateWorkbookWithMergedCells("Sheet1");
        try
        {
            MarkdownSheet result = _cellService.LoadRange(tempFile, "Sheet1", "A1:Q60");
            var contents = ParseContents(result.Contents);

            // The merged region is emitted once, keyed by its full range address.
            Assert.Equal("Merged", contents["A1:C2"]);

            // The anchor's individual address is not emitted separately.
            Assert.False(contents.ContainsKey("A1"));

            // D1 is a plain cell — keyed by its own address.
            Assert.Equal("Solo", contents["D1"]);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void LoadRange_ReturnsFlattenedMarkdownMap()
    {
        string tempFile = TestFixtureFactory.CreateWorkbookForMarkdown("Sheet1");
        try
        {
            MarkdownSheet result = _cellService.LoadRange(tempFile, "Sheet1", "A1:Q60");
            var contents = ParseContents(result.Contents);

            Assert.Equal("Sheet1", result.SheetName);

            // Plain text — no markers.
            Assert.Equal("Plain", contents["A1"]);

            // Partial rich text — only the bold run is wrapped.
            Assert.Equal("sample **string**", contents["A2"]);

            // Whole-cell italic and strikethrough.
            Assert.Equal("*ItalicAll*", contents["A3"]);
            Assert.Equal("~~StruckAll~~", contents["A4"]);

            // Empty cell is excluded.
            Assert.False(contents.ContainsKey("A5"));

            // Merged region keyed by its full range address; member cell not emitted separately.
            Assert.Equal("MergedVal", contents["B1:C1"]);
            Assert.False(contents.ContainsKey("B1"));
            Assert.False(contents.ContainsKey("C1"));
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ManageTables_Create_CreatesTableWithAutoFilterAndStyle()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            var table = new SerializableTable
            {
                Name = "TestTable",
                Reference = "A1:C1",
                StyleName = "TableStyleMedium9",
                ShowRowStripes = true,
                Columns = ["Hello", "World", "Value"],
            };

            _tableService.ManageTables(tempFile, "Sheet1", "create", table: table);

            using var workbook = new XLWorkbook(tempFile);
            var worksheet = workbook.Worksheet("Sheet1");
            var createdTable = worksheet.Table("TestTable");

            Assert.NotNull(createdTable);
            Assert.True(createdTable.ShowAutoFilter);
            Assert.Equal("TableStyleMedium9", createdTable.Theme.Name);
            Assert.True(createdTable.ShowRowStripes);
            Assert.Equal(["Hello", "World", "Value"], createdTable.Fields.Select(f => f.Name).ToList());
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void LoadRange_OrdersEntriesByRowThenColumn_Naturally()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                // Write in deliberately out-of-order positions, including A10 before A2 so a
                // lexicographic sort ("A10" < "A2") would order them incorrectly.
                ws.Cell("B1").Value = "B1";
                ws.Cell("A10").Value = "A10";
                ws.Cell("A1").Value = "A1";
                ws.Cell("A2").Value = "A2";
                workbook.SaveAs(tempFile);
            }

            MarkdownSheet result = _cellService.LoadRange(tempFile, "Sheet1", "A1:Q60");

            // Row first, then column, with natural (numeric) row ordering: A1, B1, A2, A10.
            Assert.Equal(["A1", "B1", "A2", "A10"], ParseContents(result.Contents).Keys.ToList());
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void LoadRange_OrdersMergedRangeByAnchorAddress()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.Cell("A1").Value = "A1";
                // Merged region anchored at A3 — should sort by its anchor (A3), i.e. after A1/B1
                // and before A5, not at the end despite the long range key.
                ws.Cell("A3").Value = "MergedA3";
                ws.Range("A3:C4").Merge();
                ws.Cell("B1").Value = "B1";
                ws.Cell("A5").Value = "A5";
                workbook.SaveAs(tempFile);
            }

            MarkdownSheet result = _cellService.LoadRange(tempFile, "Sheet1", "A1:Q60");

            // Anchor A3 places the merged key between B1 (row 1) and A5 (row 5).
            Assert.Equal(["A1", "B1", "A3:C4", "A5"], ParseContents(result.Contents).Keys.ToList());
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void LoadRange_TruncateStopsBeforeEntryThatWouldExceedLimit()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.Cell("A1").Value = "A1value";
                ws.Cell("A2").Value = "A2value";
                ws.Cell("A3").Value = "A3value";
                workbook.SaveAs(tempFile);
            }

            // Pick a limit (in rendered-table characters) that fits the header + A1 + A2 but not A3.
            // The rendered table is the header followed by one "\n| <addr> | <value> |" row per cell.
            int headerLen = BuildContents().Length;
            int rowA1 = BuildContents(("A1", "A1value")).Length - headerLen;
            int rowA2 = BuildContents(("A2", "A2value")).Length - headerLen;
            int limit = headerLen + rowA1 + rowA2; // including A3 would exceed this.

            MarkdownSheet result = _cellService.LoadRange(tempFile, "Sheet1", "A1:A3", truncate: limit);

            // Ordered A1, A2, A3 — including A3 would push past the limit, so only A1 and A2 return.
            Assert.Equal(["A1", "A2"], ParseContents(result.Contents).Keys.ToList());
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void LoadRange_TruncateNonPositive_ReturnsEverything()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                for (int i = 1; i <= 5; i++)
                    ws.Cell(i, 1).Value = $"Value{i}";
                workbook.SaveAs(tempFile);
            }

            MarkdownSheet zero = _cellService.LoadRange(tempFile, "Sheet1", "A1:A5", truncate: 0);
            MarkdownSheet negative = _cellService.LoadRange(tempFile, "Sheet1", "A1:A5", truncate: -1);

            Assert.Equal(5, ParseContents(zero.Contents).Count);
            Assert.Equal(5, ParseContents(negative.Contents).Count);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void LoadRange_TargetRange_ReflectsEffectiveRangeReadWhenSmallerThanRequested()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                // Used range is A1:B2 — much smaller than the requested A1:Q60.
                ws.Cell("A1").Value = "A1";
                ws.Cell("B2").Value = "B2";
                workbook.SaveAs(tempFile);
            }

            MarkdownSheet result = _cellService.LoadRange(tempFile, "Sheet1", "A1:Q60");

            // TargetRange is the tight bounding box of the emitted cells, not the request.
            Assert.Equal("A1:B2", result.TargetRange);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void LoadRange_TargetRange_IsBoundingBoxOfEmittedCellsNotUsedRange()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                // Two clustered cells plus a far-flung one stretch the used range out to U92,
                // but the requested A1:Q60 only contains the first two — so the box must be A1:B3.
                ws.Cell("A1").Value = "A1";
                ws.Cell("B3").Value = "B3";
                ws.Cell("U92").Value = "FarAway";
                workbook.SaveAs(tempFile);
            }

            MarkdownSheet result = _cellService.LoadRange(tempFile, "Sheet1", "A1:Q60");

            Assert.Equal("A1:B3", result.TargetRange);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void LoadRange_TargetRange_ShrinksWhenTruncationDropsTrailingRows()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.Cell("A1").Value = "A1";
                ws.Cell("B3").Value = "B3";
                ws.Cell("C50").Value = "C50";
                workbook.SaveAs(tempFile);
            }

            // A tight truncation budget keeps only the first couple of rows, so the trailing C50
            // is dropped and the reported box must shrink to exclude it.
            MarkdownSheet result = _cellService.LoadRange(tempFile, "Sheet1", "A1:Q60", 60);

            var contents = ParseContents(result.Contents);
            Assert.DoesNotContain("C50", contents.Keys);
            Assert.Equal("A1:B3", result.TargetRange);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void LoadRange_TargetRange_EchoesRequestedRangeWhenNoUsedRange()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                workbook.Worksheets.Add("Sheet1"); // blank sheet — no used range.
                workbook.SaveAs(tempFile);
            }

            MarkdownSheet result = _cellService.LoadRange(tempFile, "Sheet1", "A1:Q60");

            Assert.Empty(ParseContents(result.Contents));
            // No used range — TargetRange echoes the requested range.
            Assert.Equal("A1:Q60", result.TargetRange);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void LoadRange_TargetRange_EchoesRequestedRangeWhenIntersectionEmpty()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                // Data lives far away from the requested range, so the intersection is empty.
                ws.Cell("Z90").Value = "FarAway";
                workbook.SaveAs(tempFile);
            }

            MarkdownSheet result = _cellService.LoadRange(tempFile, "Sheet1", "A1:C3");

            Assert.Empty(ParseContents(result.Contents));
            Assert.Equal("A1:C3", result.TargetRange);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void UpdateRange_EscapesPipeInContents_AndRoundTrips()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            const string sheetName = "Sheet1";
            const string valueWithPipe = "a | b | c";

            var sheet = new MarkdownSheet { SheetName = sheetName, Contents = BuildContents(("A1", valueWithPipe)) };

            // The serialized table escapes the literal pipes so they do not split the row.
            Assert.Contains("a \\| b \\| c", sheet.Contents);

            _cellService.UpdateRange(tempFile, sheet);
            MarkdownSheet result = _cellService.LoadRange(tempFile, sheetName, "A1:Q60");

            // The literal pipes survive the round trip and the emitted table re-escapes them.
            Assert.Contains("a \\| b \\| c", result.Contents);
            Assert.Equal(valueWithPipe, ParseContents(result.Contents)["A1"]);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void UpdateRange_WritesNumbersAsNumericCells()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            _cellService.UpdateRange(
                tempFile,
                new MarkdownSheet
                {
                    SheetName = "Sheet1",
                    Contents = BuildContents(("A1", "42"), ("A2", "3.14159"), ("A3", "-7")),
                }
            );

            using var workbook = new XLWorkbook(tempFile);
            var ws = workbook.Worksheet("Sheet1");
            Assert.Equal(XLDataType.Number, ws.Cell("A1").DataType);
            Assert.Equal(42.0, ws.Cell("A1").GetDouble());
            Assert.Equal(XLDataType.Number, ws.Cell("A2").DataType);
            Assert.Equal(3.14159, ws.Cell("A2").GetDouble());
            Assert.Equal(XLDataType.Number, ws.Cell("A3").DataType);
            Assert.Equal(-7.0, ws.Cell("A3").GetDouble());
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void UpdateRange_WritesFormulaCell()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            using (var workbook = new XLWorkbook(tempFile))
            {
                var ws = workbook.Worksheet("Sheet1");
                ws.Cell("A1").Value = 20;
                ws.Cell("A2").Value = 22;
                workbook.Save();
            }

            _cellService.UpdateRange(
                tempFile,
                new MarkdownSheet { SheetName = "Sheet1", Contents = BuildContents(("A3", "=SUM(A1:A2)")) }
            );

            using var readWorkbook = new XLWorkbook(tempFile);
            var readWs = readWorkbook.Worksheet("Sheet1");
            Assert.True(readWs.Cell("A3").HasFormula);
            Assert.Equal("SUM(A1:A2)", readWs.Cell("A3").FormulaA1);
            Assert.Equal(42.0, readWs.Cell("A3").GetDouble());
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void UpdateRange_WritesBooleanAndDateCells()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            _cellService.UpdateRange(
                tempFile,
                new MarkdownSheet
                {
                    SheetName = "Sheet1",
                    Contents = BuildContents(("A1", "true"), ("A2", "FALSE"), ("A3", "2026-09-07")),
                }
            );

            using var workbook = new XLWorkbook(tempFile);
            var ws = workbook.Worksheet("Sheet1");
            Assert.Equal(XLDataType.Boolean, ws.Cell("A1").DataType);
            Assert.True(ws.Cell("A1").GetBoolean());
            Assert.Equal(XLDataType.Boolean, ws.Cell("A2").DataType);
            Assert.False(ws.Cell("A2").GetBoolean());
            Assert.Equal(XLDataType.DateTime, ws.Cell("A3").DataType);
            Assert.Equal(new DateTime(2026, 9, 7), ws.Cell("A3").GetDateTime());
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void UpdateRange_ApostrophePrefix_WritesLiteralText()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            _cellService.UpdateRange(
                tempFile,
                new MarkdownSheet
                {
                    SheetName = "Sheet1",
                    Contents = BuildContents(("A1", "'=SUM(A1:A2)"), ("A2", "'007"), ("A3", "'true")),
                }
            );

            using var workbook = new XLWorkbook(tempFile);
            var ws = workbook.Worksheet("Sheet1");
            Assert.Equal(XLDataType.Text, ws.Cell("A1").DataType);
            Assert.Equal("=SUM(A1:A2)", ws.Cell("A1").GetString());
            Assert.False(ws.Cell("A1").HasFormula);
            Assert.Equal(XLDataType.Text, ws.Cell("A2").DataType);
            Assert.Equal("007", ws.Cell("A2").GetString());
            Assert.Equal(XLDataType.Text, ws.Cell("A3").DataType);
            Assert.Equal("true", ws.Cell("A3").GetString());
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ClearRange_RemovesContentsButKeepsFormatting()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.Cell("A1").Value = "Hello";
                ws.Cell("A1").Style.Font.Bold = true;
                workbook.SaveAs(tempFile);
            }

            string message = _cellService.ClearRange(tempFile, "Sheet1", "A1:B10");

            using var readWorkbook = new XLWorkbook(tempFile);
            var readWs = readWorkbook.Worksheet("Sheet1");
            Assert.Contains("A1", message);
            Assert.True(readWs.Cell("A1").IsEmpty());
            Assert.True(readWs.Cell("A1").Style.Font.Bold);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ClearRange_ClearFormats_AlsoRemovesFormatting()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.Cell("A1").Value = "Hello";
                ws.Cell("A1").Style.Font.Bold = true;
                workbook.SaveAs(tempFile);
            }

            _cellService.ClearRange(tempFile, "Sheet1", "A1:B10", clearFormats: true);

            using var readWorkbook = new XLWorkbook(tempFile);
            var readWs = readWorkbook.Worksheet("Sheet1");
            Assert.True(readWs.Cell("A1").IsEmpty());
            Assert.False(readWs.Cell("A1").Style.Font.Bold);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void LoadRangeInMarkdownTable_SimpleGrid_LaysOutValuesByPosition()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.Cell("A1").Value = "Tom";
                ws.Cell("B1").Value = "John";
                ws.Cell("A2").Value = "Matthew";
                ws.Cell("B2").Value = "Luke";
                workbook.SaveAs(tempFile);
            }

            MarkdownSheet result = _cellService.LoadRangeInMarkdownTable(tempFile, "Sheet1", "A1:Q60");

            const string expected = "| | A | B |\n| --- | --- | --- |\n| 1 | Tom | John |\n| 2 | Matthew | Luke |";
            Assert.Equal("Sheet1", result.SheetName);
            Assert.Equal(expected, result.Contents);
            Assert.Equal("A1:B2", result.TargetRange);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void LoadRangeInMarkdownTable_MergedAndFormattedAndBlankInterior()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                // A1:B1 merged with value "Header"; A2 empty; B2 bold "Luke"; A3 "End".
                ws.Cell("A1").Value = "Header";
                ws.Range("A1:B1").Merge();
                var b2 = ws.Cell("B2");
                b2.Value = "Luke";
                b2.Style.Font.Bold = true;
                ws.Cell("A3").Value = "End";
                workbook.SaveAs(tempFile);
            }

            MarkdownSheet result = _cellService.LoadRangeInMarkdownTable(tempFile, "Sheet1", "A1:Q60");

            const string expected =
                "| | A | B |\n| --- | --- | --- |\n| 1 | Header |  |\n| 2 |  | **Luke** |\n| 3 | End |  |";
            Assert.Equal(expected, result.Contents);
            Assert.Equal("A1:B3", result.TargetRange);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void LoadRangeInMarkdownTable_DropsTrailingEmptyRowsAndColumns_KeepsInteriorBlanks()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                // Content forms a tight box A1:C2 (with an interior gap at B1/B2), while a far cell
                // F10 stretches the used range — the requested range excludes it.
                ws.Cell("A1").Value = "A1";
                ws.Cell("C1").Value = "C1";
                ws.Cell("A2").Value = "A2";
                ws.Cell("C2").Value = "C2";
                ws.Cell("F10").Value = "Far";
                workbook.SaveAs(tempFile);
            }

            MarkdownSheet result = _cellService.LoadRangeInMarkdownTable(tempFile, "Sheet1", "A1:E5");

            const string expected =
                "| | A | B | C |\n| --- | --- | --- | --- |\n| 1 | A1 |  | C1 |\n| 2 | A2 |  | C2 |";
            Assert.Equal(expected, result.Contents);
            Assert.Equal("A1:C2", result.TargetRange);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void LoadRangeInMarkdownTable_EscapesPipeAndCollapsesNewlines()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.Cell("A1").Value = "a | b";
                ws.Cell("B1").Value = "line1\nline2";
                workbook.SaveAs(tempFile);
            }

            MarkdownSheet result = _cellService.LoadRangeInMarkdownTable(tempFile, "Sheet1", "A1:B1");

            Assert.Equal("| | A | B |\n| --- | --- | --- |\n| 1 | a \\| b | line1 line2 |", result.Contents);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void LoadRangeInMarkdownTable_TruncateDropsTrailingRows()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                for (int i = 1; i <= 20; i++)
                    ws.Cell(i, 1).Value = $"Value{i}";
                workbook.SaveAs(tempFile);
            }

            MarkdownSheet result = _cellService.LoadRangeInMarkdownTable(tempFile, "Sheet1", "A1:A20", truncate: 80);

            // Body rows are "| <rowNumber> | … |"; the header and "--- " separator are excluded.
            int bodyRows = result.Contents.Split('\n').Count(l => l.Length > 2 && char.IsDigit(l[2]));
            Assert.True(bodyRows > 0, "Expected some rows within the budget.");
            Assert.True(bodyRows < 20, "Expected truncation with a small budget.");
            // TargetRange shrinks to the rows that survived.
            Assert.Equal($"A1:A{bodyRows}", result.TargetRange);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void LoadRangeInMarkdownTable_EmptyIntersection_EchoesRequestedRange()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.Cell("Z90").Value = "FarAway";
                workbook.SaveAs(tempFile);
            }

            MarkdownSheet result = _cellService.LoadRangeInMarkdownTable(tempFile, "Sheet1", "A1:C3");

            Assert.Equal(string.Empty, result.Contents);
            Assert.Equal("A1:C3", result.TargetRange);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void LoadRange_EmptyCellsProduceNoTableRow()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.Cell("A1").Value = "Value";
                ws.Cell("B1").Value = ""; // empty string — should produce no row
                ws.Cell("C1").Value = "Other";
                workbook.SaveAs(tempFile);
            }

            MarkdownSheet result = _cellService.LoadRange(tempFile, "Sheet1", "A1:Q60");

            // Exactly two data rows (A1, C1) on top of the header + separator — no blank-cell row.
            int dataRows =
                result.Contents.Split('\n').Count(l => l.TrimStart().StartsWith("| ", StringComparison.Ordinal)) - 2; // header row + separator row
            Assert.Equal(2, dataRows);
            Assert.DoesNotContain("| B1 |", result.Contents);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void FindStringsInSheets_MatchesAcrossMultipleSheets()
    {
        string tempFile = TestFixtureFactory.CreateMultiSheetWorkbook();
        try
        {
            List<MarkdownSheet> results = _cellService.FindStringsInSheets(new List<string> { "Hello" }, tempFile);

            // "Hello" matches A1 on Sheet1 and the substring in B1 on Sheet2.
            Assert.Equal(2, results.Count);

            var sheet1 = results.Single(r => r.SheetName == "Sheet1");
            var sheet2 = results.Single(r => r.SheetName == "Sheet2");

            Assert.Equal("Hello", ParseContents(sheet1.Contents)["A1"]);
            Assert.Equal("Hello again", ParseContents(sheet2.Contents)["B1"]);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void FindStringsInSheets_SubstringMatchIsCaseInsensitive()
    {
        string tempFile = TestFixtureFactory.CreateMultiSheetWorkbook();
        try
        {
            // "ell" is a substring of "Hello"; lower case should still match.
            List<MarkdownSheet> results = _cellService.FindStringsInSheets(new List<string> { "ELL" }, tempFile);

            var sheet1 = results.Single(r => r.SheetName == "Sheet1");
            Assert.Equal("Hello", ParseContents(sheet1.Contents)["A1"]);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void FindStringsInSheets_AnyNeedleMatches()
    {
        string tempFile = TestFixtureFactory.CreateMultiSheetWorkbook();
        try
        {
            List<MarkdownSheet> results = _cellService.FindStringsInSheets(
                new List<string> { "apple", "Goodbye" },
                tempFile
            );

            var sheet1 = ParseContents(results.Single(r => r.SheetName == "Sheet1").Contents);
            var sheet2 = ParseContents(results.Single(r => r.SheetName == "Sheet2").Contents);

            Assert.Equal("apple", sheet1["C1"]);
            Assert.Equal("Goodbye", sheet2["A1"]);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void FindStringsInSheets_NoMatch_ReturnsEmptyList()
    {
        string tempFile = TestFixtureFactory.CreateMultiSheetWorkbook();
        try
        {
            List<MarkdownSheet> results = _cellService.FindStringsInSheets(
                new List<string> { "no-such-value" },
                tempFile
            );

            Assert.Empty(results);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void FindStringsInSheets_EmptyNeedleList_ReturnsEmptyList()
    {
        string tempFile = TestFixtureFactory.CreateMultiSheetWorkbook();
        try
        {
            Assert.Empty(_cellService.FindStringsInSheets(new List<string>(), tempFile));
            Assert.Empty(_cellService.FindStringsInSheets(new List<string> { "" }, tempFile));
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void FindStringsInSheets_LeavesTargetRangeEmpty()
    {
        string tempFile = TestFixtureFactory.CreateMultiSheetWorkbook();
        try
        {
            List<MarkdownSheet> results = _cellService.FindStringsInSheets(new List<string> { "Hello" }, tempFile);

            Assert.NotEmpty(results);
            Assert.All(results, r => Assert.Equal(string.Empty, r.TargetRange));
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void FindStringsInSheets_ExcludesNonMatchingCells()
    {
        string tempFile = TestFixtureFactory.CreateMultiSheetWorkbook();
        try
        {
            // Only "apple" matches on Sheet1; "Hello" and "World" must not appear.
            List<MarkdownSheet> results = _cellService.FindStringsInSheets(new List<string> { "apple" }, tempFile);

            var sheet1 = ParseContents(results.Single(r => r.SheetName == "Sheet1").Contents);
            Assert.Single(sheet1);
            Assert.Equal("apple", sheet1["C1"]);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void FindStringsInSheets_SheetList_RestrictsSearchToNamedSheets()
    {
        string tempFile = TestFixtureFactory.CreateMultiSheetWorkbook();
        try
        {
            // "Hello" matches on both sheets, but only Sheet2 is requested.
            List<MarkdownSheet> results = _cellService.FindStringsInSheets(
                new List<string> { "Hello" },
                tempFile,
                new List<string> { "Sheet2" }
            );

            var only = Assert.Single(results);
            Assert.Equal("Sheet2", only.SheetName);
            Assert.Equal("Hello again", ParseContents(only.Contents)["B1"]);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void FindStringsInSheets_EmptySheetList_SearchesAllSheets()
    {
        string tempFile = TestFixtureFactory.CreateMultiSheetWorkbook();
        try
        {
            // An empty sheet list behaves the same as omitting it: search every worksheet.
            List<MarkdownSheet> results = _cellService.FindStringsInSheets(
                new List<string> { "Hello" },
                tempFile,
                new List<string>()
            );

            Assert.Equal(2, results.Count);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    /// <summary>
    /// Finds the single entry addressed <paramref name="address"/>, failing the test when absent.
    /// </summary>
    private static CellFormatSpec Entry(List<CellFormatSpec> specs, string address)
    {
        var match = specs.SingleOrDefault(f => f.Address == address);
        Assert.True(match != null, $"No entry for '{address}'. Got: {string.Join(", ", specs.Select(f => f.Address))}");
        return match!;
    }

    [Fact]
    public void ReadCellFormatting_OmitsDefaultStyledCells()
    {
        // CreateSimpleWorkbook writes values but never touches styling.
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            var result = _cellFormatService.ReadCellFormatting(tempFile, "Sheet1", "A1:Q60");

            Assert.Empty(result);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ReadCellFormatting_CoalescesAdjacentIdenticalCellsIntoRange()
    {
        string tempFile = TestFixtureFactory.CreateStyledWorkbookForFormatReading("Sheet1");
        try
        {
            var result = _cellFormatService.ReadCellFormatting(tempFile, "Sheet1", "A1:Q60");

            // The header run is one entry, not four; the number column is one entry, not nine.
            var header = Entry(result, "A1:D1");
            Assert.True(header.Bold);
            Assert.Equal("#FFFF00", header.BackgroundColor);
            Assert.Equal("Center", header.HorizontalAlignment);

            var numbers = Entry(result, "A2:A10");
            Assert.Equal("#,##0", numbers.NumberFormat);
            Assert.Equal("Thin", numbers.TopBorder?.Style);
            Assert.Equal("Thin", numbers.RightBorder?.Style);

            // A cell styled differently from its neighbours stays on its own.
            Assert.True(Entry(result, "B5").Italic);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ReadCellFormatting_ReportsRowHeightAndColumnWidth()
    {
        string tempFile = TestFixtureFactory.CreateStyledWorkbookForFormatReading("Sheet1");
        try
        {
            var result = _cellFormatService.ReadCellFormatting(tempFile, "Sheet1", "A1:Q60");

            // Row 1's height is addressed to the first column of the read range, column C's width to
            // the first row — the cells ApplyCellFormatting resolves back to that row/column.
            var height = result.Single(f => f.Height.HasValue);
            Assert.Equal("A1", height.Address);
            Assert.Equal(30, height.Height!.Value, 3);

            var width = result.Single(f => f.Width.HasValue);
            Assert.Equal("C1", width.Address);
            Assert.Equal(25, width.Width!.Value, 3);

            // Dimension entries come before the style entries.
            Assert.True(result.IndexOf(height) < result.FindIndex(f => f.Address == "A1:D1"));
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ReadCellFormatting_ReadsFontFillAlignment()
    {
        // A1 = "Styled": bold, italic, strikethrough, single underline, red font, yellow fill,
        // centered, wrap-text.
        string tempFile = TestFixtureFactory.CreateFormattedWorkbook("Sheet1");
        try
        {
            var result = _cellFormatService.ReadCellFormatting(tempFile, "Sheet1", "A1:Q60");

            var a1 = Entry(result, "A1");
            Assert.True(a1.Bold);
            Assert.True(a1.Italic);
            Assert.True(a1.Strikethrough);
            Assert.True(a1.Underline);
            Assert.Equal("#FF0000", a1.FontColor);
            Assert.Equal("#FFFF00", a1.BackgroundColor);
            Assert.Equal("Center", a1.HorizontalAlignment);
            Assert.True(a1.WrapText);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ReadCellFormatting_ReadsStylingOfCellWithNoContents()
    {
        // A cell can carry styling and no value at all — a colored banner row, a bordered box. The
        // used range must be computed from formats as well as contents, or these vanish.
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.Cell("A1").Value = "Anchor";
                ws.Cell("C3").Style.Fill.BackgroundColor = XLColor.Cyan; // styled, but empty
                workbook.SaveAs(tempFile);
            }

            var result = _cellFormatService.ReadCellFormatting(tempFile, "Sheet1", "A1:Q60");

            Assert.Equal("#00FFFF", Entry(result, "C3").BackgroundColor);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ReadCellFormatting_RoundTripsThroughApplyCellFormatting()
    {
        string source = TestFixtureFactory.CreateStyledWorkbookForFormatReading("Sheet1");
        string target = Path.Combine(Path.GetTempPath(), $"fixture_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                workbook.Worksheets.Add("Sheet1");
                workbook.SaveAs(target);
            }

            var original = _cellFormatService.ReadCellFormatting(source, "Sheet1", "A1:Q60");
            Assert.NotEmpty(original);

            // The read output is fed back verbatim — this is the contract the tool exists for.
            _cellFormatService.ApplyCellFormatting(target, original);

            var reapplied = _cellFormatService.ReadCellFormatting(target, "Sheet1", "A1:Q60");

            Assert.Equal(
                original.Select(f => f.Address).OrderBy(a => a, StringComparer.Ordinal),
                reapplied.Select(f => f.Address).OrderBy(a => a, StringComparer.Ordinal)
            );

            foreach (var before in original)
            {
                var after = Entry(reapplied, before.Address);
                Assert.Equal(before.Bold, after.Bold);
                Assert.Equal(before.Italic, after.Italic);
                Assert.Equal(before.Underline, after.Underline);
                Assert.Equal(before.Strikethrough, after.Strikethrough);
                Assert.Equal(before.FontColor, after.FontColor);
                Assert.Equal(before.FontName, after.FontName);
                Assert.Equal(before.FontSize, after.FontSize);
                Assert.Equal(before.BackgroundColor, after.BackgroundColor);
                Assert.Equal(before.HorizontalAlignment, after.HorizontalAlignment);
                Assert.Equal(before.VerticalAlignment, after.VerticalAlignment);
                Assert.Equal(before.WrapText, after.WrapText);
                Assert.Equal(before.NumberFormat, after.NumberFormat);
                Assert.Equal(before.TopBorder?.Style, after.TopBorder?.Style);
                Assert.Equal(before.BottomBorder?.Style, after.BottomBorder?.Style);
                Assert.Equal(before.LeftBorder?.Style, after.LeftBorder?.Style);
                Assert.Equal(before.RightBorder?.Style, after.RightBorder?.Style);
                Assert.Equal(before.Height, after.Height);
                Assert.Equal(before.Width, after.Width);
            }
        }
        finally
        {
            File.Delete(source);
            File.Delete(target);
        }
    }

    [Fact]
    public void ReadCellFormatting_TruncateCapsEntryCount()
    {
        string tempFile = TestFixtureFactory.CreateStyledWorkbookForFormatReading("Sheet1");
        try
        {
            var all = _cellFormatService.ReadCellFormatting(tempFile, "Sheet1", "A1:Q60", truncate: 0);
            Assert.True(all.Count > 2);

            var capped = _cellFormatService.ReadCellFormatting(tempFile, "Sheet1", "A1:Q60", truncate: 2);

            Assert.Equal(2, capped.Count);
            Assert.Equal(all.Take(2).Select(f => f.Address), capped.Select(f => f.Address));
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ReadCellFormatting_RangeSmallerThanUsedRange_ReadsOnlyRequestedRange()
    {
        string tempFile = TestFixtureFactory.CreateStyledWorkbookForFormatReading("Sheet1");
        try
        {
            var result = _cellFormatService.ReadCellFormatting(tempFile, "Sheet1", "A1:B1", truncate: 0);

            // The header run is clipped to the requested range, and the number column is outside it.
            Assert.Contains(result, f => f.Address == "A1:B1");
            Assert.DoesNotContain(result, f => f.Address == "A2:A10");
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ReadCellFormatting_UnknownSheet_Throws()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            Assert.Throws<InvalidOperationException>(() =>
                _cellFormatService.ReadCellFormatting(tempFile, "NoSuchSheet", "A1:Q60")
            );
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ApplyCellFormatting_BorderStyleWithoutExplicitCase_RoundTrips()
    {
        // MediumDashDotDot has no hand-written case in ParseBorderStyle; it must still survive being
        // read and applied back, rather than degrading to no border.
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            _cellFormatService.ApplyCellFormatting(
                tempFile,
                [
                    new CellFormatSpec
                    {
                        Address = "A1",
                        TopBorder = new BorderSpec { Style = "MediumDashDotDot" },
                    },
                ]
            );

            var result = _cellFormatService.ReadCellFormatting(tempFile, "Sheet1", "A1:Q60");

            Assert.Equal("MediumDashDotDot", Entry(result, "A1").TopBorder?.Style);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ApplyCellFormatting_AppliesFontFillAlignmentAndDimensions()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            _cellFormatService.ApplyCellFormatting(
                tempFile,
                [
                    new CellFormatSpec
                    {
                        Address = "A1:B2",
                        Bold = true,
                        FontColor = "#FF0000",
                        BackgroundColor = "#00FF00",
                        HorizontalAlignment = "Right",
                        NumberFormat = "0.00",
                    },
                    new CellFormatSpec { Address = "A1", Height = 40 },
                ]
            );

            using var workbook = new XLWorkbook(tempFile);
            var ws = workbook.Worksheet("Sheet1");

            Assert.True(ws.Cell("B2").Style.Font.Bold);
            Assert.Equal(XLColor.FromHtml("#FF0000"), ws.Cell("A1").Style.Font.FontColor);
            Assert.Equal(XLColor.FromHtml("#00FF00"), ws.Cell("A1").Style.Fill.BackgroundColor);
            Assert.Equal(XLAlignmentHorizontalValues.Right, ws.Cell("A1").Style.Alignment.Horizontal);
            Assert.Equal("0.00", ws.Cell("A1").Style.NumberFormat.Format);
            Assert.Equal(40, ws.Row(1).Height, 3);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void PathResolver_GetWorkingDirectory_UsesBasePathWhenSet()
    {
        string baseDir = Path.Combine(Path.GetTempPath(), $"base_dir_{Guid.NewGuid():N}");
        Directory.CreateDirectory(baseDir);
        try
        {
            string wd = BasePathFixture.WithBasePath(
                baseDir,
                () => SpreadsheetMcpServer.Core.Helpers.PathResolver.GetWorkingDirectory()
            );
            Assert.Equal(Path.GetFullPath(baseDir), wd);
        }
        finally
        {
            Directory.Delete(baseDir);
        }
    }

    [Fact]
    public void PathResolver_ResolveRelativePath_CombinesWithBase()
    {
        string baseDir = Path.Combine(Path.GetTempPath(), $"base_dir_{Guid.NewGuid():N}");
        Directory.CreateDirectory(baseDir);
        try
        {
            string resolved = BasePathFixture.WithBasePath(
                baseDir,
                () => SpreadsheetMcpServer.Core.Helpers.PathResolver.Resolve("books/x.xlsx")
            );

            Assert.Equal(Path.Combine(Path.GetFullPath(baseDir), "books", "x.xlsx"), resolved);
        }
        finally
        {
            Directory.Delete(baseDir);
        }
    }

    [Fact]
    public void PathResolver_ResolveAbsoluteInsideBase_Accepted()
    {
        string baseDir = Path.Combine(Path.GetTempPath(), $"base_dir_{Guid.NewGuid():N}");
        string inner = Path.Combine(baseDir, "inner.xlsx");
        Directory.CreateDirectory(baseDir);
        try
        {
            string resolved = BasePathFixture.WithBasePath(
                baseDir,
                () => SpreadsheetMcpServer.Core.Helpers.PathResolver.Resolve(inner)
            );

            Assert.Equal(Path.GetFullPath(inner), resolved);
        }
        finally
        {
            Directory.Delete(baseDir);
        }
    }

    [Fact]
    public void PathResolver_ResolveTraversal_EscapesBase_Throws()
    {
        Assert.Throws<InvalidOperationException>(() =>
            SpreadsheetMcpServer.Core.Helpers.PathResolver.Resolve("../../etc/passwd")
        );
    }

    [Fact]
    public void PathResolver_ResolveAbsoluteOutsideBase_Throws()
    {
        // The base is the temp directory (see the class initializer), so the repo's working
        // directory is guaranteed to be outside it.
        Assert.Throws<InvalidOperationException>(() =>
            SpreadsheetMcpServer.Core.Helpers.PathResolver.Resolve(Directory.GetCurrentDirectory())
        );
    }

    [Fact]
    public void ClearRange_NonexistentSheet_Throws()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            // ClosedXML's Worksheet(name) throws ArgumentException, which the service's catch filter
            // lets through unwrapped — the lookup has to go through WorkbookReader.GetWorksheet.
            var ex = Assert.Throws<InvalidOperationException>(() =>
                _cellService.ClearRange(tempFile, "NoSuchSheet", "A1:B2")
            );
            Assert.Contains("NoSuchSheet", ex.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ClearRange_ClearFormats_RemovesCellCarryingOnlyStyling()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.Cell("A1").Value = "hdr";
                // No value — only a fill. This cell is outside the contents-only used range, so it
                // is exactly the cell clearFormats: true exists to remove.
                ws.Cell("C3").Style.Fill.BackgroundColor = XLColor.Red;
                workbook.SaveAs(tempFile);
            }

            string message = _cellService.ClearRange(tempFile, "Sheet1", "A1:E10", clearFormats: true);

            using var readWorkbook = new XLWorkbook(tempFile);
            var readWs = readWorkbook.Worksheet("Sheet1");
            Assert.Contains("C3", message, StringComparison.Ordinal);
            // Compare against an untouched cell rather than XLColor.NoColor — a cleared fill comes
            // back as ClosedXML's default indexed color, not "no color".
            Assert.Equal(readWs.Cell("Z99").Style.Fill.BackgroundColor, readWs.Cell("C3").Style.Fill.BackgroundColor);
            Assert.NotEqual(XLColor.Red, readWs.Cell("C3").Style.Fill.BackgroundColor);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ClearRange_EmptySheet_ReportsNothingToClear()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                workbook.Worksheets.Add("Sheet1");
                workbook.SaveAs(tempFile);
            }

            string message = _cellService.ClearRange(tempFile, "Sheet1", "A1:E10");

            Assert.Contains("No cells to clear", message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ClearRange_ClearFormats_AlsoUnmergesButDefaultKeepsTheMerge()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.Cell("A1").Value = "Merged";
                ws.Range("A1:C1").Merge();
                workbook.SaveAs(tempFile);
            }

            // Contents-only clear leaves the merged region in place...
            _cellService.ClearRange(tempFile, "Sheet1", "A1:C1");
            using (var afterContents = new XLWorkbook(tempFile))
                Assert.Single(afterContents.Worksheet("Sheet1").MergedRanges);

            // ...while XLClearOptions.All drops it along with the styling.
            _cellService.ClearRange(tempFile, "Sheet1", "A1:C1", clearFormats: true);
            using var afterAll = new XLWorkbook(tempFile);
            Assert.Empty(afterAll.Worksheet("Sheet1").MergedRanges);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Theory]
    [InlineData("Infinity")]
    [InlineData("-Infinity")]
    [InlineData("NaN")]
    public void UpdateRange_NonFiniteNumberText_StaysText(string value)
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            // double.TryParse accepts these, but no cell can hold them — XLCellValue's constructor
            // throws ArgumentException, which the catch filter would let escape unwrapped.
            var sheet = new MarkdownSheet { SheetName = "Sheet1", Contents = BuildContents(("A5", value)) };

            _cellService.UpdateRange(tempFile, sheet);

            using var workbook = new XLWorkbook(tempFile);
            var cell = workbook.Worksheet("Sheet1").Cell("A5");
            Assert.Equal(XLDataType.Text, cell.DataType);
            Assert.Equal(value, cell.GetString());
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Theory]
    [InlineData("007")]
    [InlineData("1.5")]
    [InlineData("true")]
    [InlineData("2026-09-07")]
    [InlineData("=SUM(A1:A2)")]
    public void LoadRangeThenUpdateRange_TextThatLooksTyped_SurvivesRoundTrip(string original)
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.Cell("A1").SetValue(original); // explicitly a text cell
                workbook.SaveAs(tempFile);
            }

            // Read the range and write the unmodified result straight back — the primary
            // read-modify-write shape. LoadRange must escape the value so it is not re-typed.
            var read = _cellService.LoadRange(tempFile, "Sheet1", "A1:A1");
            _cellService.UpdateRange(tempFile, new MarkdownSheet { SheetName = "Sheet1", Contents = read.Contents });

            using var after = new XLWorkbook(tempFile);
            var cell = after.Worksheet("Sheet1").Cell("A1");
            Assert.Equal(XLDataType.Text, cell.DataType);
            Assert.Equal(original, cell.GetString());
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void LoadRangeThenUpdateRange_DateCells_SurviveRoundTrip()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"fixture_{Guid.NewGuid():N}.xlsx");
        var date = new DateTime(2026, 9, 7, 0, 0, 0, DateTimeKind.Unspecified);
        var timestamp = new DateTime(2026, 9, 7, 13, 45, 30, DateTimeKind.Unspecified);
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.Cell("A1").Value = date;
                ws.Cell("A2").Value = timestamp;
                workbook.SaveAs(tempFile);
            }

            var read = _cellService.LoadRange(tempFile, "Sheet1", "A1:A2");
            var contents = ParseContents(read.Contents);

            // Rendered in an ISO format SetTypedValue recognises, not the culture-invariant default.
            Assert.Equal("2026-09-07", contents["A1"]);
            Assert.Equal("2026-09-07 13:45:30", contents["A2"]);

            _cellService.UpdateRange(tempFile, new MarkdownSheet { SheetName = "Sheet1", Contents = read.Contents });

            using var after = new XLWorkbook(tempFile);
            var ws2 = after.Worksheet("Sheet1");
            Assert.Equal(XLDataType.DateTime, ws2.Cell("A1").DataType);
            Assert.Equal(date, ws2.Cell("A1").GetDateTime());
            Assert.Equal(XLDataType.DateTime, ws2.Cell("A2").DataType);
            Assert.Equal(timestamp, ws2.Cell("A2").GetDateTime());
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void PathResolver_ResolveBlankPath_Throws(string path)
    {
        // Without the guard these resolve to the base directory itself, pass containment, and then
        // fail deep inside ClosedXML with an opaque message.
        Assert.Throws<ArgumentException>(() => SpreadsheetMcpServer.Core.Helpers.PathResolver.Resolve(path));
    }

    [Fact]
    public void PathResolver_SymlinkInsideBase_IsTrustedAndNotResolved()
    {
        string baseDir = Path.Combine(Path.GetTempPath(), $"base_dir_{Guid.NewGuid():N}");
        string outsideDir = Path.Combine(Path.GetTempPath(), $"outside_dir_{Guid.NewGuid():N}");
        Directory.CreateDirectory(baseDir);
        Directory.CreateDirectory(outsideDir);
        string link = Path.Combine(baseDir, "escape");
        try
        {
            Directory.CreateSymbolicLink(link, outsideDir);

            string resolved = BasePathFixture.WithBasePath(
                baseDir,
                () => SpreadsheetMcpServer.Core.Helpers.PathResolver.Resolve("escape/book.xlsx")
            );

            // Documents the deliberate scope of the guard: the check is lexical, so a symlink placed
            // inside the base directory is followed at open time. The base directory's contents are
            // trusted because the operator chooses what to expose; the resolver guards the *path
            // argument*. If this ever starts throwing, the doc comment on PathResolver must change too.
            Assert.Equal(Path.Combine(Path.GetFullPath(baseDir), "escape", "book.xlsx"), resolved);
        }
        finally
        {
            Directory.Delete(link);
            Directory.Delete(baseDir, recursive: true);
            Directory.Delete(outsideDir, recursive: true);
        }
    }

    [Fact]
    public void ApplyCellFormatting_ColonQualifiedColumnAndRow_SetsWholeColumnAndRow()
    {
        // "A:A"/"1:1" are ClosedXML's own whole-column/whole-row shorthand and must resolve
        // correctly through the unconditional worksheet.Range(address) resolution path.
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            _cellFormatService.ApplyCellFormatting(
                tempFile,
                [
                    new CellFormatSpec { Address = "A:A", Width = 26 },
                    new CellFormatSpec { Address = "1:1", Height = 40 },
                ]
            );

            using var workbook = new XLWorkbook(tempFile);
            var ws = workbook.Worksheet("Sheet1");

            Assert.Equal(26, ws.Column("A").Width, 3);
            Assert.Equal(40, ws.Row(1).Height, 3);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ApplyCellFormatting_BareColumnAddress_ThrowsInsteadOfSilentlyNoOp()
    {
        // The original bug: a bare "A" (no colon) is not a valid ClosedXML address and used to be
        // silently swallowed, applying to nothing. It must now surface as a clear, loud failure.
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            Assert.ThrowsAny<Exception>(() =>
                _cellFormatService.ApplyCellFormatting(tempFile, [new CellFormatSpec { Address = "A", Width = 26 }])
            );
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void UpdateRange_BareColumnAddress_ThrowsClearErrorInsteadOfNullReference()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            var sheet = new MarkdownSheet { SheetName = "Sheet1", Contents = BuildContents(("A", "value")) };

            // A bare column address has no single anchor cell to write a value into — ClosedXML
            // rejects it, and that rejection must surface clearly rather than as an NRE.
            Assert.ThrowsAny<Exception>(() => _cellService.UpdateRange(tempFile, sheet));
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ManageWorksheet_RenameToSameName_IsNoOp()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            string message = _worksheetService.ManageWorksheet(tempFile, "Sheet1", "rename", "Sheet1");

            Assert.Contains("already named", message, StringComparison.Ordinal);

            using var workbook = new XLWorkbook(tempFile);
            Assert.True(workbook.TryGetWorksheet("Sheet1", out _));
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ManageWorksheet_DeleteOnlySheet_Throws()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            var ex = Assert.Throws<InvalidOperationException>(() =>
                _worksheetService.ManageWorksheet(tempFile, "Sheet1", "delete")
            );
            Assert.Contains("only worksheet", ex.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ManageTables_ColumnCountMismatch_Throws()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            var ex = Assert.Throws<ArgumentException>(() =>
                _tableService.ManageTables(
                    tempFile,
                    "Sheet1",
                    "create",
                    table: new SerializableTable
                    {
                        Name = "T1",
                        Reference = "A1:C2",
                        Columns = ["OnlyOneName"],
                    }
                )
            );
            Assert.Contains("3 column", ex.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void PastePictures_EmptyList_ReturnsCleanlyWithoutTouchingFile()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            string message = _pictureService.PastePictures(tempFile, []);
            Assert.Contains("No pictures", message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ManageFreezePanes_Set_FreezesRowsAndColumnsAboveAndLeftOfAnchor()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            _worksheetService.ManageFreezePanes(tempFile, "Sheet1", "set", "B2");

            using var workbook = new XLWorkbook(tempFile);
            var ws = workbook.Worksheet("Sheet1");
            Assert.Equal(1, ws.SheetView.SplitRow);
            Assert.Equal(1, ws.SheetView.SplitColumn);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ManageFreezePanes_Clear_RemovesExistingFreeze()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            _worksheetService.ManageFreezePanes(tempFile, "Sheet1", "set", "B2");
            _worksheetService.ManageFreezePanes(tempFile, "Sheet1", "clear");

            using var workbook = new XLWorkbook(tempFile);
            var ws = workbook.Worksheet("Sheet1");
            Assert.Equal(0, ws.SheetView.SplitRow);
            Assert.Equal(0, ws.SheetView.SplitColumn);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ManageRowsColumns_InsertRows_ShiftsExistingContentDown()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            _worksheetService.ManageRowsColumns(tempFile, "Sheet1", "insertRows", "1", 2);

            using var workbook = new XLWorkbook(tempFile);
            var ws = workbook.Worksheet("Sheet1");
            Assert.Equal("Hello", ws.Cell("A3").GetString());
            Assert.True(ws.Cell("A1").IsEmpty());
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ManageRowsColumns_DeleteColumns_RemovesAndShiftsLeft()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            _worksheetService.ManageRowsColumns(tempFile, "Sheet1", "deleteColumns", "A");

            using var workbook = new XLWorkbook(tempFile);
            var ws = workbook.Worksheet("Sheet1");
            Assert.Equal("World", ws.Cell("A1").GetString());
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ManageMerge_MergeThenUnmerge_RoundTrips()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            _cellService.ManageMerge(tempFile, "Sheet1", "A1:B1", "merge");

            using (var workbook = new XLWorkbook(tempFile))
            {
                var ws = workbook.Worksheet("Sheet1");
                Assert.True(ws.Range("A1:B1").IsMerged());
            }

            _cellService.ManageMerge(tempFile, "Sheet1", "A1:B1", "unmerge");

            using (var workbook = new XLWorkbook(tempFile))
            {
                var ws = workbook.Worksheet("Sheet1");
                Assert.False(ws.Range("A1:B1").IsMerged());
            }
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void AutofitRange_Both_AdjustsColumnWidthAndRowHeight()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            using (var workbook = new XLWorkbook(tempFile))
            {
                var ws = workbook.Worksheet("Sheet1");
                ws.Column("A").Width = 1;
                workbook.Save();
            }

            _cellService.AutofitRange(tempFile, "Sheet1", "A1:C1", "columns");

            using var result = new XLWorkbook(tempFile);
            var resultWs = result.Worksheet("Sheet1");
            Assert.True(resultWs.Column("A").Width > 1);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }
}
