using ClosedXML.Excel;
using SpreadsheetMcpServer.Core.Models;
using SpreadsheetMcpServer.Core.Services;

namespace SpreadsheetMcpServer.Tests;

public class SpreadsheetMcpServerUnitTest
{
    private readonly IWorksheetService _worksheetService = new WorksheetService();
    private readonly ICellService _cellService = new CellService();
    private readonly ITableService _tableService = new TableService();
    private readonly ISpreadsheetExportService _exportService = new SpreadsheetExportService();
    private readonly IPictureService _pictureService = new PictureService();

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
            sb.Append($"\n| {address} | {markdown.Replace("|", "\\|")} |");
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
    public void ExportJsonArray_CreatesNewFileWithHeadersAndRows()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"export_{Guid.NewGuid():N}.xlsx");
        try
        {
            const string json = """[{"Name":"Alice","Age":30},{"Name":"Bob","Age":25}]""";

            _exportService.ExportJsonArrayToSpreadSheet(tempFile, "People", json);

            using var workbook = new XLWorkbook(tempFile);
            var sheet = workbook.Worksheet("People");

            Assert.Equal("Name", sheet.Cell("A1").GetString());
            Assert.Equal("Age", sheet.Cell("B1").GetString());
            Assert.Equal("Alice", sheet.Cell("A2").GetString());
            Assert.Equal("Bob", sheet.Cell("A3").GetString());
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ExportJsonSingleObject_CreatesSheetWithOneRow()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"export_{Guid.NewGuid():N}.xlsx");
        try
        {
            const string json = """{"Name":"Alice","Age":30}""";

            _exportService.ExportJsonArrayToSpreadSheet(tempFile, "People", json);

            using var workbook = new XLWorkbook(tempFile);
            var sheet = workbook.Worksheet("People");

            Assert.Equal("Name", sheet.Cell("A1").GetString());
            Assert.Equal("Alice", sheet.Cell("A2").GetString());
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ExportJsonArray_AppendsSheetToExistingWorkbook()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Existing");
        try
        {
            const string json = """[{"City":"Paris","Pop":2000000}]""";

            _exportService.ExportJsonArrayToSpreadSheet(tempFile, "Cities", json);

            using var workbook = new XLWorkbook(tempFile);
            Assert.True(workbook.TryGetWorksheet("Existing", out _));
            Assert.True(workbook.TryGetWorksheet("Cities", out _));
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ExportJsonArray_ThrowsWhenSheetAlreadyExists()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            const string json = """[{"Name":"Alice"}]""";

            Assert.Throws<InvalidOperationException>(() =>
                _exportService.ExportJsonArrayToSpreadSheet(tempFile, "Sheet1", json)
            );
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
            int dataRows = result.Contents.Split('\n').Count(l => l.TrimStart().StartsWith("| ")) - 2; // header row + separator row
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
}
