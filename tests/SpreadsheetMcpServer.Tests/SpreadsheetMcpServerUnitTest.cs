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

    [Fact]
    public void GetAllWorksheets_WithValidFile_ReturnsSheetNames()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            List<string> result = _worksheetService.GetAllWorksheets(tempFile);

            Assert.Contains("Sheet1", result);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ReadSpreadSheet_WithValidFile_ReturnsCells()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            List<SerializableCell> result = _cellService.ReadSpreadSheet(tempFile, "Sheet1");

            Assert.True(result.Count > 0);
            Assert.Contains(result, c => c.Value == "Hello");
            Assert.Contains(result, c => c.Value == "42");
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void CreateTable_CreatesTableWithExpectedProperties()
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

            _tableService.CreateTable(tempFile, table, "Sheet1");

            List<SerializableTable> result = _tableService.GetTables(tempFile, "Sheet1");

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
    public void DeleteTable_RemovesTable()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            _tableService.CreateTable(
                tempFile,
                new SerializableTable
                {
                    Name = "ToDelete",
                    Reference = "A1:C1",
                    Columns = ["Hello", "World", "Value"],
                },
                "Sheet1"
            );

            _tableService.DeleteTable(tempFile, "Sheet1", "ToDelete");

            List<SerializableTable> result = _tableService.GetTables(tempFile, "Sheet1");
            Assert.Empty(result);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void GetTables_ReturnsEmptyWhenNoTables()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            List<SerializableTable> result = _tableService.GetTables(tempFile, "Sheet1");
            Assert.Empty(result);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void UpdateCells_WritesValuesAndRoundTripsCorrectly()
    {
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            const string sheetName = "Sheet1";

            var updates = new List<SerializableCell>
            {
                new() { Address = $"{sheetName}!A1", Value = "UpdatedString" },
                new()
                {
                    Address = $"{sheetName}!B1",
                    Value = "42",
                    DataType = "Number",
                },
                new() { Address = $"{sheetName}!C1", Value = "NewCell" },
            };

            _cellService.UpdateCells(tempFile, updates);

            List<SerializableCell> result = _cellService.ReadSpreadSheet(tempFile, sheetName);

            SerializableCell? a1 = result.FirstOrDefault(c => c.Address == $"{sheetName}!A1");
            SerializableCell? b1 = result.FirstOrDefault(c => c.Address == $"{sheetName}!B1");
            SerializableCell? c1 = result.FirstOrDefault(c => c.Address == $"{sheetName}!C1");

            Assert.NotNull(a1);
            Assert.Equal("UpdatedString", a1.Value);

            Assert.NotNull(b1);
            Assert.Equal("42", b1.Value);

            Assert.NotNull(c1);
            Assert.Equal("NewCell", c1.Value);
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
    public void CreateTable_CreatesTableWithAutoFilterAndStyle()
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

            _tableService.CreateTable(tempFile, table, "Sheet1");

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
}
