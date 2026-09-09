using ClosedXML.Excel;
using MiniExcelLibs;
using Newtonsoft.Json.Linq;

namespace SpreadsheetMcpServer.Core.Services;

/// <summary>
/// Service for exporting data to Excel files.
/// </summary>
public class SpreadsheetExportService : ISpreadsheetExportService
{
    /// <summary>
    /// Exports a JSON string to an Excel file.
    /// </summary>
    public void ExportJsonArrayToSpreadSheet(string spreadSheetPath, string sheetName, string json)
    {
        try
        {
            var token = JToken.Parse(json);
            var array = token is JArray arr ? arr : new JArray(token);

            // MiniExcel accepts IEnumerable<IDictionary<string,object>>
            var rows = array
                .OfType<JObject>()
                .Select(o => o.Properties().ToDictionary(p => p.Name, p => (object?)p.Value.ToObject<object>()))
                .ToList();

            // Render the JSON rows through MiniExcel into an in-memory workbook, then copy its sheet into
            // the target. No temp file is used, so nothing can leak if a later step throws.
            var tempSheetName = "tempSheet";
            using var memory = new MemoryStream();
            MiniExcel.SaveAs(memory, rows, sheetName: tempSheetName);
            memory.Position = 0;

            using var tempWorkbook = new XLWorkbook(memory);
            var tempSheet = tempWorkbook.Worksheet(tempSheetName);

            // Open the target workbook if it exists; otherwise start a new one
            using var targetWorkbook = File.Exists(spreadSheetPath)
                ? new XLWorkbook(spreadSheetPath)
                : new XLWorkbook();

            // Throw if a sheet with the same name already exists
            if (targetWorkbook.TryGetWorksheet(sheetName, out _))
                throw new InvalidOperationException($"Sheet '{sheetName}' already exists in '{spreadSheetPath}'.");

            tempSheet.CopyTo(targetWorkbook, sheetName);
            targetWorkbook.SaveAs(spreadSheetPath);
        }
        catch (Exception ex) when (ex is not InvalidOperationException and not ArgumentException)
        {
            throw new InvalidOperationException($"Error exporting to Excel file: {ex.Message}", ex);
        }
    }
}
