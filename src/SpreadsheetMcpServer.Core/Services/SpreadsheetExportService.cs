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
        var token = JToken.Parse(json);
        var array = token is JArray arr ? arr : new JArray(token);

        // MiniExcel accepts IEnumerable<IDictionary<string,object>>
        var rows = array
            .OfType<JObject>()
            .Select(o => o.Properties().ToDictionary(p => p.Name, p => (object?)p.Value.ToObject<object>()))
            .ToList();

        var tempSpreadSheetPath = Path.Combine(Path.GetTempPath(), $"temp_{Guid.NewGuid():N}.xlsx");
        var tempSheetName = "tempSheet";

        MiniExcel.SaveAs(tempSpreadSheetPath, rows, sheetName: tempSheetName);

        // Open the temp workbook written by MiniExcel and grab the sheet
        using var tempWorkbook = new XLWorkbook(tempSpreadSheetPath);
        var tempSheet = tempWorkbook.Worksheet(tempSheetName);

        // Open the target workbook if it exists; otherwise start a new one
        XLWorkbook targetWorkbook = File.Exists(spreadSheetPath) ? new XLWorkbook(spreadSheetPath) : new XLWorkbook();

        using (targetWorkbook)
        {
            // Throw if a sheet with the same name already exists
            if (targetWorkbook.TryGetWorksheet(sheetName, out _))
                throw new InvalidOperationException($"Sheet '{sheetName}' already exists in '{spreadSheetPath}'.");

            tempSheet.CopyTo(targetWorkbook, sheetName);
            targetWorkbook.SaveAs(spreadSheetPath);
        }

        File.Delete(tempSpreadSheetPath);
    }
}
