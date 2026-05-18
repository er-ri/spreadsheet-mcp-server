using ClosedXML.Excel;

namespace SpreadsheetMcpServer.Core.Services;

/// <summary>
/// Service for managing Excel worksheets.
/// </summary>
public class WorksheetService : IWorksheetService
{
    /// <summary>
    /// Gets a list of all worksheet names in the spreadsheet.
    /// </summary>
    public List<string> GetAllWorksheets(string spreadSheetPath)
    {
        try
        {
            using var workbook = new XLWorkbook(spreadSheetPath);
            return workbook.Worksheets.Select(ws => ws.Name).ToList();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Error reading Excel file: {ex.Message}", ex);
        }
    }
}
