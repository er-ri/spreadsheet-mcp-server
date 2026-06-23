using ClosedXML.Excel;
using SpreadsheetMcpServer.Core.Models;

namespace SpreadsheetMcpServer.Core.Services;

/// <summary>
/// Service for managing Excel worksheets.
/// </summary>
public class WorksheetService : IWorksheetService
{
    /// <summary>
    /// Gets a list of all worksheets in the spreadsheet, each with its name and used range.
    /// </summary>
    public List<SerializableSheet> GetAllWorksheets(string spreadSheetPath)
    {
        try
        {
            using var workbook = new XLWorkbook(spreadSheetPath);
            return workbook
                .Worksheets.Select(ws => new SerializableSheet
                {
                    Name = ws.Name,
                    UsedRange = ws.RangeUsed()?.RangeAddress.ToStringRelative(false) ?? string.Empty,
                    PictureAddresses = ws
                        .Pictures.Select(picture => picture.TopLeftCell.Address.ToStringRelative(false))
                        .ToList(),
                })
                .ToList();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Error reading Excel file: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Creates, deletes, or renames a worksheet. 'action' must be "create", "delete", or
    /// "rename"; 'newSheetName' is required only when renaming. Returns a confirmation message.
    /// </summary>
    public string ManageWorksheet(
        string spreadSheetPath,
        string spreadSheetName,
        string action,
        string? newSheetName = null
    )
    {
        try
        {
            using var workbook = new XLWorkbook(spreadSheetPath);
            string message;
            switch (action.Trim().ToLowerInvariant())
            {
                case "create":
                    if (workbook.TryGetWorksheet(spreadSheetName, out _))
                        throw new InvalidOperationException($"A worksheet named '{spreadSheetName}' already exists.");
                    workbook.Worksheets.Add(spreadSheetName);
                    message = $"Created worksheet '{spreadSheetName}'.";
                    break;

                case "delete":
                    if (!workbook.TryGetWorksheet(spreadSheetName, out var sheetToDelete))
                        throw new InvalidOperationException($"No worksheet named '{spreadSheetName}' was found.");
                    sheetToDelete.Delete();
                    message = $"Deleted worksheet '{spreadSheetName}'.";
                    break;

                case "rename":
                    ArgumentException.ThrowIfNullOrWhiteSpace(newSheetName);
                    if (!workbook.TryGetWorksheet(spreadSheetName, out var sheetToRename))
                        throw new InvalidOperationException($"No worksheet named '{spreadSheetName}' was found.");
                    if (workbook.TryGetWorksheet(newSheetName, out _))
                        throw new InvalidOperationException($"A worksheet named '{newSheetName}' already exists.");
                    sheetToRename.Name = newSheetName;
                    message = $"Renamed worksheet '{spreadSheetName}' to '{newSheetName}'.";
                    break;

                default:
                    throw new InvalidOperationException(
                        $"Unknown action '{action}'. Expected 'create', 'delete', or 'rename'."
                    );
            }

            workbook.Save();
            return message;
        }
        catch (Exception ex) when (ex is not InvalidOperationException and not ArgumentException)
        {
            throw new InvalidOperationException($"Error updating Excel file: {ex.Message}", ex);
        }
    }
}
