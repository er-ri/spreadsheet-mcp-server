using ClosedXML.Excel;
using SpreadsheetMcpServer.Core.Helpers;
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
        catch (Exception ex) when (ex is not InvalidOperationException and not ArgumentException)
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
            // Create new workbook if file doesn't exist, otherwise open existing (through the
            // phonetic-run sanitizer, like every other path that reads existing content).
            var fileExists = File.Exists(spreadSheetPath);
            using var sanitized = fileExists ? WorkbookReader.SanitizePhoneticRuns(spreadSheetPath) : null;
            using var workbook = fileExists ? new XLWorkbook(sanitized!) : new XLWorkbook();
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
                    if (workbook.Worksheets.Count <= 1)
                        throw new InvalidOperationException("Cannot delete the only worksheet in the workbook.");
                    sheetToDelete.Delete();
                    message = $"Deleted worksheet '{spreadSheetName}'.";
                    break;

                case "rename":
                    ArgumentException.ThrowIfNullOrWhiteSpace(newSheetName);
                    if (!workbook.TryGetWorksheet(spreadSheetName, out var sheetToRename))
                        throw new InvalidOperationException($"No worksheet named '{spreadSheetName}' was found.");
                    if (string.Equals(newSheetName, spreadSheetName, StringComparison.Ordinal))
                    {
                        message = $"Worksheet '{spreadSheetName}' is already named '{newSheetName}'.";
                        break;
                    }
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

            workbook.SaveAs(spreadSheetPath);
            return message;
        }
        catch (Exception ex) when (ex is not InvalidOperationException and not ArgumentException)
        {
            throw new InvalidOperationException($"Error updating Excel file: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Sets or clears frozen panes on a worksheet. 'action' must be "set" or "clear". For "set",
    /// 'freezeAtAddress' is required and freezes every row above and column left of that cell.
    /// For "clear", any existing freeze is removed. Returns a confirmation message.
    /// </summary>
    public string ManageFreezePanes(
        string spreadSheetPath,
        string spreadSheetName,
        string action,
        string? freezeAtAddress = null
    )
    {
        try
        {
            using var sanitized = WorkbookReader.SanitizePhoneticRuns(spreadSheetPath);
            using var workbook = new XLWorkbook(sanitized);
            var worksheet = WorkbookReader.GetWorksheet(workbook, spreadSheetName);
            string message;

            switch (action.Trim().ToLowerInvariant())
            {
                case "set":
                    ArgumentException.ThrowIfNullOrWhiteSpace(freezeAtAddress);
                    var anchor = worksheet.Cell(freezeAtAddress).Address;
                    worksheet.SheetView.Freeze(anchor.RowNumber - 1, anchor.ColumnNumber - 1);
                    message = $"Froze panes on '{spreadSheetName}' at '{freezeAtAddress}'.";
                    break;

                case "clear":
                    worksheet.SheetView.Freeze(0, 0);
                    message = $"Cleared frozen panes on '{spreadSheetName}'.";
                    break;

                default:
                    throw new InvalidOperationException($"Unknown action '{action}'. Expected 'set' or 'clear'.");
            }

            workbook.SaveAs(spreadSheetPath);
            return message;
        }
        catch (Exception ex) when (ex is not InvalidOperationException and not ArgumentException)
        {
            throw new InvalidOperationException($"Error updating Excel file: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Inserts or deletes whole rows or columns on a worksheet, shifting subsequent cells. 'action'
    /// must be "insertRows", "deleteRows", "insertColumns", or "deleteColumns". 'target' is a
    /// 1-based row number (Rows actions) or a column letter (Columns actions). 'count' (default 1)
    /// is how many rows/columns to insert or delete. Returns a confirmation message.
    /// </summary>
    public string ManageRowsColumns(
        string spreadSheetPath,
        string spreadSheetName,
        string action,
        string target,
        int count = 1
    )
    {
        try
        {
            using var sanitized = WorkbookReader.SanitizePhoneticRuns(spreadSheetPath);
            using var workbook = new XLWorkbook(sanitized);
            var worksheet = WorkbookReader.GetWorksheet(workbook, spreadSheetName);
            string message;

            if (count <= 0)
                throw new ArgumentException($"'count' must be at least 1 (got {count}).");

            switch (action.Trim().ToLowerInvariant())
            {
                case "insertrows":
                    worksheet.Row(ParseRowTarget(target)).InsertRowsAbove(count);
                    message = $"Inserted {count} row(s) above row {target} on '{spreadSheetName}'.";
                    break;

                case "deleterows":
                    worksheet.Rows(ParseRowTarget(target), ParseRowTarget(target) + count - 1).Delete();
                    message = $"Deleted {count} row(s) starting at row {target} on '{spreadSheetName}'.";
                    break;

                case "insertcolumns":
                    worksheet.Column(target).InsertColumnsBefore(count);
                    message = $"Inserted {count} column(s) before column {target} on '{spreadSheetName}'.";
                    break;

                case "deletecolumns":
                    int startCol = CellReferenceParser.ColumnLetterToIndex(target);
                    worksheet.Columns(startCol, startCol + count - 1).Delete();
                    message = $"Deleted {count} column(s) starting at column {target} on '{spreadSheetName}'.";
                    break;

                default:
                    throw new InvalidOperationException(
                        $"Unknown action '{action}'. Expected 'insertRows', 'deleteRows', 'insertColumns', or 'deleteColumns'."
                    );
            }

            workbook.SaveAs(spreadSheetPath);
            return message;
        }
        catch (Exception ex) when (ex is not InvalidOperationException and not ArgumentException)
        {
            throw new InvalidOperationException($"Error updating Excel file: {ex.Message}", ex);
        }
    }

    private static int ParseRowTarget(string target) =>
        int.TryParse(target, out var row) && row >= 1
            ? row
            : throw new ArgumentException($"'target' must be a positive row number for this action (got '{target}').");
}
