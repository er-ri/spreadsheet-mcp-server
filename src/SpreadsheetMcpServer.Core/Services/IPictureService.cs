using SpreadsheetMcpServer.Core.Models;

namespace SpreadsheetMcpServer.Core.Services;

/// <summary>
/// Service for reading pictures embedded in Excel worksheets.
/// </summary>
public interface IPictureService
{
    /// <summary>
    /// Returns the first picture whose top-left anchor cell falls within <paramref name="range"/>,
    /// or <c>null</c> when the worksheet has no such picture. "First" follows the worksheet's picture
    /// enumeration order.
    /// </summary>
    PictureData? GetFirstPictureInRange(string spreadSheetPath, string spreadSheetName, string range);

    /// <summary>
    /// Inserts one or more image files into the workbook at <paramref name="spreadSheetPath"/>.
    /// Each spec names its own target worksheet and anchor cell. The inverse of
    /// <see cref="GetFirstPictureInRange"/>. Returns a confirmation message.
    /// </summary>
    string PastePictures(string spreadSheetPath, List<PicturePasteSpec> pictures);
}
