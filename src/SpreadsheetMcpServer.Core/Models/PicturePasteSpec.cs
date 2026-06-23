namespace SpreadsheetMcpServer.Core.Models;

/// <summary>
/// Describes a single picture to insert into a worksheet via PastePictures.
/// </summary>
public class PicturePasteSpec
{
    /// <summary>Source image file path, resolved relative to GetWorkingDirectory (SPREADSHEET_BASE_PATH) when not rooted.</summary>
    public string PicturePath { get; set; } = string.Empty;

    /// <summary>Target worksheet name within the workbook.</summary>
    public string TargetSheet { get; set; } = string.Empty;

    /// <summary>Top-left anchor cell address, e.g. "B2".</summary>
    public string TopLeftAddress { get; set; } = string.Empty;

    /// <summary>Optional width in pixels; the original image width is used when null.</summary>
    public int? Width { get; set; }

    /// <summary>Optional height in pixels; the original image height is used when null.</summary>
    public int? Height { get; set; }
}
