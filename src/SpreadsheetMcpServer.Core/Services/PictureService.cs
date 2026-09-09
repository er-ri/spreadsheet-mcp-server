using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using SpreadsheetMcpServer.Core.Helpers;
using SpreadsheetMcpServer.Core.Models;

namespace SpreadsheetMcpServer.Core.Services;

/// <summary>
/// Service for reading pictures embedded in Excel worksheets.
/// </summary>
public class PictureService : IPictureService
{
    /// <summary>
    /// Returns the first picture whose top-left anchor cell falls within <paramref name="range"/>,
    /// or <c>null</c> when no picture is anchored there. "First" follows the worksheet's picture
    /// enumeration order.
    /// </summary>
    public PictureData? GetFirstPictureInRange(string spreadSheetPath, string spreadSheetName, string range)
    {
        try
        {
            using var workbook = new XLWorkbook(spreadSheetPath);
            var worksheet = WorkbookReader.GetWorksheet(workbook, spreadSheetName);

            var bounds = worksheet.Range(range).RangeAddress;
            int firstRow = bounds.FirstAddress.RowNumber;
            int lastRow = bounds.LastAddress.RowNumber;
            int firstCol = bounds.FirstAddress.ColumnNumber;
            int lastCol = bounds.LastAddress.ColumnNumber;

            foreach (var picture in worksheet.Pictures)
            {
                var anchor = picture.TopLeftCell.Address;
                if (
                    anchor.RowNumber < firstRow
                    || anchor.RowNumber > lastRow
                    || anchor.ColumnNumber < firstCol
                    || anchor.ColumnNumber > lastCol
                )
                    continue;

                using var memory = new MemoryStream();
                picture.ImageStream.Position = 0;
                picture.ImageStream.CopyTo(memory);

                return new PictureData
                {
                    Base64 = Convert.ToBase64String(memory.ToArray()),
                    MimeType = MimeFromFormat(picture.Format),
                    AnchorAddress = anchor.ToStringRelative(false),
                };
            }

            return null;
        }
        catch (Exception ex) when (ex is not InvalidOperationException and not ArgumentException)
        {
            throw new InvalidOperationException($"Error reading Excel file: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Inserts one or more image files into the workbook at <paramref name="spreadSheetPath"/>.
    /// Each spec names its own target worksheet and anchor cell. Optional width/height are applied
    /// in pixels; omitted dimensions keep the original image size.
    /// </summary>
    public string PastePictures(string spreadSheetPath, List<PicturePasteSpec> pictures)
    {
        try
        {
            using var workbook = new XLWorkbook(spreadSheetPath);

            int index = 0;
            foreach (var spec in pictures)
            {
                // Resolve the picture against the shared base directory and reject any path that
                // escapes it, the same containment as spreadsheet paths.
                string picturePath = PathResolver.Resolve(spec.PicturePath);

                if (!File.Exists(picturePath))
                    throw new FileNotFoundException($"Picture file not found: {picturePath}");

                var worksheet = WorkbookReader.GetWorksheet(workbook, spec.TargetSheet);

                // Picture names must be unique per worksheet and at most 31 characters.
                string name = UniquePictureName(worksheet, index);

                using var stream = new FileStream(picturePath, FileMode.Open, FileAccess.Read);
                var picture = worksheet.AddPicture(stream, name).MoveTo(worksheet.Cell(spec.TopLeftAddress));

                // Apply explicit pixel dimensions when provided; omitted ones keep the original size.
                if (spec.Width is int w)
                    picture.Width = w;
                if (spec.Height is int h)
                    picture.Height = h;

                index++;
            }

            workbook.Save();

            return $"Inserted {pictures.Count} picture(s) into {spreadSheetPath}.";
        }
        catch (Exception ex) when (ex is not InvalidOperationException and not ArgumentException)
        {
            throw new InvalidOperationException($"Error writing Excel file: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Builds a picture name that is unique within the worksheet and within ClosedXML's
    /// 31-character limit. Starts from "Pasted_{index}" and bumps the suffix on collision.
    /// </summary>
    private static string UniquePictureName(IXLWorksheet worksheet, int index)
    {
        var existing = new HashSet<string>(worksheet.Pictures.Select(p => p.Name), StringComparer.OrdinalIgnoreCase);

        int suffix = index;
        string name = $"Pasted_{suffix}";
        while (existing.Contains(name))
            name = $"Pasted_{++suffix}";

        return name;
    }

    /// <summary>
    /// Maps a ClosedXML picture format to its image MIME type.
    /// </summary>
    private static string MimeFromFormat(XLPictureFormat format) =>
        format switch
        {
            XLPictureFormat.Png => "image/png",
            XLPictureFormat.Jpeg => "image/jpeg",
            XLPictureFormat.Gif => "image/gif",
            XLPictureFormat.Bmp => "image/bmp",
            XLPictureFormat.Tiff => "image/tiff",
            XLPictureFormat.Emf => "image/emf",
            XLPictureFormat.Wmf => "image/wmf",
            XLPictureFormat.Icon => "image/x-icon",
            _ => "application/octet-stream",
        };
}
