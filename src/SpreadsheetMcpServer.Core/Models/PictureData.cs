namespace SpreadsheetMcpServer.Core.Models;

/// <summary>
/// A format-neutral representation of a single picture read from a worksheet. The host layer maps
/// this to its own image content type; keeping it MCP-SDK-free lets the Core library stay free of
/// any transport dependency.
/// </summary>
public class PictureData
{
    /// <summary>The picture bytes, Base64-encoded.</summary>
    public string Base64 { get; set; } = string.Empty;

    /// <summary>The image MIME type, e.g. "image/png".</summary>
    public string MimeType { get; set; } = string.Empty;

    /// <summary>The top-left anchor cell address, e.g. "B2".</summary>
    public string AnchorAddress { get; set; } = string.Empty;
}
