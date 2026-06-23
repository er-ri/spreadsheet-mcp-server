namespace SpreadsheetMcpServer.Core.Models;

public class SerializableSheet
{
    public string Name { get; set; } = string.Empty;
    public string UsedRange { get; set; } = string.Empty;

    /// <summary>
    /// Top-left anchor cell address (e.g. "B2") of every picture on the worksheet.
    /// </summary>
    public List<string> PictureAddresses { get; set; } = new();
}
