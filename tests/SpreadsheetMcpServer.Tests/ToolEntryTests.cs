using ModelContextProtocol.Protocol;
using SpreadsheetMcpServer;

namespace SpreadsheetMcpServer.Tests;

/// <summary>
/// Exercises <see cref="ToolEntry"/> methods directly, rather than the <c>Core</c> services they
/// delegate to. <see cref="ToolEntryPathSandboxTests"/> covers the sandbox behavior shared by every
/// tool; this file covers the handful of tools with their own Host-layer logic worth asserting on
/// directly — <c>GetWorkingDirectory</c> (a thin static call with no Core-level equivalent) and
/// <c>GetPictures</c> (the response-shape conversion from <c>PictureData</c> to
/// <see cref="ImageContentBlock"/>).
/// </summary>
[Collection(BasePathCollectionDefinition.Name)]
public class ToolEntryTests
{
    [Fact]
    public void GetWorkingDirectory_ReturnsResolvedBasePath()
    {
        string baseDir = Path.Combine(Path.GetTempPath(), $"base_dir_{Guid.NewGuid():N}");
        Directory.CreateDirectory(baseDir);
        try
        {
            string wd = BasePathFixture.WithBasePath(baseDir, ToolEntry.GetWorkingDirectory);
            Assert.Equal(Path.GetFullPath(baseDir), wd);
        }
        finally
        {
            Directory.Delete(baseDir);
        }
    }

    [Fact]
    public void GetPictures_RangeContainingPicture_ReturnsImageContentBlock()
    {
        string tempFile = TestFixtureFactory.CreateWorkbookWithPicture("Sheet1", "B2");
        try
        {
            var result = ToolEntry.GetPictures(tempFile, "Sheet1", "A1:C3");

            Assert.NotNull(result);
            Assert.Equal("image/png", result.MimeType);
            Assert.Equal(TestFixtureFactory.OnePixelPng, result.Data.ToArray());
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void GetPictures_RangeWithoutPicture_ReturnsNull()
    {
        string tempFile = TestFixtureFactory.CreateWorkbookWithPicture("Sheet1", "B2");
        try
        {
            var result = ToolEntry.GetPictures(tempFile, "Sheet1", "D4:E5");

            Assert.Null(result);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }
}
