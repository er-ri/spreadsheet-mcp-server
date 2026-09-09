namespace SpreadsheetMcpServer.Tests;

/// <summary>
/// Owns the process-wide <c>SPREADSHEET_BASE_PATH</c> environment variable that
/// <see cref="Core.Helpers.PathResolver"/> reads.
/// <para>
/// Fixtures live in <see cref="Path.GetTempPath"/> (absolute paths), so the sandbox base has to
/// cover the temp directory for any test that goes through a service. That is process-wide state,
/// which means no two tests may hold different values at the same time. Every test class that
/// touches a spreadsheet path therefore joins <see cref="BasePathCollectionDefinition"/>, so xUnit runs them
/// in one collection — sequentially — rather than that being an accident of them all living in a
/// single class.
/// </para>
/// </summary>
public sealed class BasePathFixture : IDisposable
{
    private readonly string? _previous;

    public BasePathFixture()
    {
        _previous = Environment.GetEnvironmentVariable("SPREADSHEET_BASE_PATH");
        Environment.SetEnvironmentVariable("SPREADSHEET_BASE_PATH", Path.GetTempPath());
    }

    /// <summary>
    /// Runs <paramref name="action"/> with the base directory temporarily pointed at
    /// <paramref name="basePath"/>, restoring the collection-wide value afterwards. Safe only
    /// because the collection serializes the tests that use it.
    /// </summary>
    public static string WithBasePath(string? basePath, Func<string> action)
    {
        string? previous = Environment.GetEnvironmentVariable("SPREADSHEET_BASE_PATH");
        Environment.SetEnvironmentVariable("SPREADSHEET_BASE_PATH", basePath);
        try
        {
            return action();
        }
        finally
        {
            Environment.SetEnvironmentVariable("SPREADSHEET_BASE_PATH", previous);
        }
    }

    public void Dispose() => Environment.SetEnvironmentVariable("SPREADSHEET_BASE_PATH", _previous);
}

/// <summary>
/// Defines the collection every test class touching a spreadsheet path must join, so that the shared
/// <see cref="BasePathFixture"/> is created once and those classes never run in parallel.
/// </summary>
[CollectionDefinition(Name)]
public sealed class BasePathCollectionDefinition : ICollectionFixture<BasePathFixture>
{
    public const string Name = "SpreadsheetBasePath";
}
