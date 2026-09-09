namespace SpreadsheetMcpServer.Core.Helpers;

/// <summary>
/// Resolves file paths passed to MCP tools against the shared base directory
/// (<c>SPREADSHEET_BASE_PATH</c>, falling back to the current directory) and rejects any path that
/// would escape it: a relative path containing <c>../</c> that climbs past the base, or an absolute
/// path pointing outside it. Spreadsheet and image files must therefore live under the base
/// directory.
/// <para>
/// The check is lexical — it normalizes the path but does not follow symlinks — so the contents of
/// the base directory are trusted: a symlink placed inside it can still point outside, and the file
/// it targets will be read or written. That matches the deployment model, where the operator chooses
/// which host directory to expose (the Docker image mounts one at <c>/data</c>). This resolver is a
/// guard against a mistaken or malicious *path argument*, not a sandbox against the directory's own
/// contents.
/// </para>
/// </summary>
public static class PathResolver
{
    /// <summary>
    /// Returns the directory all spreadsheet/image paths are resolved against: the
    /// <c>SPREADSHEET_BASE_PATH</c> environment variable when set, otherwise the current directory.
    /// </summary>
    public static string GetWorkingDirectory() =>
        Environment.GetEnvironmentVariable("SPREADSHEET_BASE_PATH") ?? Directory.GetCurrentDirectory();

    /// <summary>
    /// Resolves <paramref name="path"/> to an absolute path inside
    /// <see cref="GetWorkingDirectory"/>. Relative paths are combined with the base directory;
    /// absolute paths must already lie under it. Throws <see cref="InvalidOperationException"/> when
    /// the resolved path falls outside the base directory.
    /// </summary>
    public static string Resolve(string path)
    {
        // An empty or blank path resolves to the base directory itself, which passes containment and
        // then fails deep inside ClosedXML with an opaque message. Reject it at the boundary.
        ArgumentException.ThrowIfNullOrWhiteSpace(path);

        string basePath = Path.GetFullPath(GetWorkingDirectory());

        string resolved = Path.IsPathRooted(path)
            ? Path.GetFullPath(path)
            : Path.GetFullPath(Path.Combine(basePath, path));

        string relative = Path.GetRelativePath(basePath, resolved);
        bool escapesBase =
            relative.StartsWith($"..{Path.DirectorySeparatorChar}", StringComparison.Ordinal) || relative == "..";
        if (escapesBase)
            throw new InvalidOperationException(
                $"Path '{path}' resolves outside the allowed base directory '{basePath}'. "
                    + "Spreadsheet files must live within SPREADSHEET_BASE_PATH."
            );

        return resolved;
    }
}
