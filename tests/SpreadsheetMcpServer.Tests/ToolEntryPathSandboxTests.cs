using System.Reflection;
using ModelContextProtocol.Server;
using SpreadsheetMcpServer;
using SpreadsheetMcpServer.Core.Models;

namespace SpreadsheetMcpServer.Tests;

/// <summary>
/// Guards the path sandbox at the layer that actually applies it.
/// <para>
/// <c>PathResolver.Resolve</c> is called from <see cref="ToolEntry"/>, not from the services, so
/// unit-testing Core alone cannot tell whether a given tool routes through it. A new tool that
/// forgets the call would lose the sandbox silently, with a green build and green tests. These
/// tests enumerate the MCP surface by reflection — the same <see cref="McpServerToolAttribute"/>
/// discovery the server itself uses — so a tool added later is covered without touching this file.
/// </para>
/// </summary>
[Collection(BasePathCollectionDefinition.Name)]
public class ToolEntryPathSandboxTests
{
    /// <summary>A relative path that climbs out of any base directory.</summary>
    private const string EscapingPath = "../../../etc/evil.xlsx";

    /// <summary>Every <c>[McpServerTool]</c> method on <see cref="ToolEntry"/>.</summary>
    private static IEnumerable<MethodInfo> ToolMethods =>
        typeof(ToolEntry)
            .GetMethods(BindingFlags.Public | BindingFlags.Static)
            .Where(m => m.GetCustomAttribute<McpServerToolAttribute>() is not null);

    public static TheoryData<string> ToolsTakingASpreadsheetPath
    {
        get
        {
            var data = new TheoryData<string>();
            foreach (var method in ToolMethods.Where(m => m.GetParameters().Any(p => p.Name == "spreadSheetPath")))
                data.Add(method.Name);
            return data;
        }
    }

    [Fact]
    public void EveryToolTakingAPath_IsCoveredByTheSandboxTheory()
    {
        // A guard on the guard: if the reflection filter ever stops matching (a renamed parameter,
        // a changed attribute), the theory below would silently shrink to zero cases and pass.
        Assert.Equal(16, ToolsTakingASpreadsheetPath.Count);
    }

    [Theory]
    [MemberData(nameof(ToolsTakingASpreadsheetPath))]
    public void Tool_WithPathEscapingTheBase_IsRejectedBeforeReachingTheService(string toolName)
    {
        var method = ToolMethods.Single(m => m.Name == toolName);
        object?[] arguments = method.GetParameters().Select(BuildArgument).ToArray();

        var ex = Assert.Throws<InvalidOperationException>(() =>
        {
            try
            {
                method.Invoke(null, arguments);
            }
            catch (TargetInvocationException tie) when (tie.InnerException is not null)
            {
                throw tie.InnerException;
            }
        });

        // Specifically the sandbox rejection — not an incidental InvalidOperationException from a
        // service that was reached because the path was never resolved.
        Assert.Contains("outside the allowed base directory", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void PastePictures_WithPicturePathEscapingTheBase_IsRejected()
    {
        // PastePictures is the one tool with a second path: the workbook is resolved at the tool
        // boundary, but each spec's PicturePath is resolved inside PictureService. Both must be
        // sandboxed, so this asserts the inner one directly.
        string tempFile = TestFixtureFactory.CreateSimpleWorkbook("Sheet1");
        try
        {
            var pictures = new List<PicturePasteSpec>
            {
                new()
                {
                    PicturePath = EscapingPath,
                    TargetSheet = "Sheet1",
                    TopLeftAddress = "B2",
                },
            };

            var ex = Assert.Throws<InvalidOperationException>(() => ToolEntry.PastePictures(tempFile, pictures));

            Assert.Contains("outside the allowed base directory", ex.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    /// <summary>
    /// Supplies a value for one tool parameter: the escaping path for <c>spreadSheetPath</c>, and
    /// for everything else the cheapest value that lets the call reach the resolver. The arguments
    /// never have to be *valid* — resolution happens first, so nothing downstream is exercised.
    /// </summary>
    private static object? BuildArgument(ParameterInfo parameter)
    {
        if (parameter.Name == "spreadSheetPath")
            return EscapingPath;

        var type = parameter.ParameterType;
        if (type == typeof(string))
            return "A1";
        if (type == typeof(MarkdownSheet))
            return new MarkdownSheet();
        if (Nullable.GetUnderlyingType(type) is not null || !type.IsValueType)
            return type.IsArray || (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(List<>))
                ? Activator.CreateInstance(type)
                : null;

        return Activator.CreateInstance(type);
    }
}
