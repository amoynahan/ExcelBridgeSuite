using ExcelDna.Integration;
using System.Reflection;

namespace JuliaExcelBridge;

public static class VersionFunctions
{
    private const string BUILD_TAG = "DEV-2026-04-12-JA";

    [ExcelFunction(Name = "JVersion", Description = "Returns the JuliaExcelBridge build/version string.", Category = "JuliaExcelBridge")]
    public static string JVersion()
    {
        string asm = Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "NA";
        return $"JuliaExcelBridge | {BUILD_TAG} | asm={asm}";
    }
}
