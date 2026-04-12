using ExcelDna.Integration;
using System.Reflection;

namespace RExcelBridge;

public static class VersionFunctions
{
    private const string BUILD_TAG = "DEV-2026-04-10-A";

    [ExcelFunction(Name = "RVersion", Description = "Returns the RExcelBridge build/version string.", Category = "RExcelBridge")]
    public static string RVersion()
    {
        string asm = Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "NA";
        return $"RExcelBridge | {BUILD_TAG} | asm={asm}";
    }
}
