using ExcelDna.Integration;
using System.Reflection;

namespace PythonExcelBridge;

public static class VersionFunctions
{
    private const string BUILD_TAG = "DEV-2026-04-12-PY";

    [ExcelFunction(Name = "PVersion", Description = "Returns the PythonExcelBridge build/version string.", Category = "PythonExcelBridge")]
    public static string JVersion()
    {
        string asm = Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "NA";
        return $"PythonExcelBridge | {BUILD_TAG} | asm={asm}";
    }
}
