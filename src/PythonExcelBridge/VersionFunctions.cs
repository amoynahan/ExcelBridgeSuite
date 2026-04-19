using ExcelDna.Integration;
using System.Reflection;

namespace PythonExcelBridge;

public static class VersionFunctions
{
    private const string BUILD_TAG = "DEV-2026-04-19-PYSYNC";

    [ExcelFunction(Name = "PVersion", Description = "Returns the PythonExcelBridge build/version string.", Category = "PythonExcelBridge")]
    public static string PVersion()
    {
        string asm = Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "NA";
        return $"PythonExcelBridge | {BUILD_TAG} | asm={asm}";
    }
}
