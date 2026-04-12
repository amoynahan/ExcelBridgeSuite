using System;
using ExcelDna.Integration;

namespace JuliaExcelBridge;

public static class JFunctions
{
    [ExcelFunction(
        Name = "JPing",
        Description = "Ping the persistent Julia worker and return a simple response.",
        Category = "JuliaExcelBridge")]
    public static object RPing()
    {
        try
        {
            return JBridge.Ping();
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [ExcelFunction(
        Name = "JEval",
        Description = "Evaluate Julia code in the persistent Julia session.",
        Category = "JuliaExcelBridge")]
    public static object REval(
        [ExcelArgument(Name = "code", Description = "Julia code to evaluate.")] string code)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(code))
                return "Error: code is blank.";

            return JBridge.Eval(code);
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    private static int ToIntOrDefault(object value, int defaultValue)
    {
        if (value is ExcelMissing || value is ExcelEmpty || value == null)
            return defaultValue;

        try
        {
            return Convert.ToInt32(value);
        }
        catch
        {
            return defaultValue;
        }
    }

    [ExcelFunction(
        Name = "JPlot",
        Description = "Render a Julia plot to a stable PNG file based on workbook, sheet, and cell, and return the file path.",
        Category = "JuliaExcelBridge")]
    public static object RPlot(
        [ExcelArgument(Name = "code", Description = "Julia plotting code or a plotting function call, such as plot(1:10) or my_plot().")] string code,
        [ExcelArgument(Name = "plot_name", Description = "Optional stable plot label used in the file name.")] object plotName,
        [ExcelArgument(Name = "width", Description = "PNG width in pixels.")] object width,
        [ExcelArgument(Name = "height", Description = "PNG height in pixels.")] object height)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(code))
                return "Error: code is blank.";

            string? label = plotName is ExcelMissing or ExcelEmpty ? null : plotName?.ToString();
            int w = ToIntOrDefault(width, 800);
            int h = ToIntOrDefault(height, 600);

            return JBridge.Plot(code, label, w, h);
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [ExcelFunction(
        Name = "JSource",
        Description = "Source an Julia script file into the persistent Julia session.",
        Category = "JuliaExcelBridge")]
    public static object RSource(
        [ExcelArgument(Name = "file", Description = "Path to an Julia script file.")] string file)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(file))
                return "Error: file is blank.";

            return JBridge.Source(file);
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [ExcelFunction(
        Name = "JCall",
        Description = "Call a Julia function with up to 10 arguments.",
        Category = "JuliaExcelBridge")]
    public static object RCall(
        [ExcelArgument(Name = "fun", Description = "Name of the Julia function.")] string fun,
        [ExcelArgument(Name = "arg1", Description = "Argument 1.")] object arg1,
        [ExcelArgument(Name = "arg2", Description = "Argument 2.")] object arg2,
        [ExcelArgument(Name = "arg3", Description = "Argument 3.")] object arg3,
        [ExcelArgument(Name = "arg4", Description = "Argument 4.")] object arg4,
        [ExcelArgument(Name = "arg5", Description = "Argument 5.")] object arg5,
        [ExcelArgument(Name = "arg6", Description = "Argument 6.")] object arg6,
        [ExcelArgument(Name = "arg7", Description = "Argument 7.")] object arg7,
        [ExcelArgument(Name = "arg8", Description = "Argument 8.")] object arg8,
        [ExcelArgument(Name = "arg9", Description = "Argument 9.")] object arg9,
        [ExcelArgument(Name = "arg10", Description = "Argument 10.")] object arg10)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(fun))
                return "Error: function name is blank.";

            return JBridge.Call(fun, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10);
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [ExcelFunction(
        Name = "JSet",
        Description = "Assign an Excel value or range to an object in the persistent Julia session.",
        Category = "JuliaExcelBridge")]
    public static object RSet(
        [ExcelArgument(Name = "name", Description = "Name of the Julia object.")] string name,
        [ExcelArgument(Name = "value", Description = "Excel value, vector, or range to assign.")] object value)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Error: object name is blank.";

            return JBridge.Set(name, value);
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [ExcelFunction(
        Name = "JGet",
        Description = "Return an object from the persistent Julia session to Excel.",
        Category = "JuliaExcelBridge")]
    public static object RGet(
        [ExcelArgument(Name = "name", Description = "Name of the Julia object.")] string name)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Error: object name is blank.";

            return JBridge.Get(name);
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [ExcelFunction(
        Name = "JExists",
        Description = "Check whether an object exists in the persistent Julia session.",
        Category = "JuliaExcelBridge")]
    public static object RExists(
        [ExcelArgument(Name = "name", Description = "Name of the Julia object.")] string name)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Error: object name is blank.";

            return JBridge.Exists(name);
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [ExcelFunction(
        Name = "JRemove",
        Description = "Remove an object from the persistent Julia session.",
        Category = "JuliaExcelBridge")]
    public static object RRemove(
        [ExcelArgument(Name = "name", Description = "Name of the Julia object.")] string name)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Error: object name is blank.";

            return JBridge.Remove(name);
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [ExcelFunction(
        Name = "JObjects",
        Description = "List objects in the persistent Julia session with type and dimensions.",
        Category = "JuliaExcelBridge")]
    public static object RObjects()
    {
        try
        {
            return JBridge.Objects();
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [ExcelFunction(
        Name = "JDescribe",
        Description = "Describe one object in the persistent Julia session.",
        Category = "JuliaExcelBridge")]
    public static object RDescribe(
        [ExcelArgument(Name = "name", Description = "Name of the Julia object.")] string name)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Error: object name is blank.";

            return JBridge.Describe(name);
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [ExcelFunction(
    Name = "JPlotPath",
    Description = "Return the resolved Julia plot output directory.",
    Category = "JuliaExcelBridge")]
    public static object RPlotPath()
    {
        try
        {
            return JBridge.GetPlotDirectory();
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

}
