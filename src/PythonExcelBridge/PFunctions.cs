using System;
using ExcelDna.Integration;

namespace PythonExcelBridge;

public static class PFunctions
{
    [ExcelFunction(
        Name = "PPing",
        Description = "Ping the persistent Python worker and return a simple response.",
        Category = "PythonExcelBridge")]
    public static object PPing()
    {
        try
        {
            return PythonBridge.Ping();
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [ExcelFunction(
        Name = "PEval",
        Description = "Evaluate Python code in the persistent Python session.",
        Category = "PythonExcelBridge")]
    public static object PEval(
        [ExcelArgument(Name = "code", Description = "Python code to evaluate.")] string code)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(code))
                return "Error: code is blank.";

            return PythonBridge.Eval(code);
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
        Name = "PPlot",
        Description = "Render a Python matplotlib plot to a stable PNG file based on workbook, sheet, and cell, and return the file path.",
        Category = "PythonExcelBridge")]
    public static object PPlot(
        [ExcelArgument(Name = "code", Description = "Python plotting code or a plotting function call, such as plot(1:10) or my_plot().")] string code,
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

            return PythonBridge.Plot(code, label, w, h);
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [ExcelFunction(
        Name = "PSource",
        Description = "Source a Python script file into the persistent Python session.",
        Category = "PythonExcelBridge")]
    public static object PSource(
        [ExcelArgument(Name = "file", Description = "Path to an Python script file.")] string file)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(file))
                return "Error: file is blank.";

            return PythonBridge.Source(file);
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [ExcelFunction(
        Name = "PCall",
        Description = "Call a Python function with up to 10 arguments.",
        Category = "PythonExcelBridge")]
    public static object PCall(
        [ExcelArgument(Name = "fun", Description = "Name of the Python function.")] string fun,
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

            return PythonBridge.Call(fun, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10);
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [ExcelFunction(
        Name = "PSet",
        Description = "Assign an Excel value or range to an object in the persistent Python session.",
        Category = "PythonExcelBridge")]
    public static object PSet(
        [ExcelArgument(Name = "name", Description = "Name of the Python object.")] string name,
        [ExcelArgument(Name = "value", Description = "Excel value, vector, or range to assign.")] object value)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Error: object name is blank.";

            return PythonBridge.Set(name, value);
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [ExcelFunction(
        Name = "PGet",
        Description = "Return an object from the persistent Python session to Excel.",
        Category = "PythonExcelBridge")]
    public static object PGet(
        [ExcelArgument(Name = "name", Description = "Name of the Python object.")] string name)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Error: object name is blank.";

            return PythonBridge.Get(name);
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [ExcelFunction(
        Name = "PExists",
        Description = "Check whether an object exists in the persistent Python session.",
        Category = "PythonExcelBridge")]
    public static object PExists(
        [ExcelArgument(Name = "name", Description = "Name of the Python object.")] string name)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Error: object name is blank.";

            return PythonBridge.Exists(name);
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [ExcelFunction(
        Name = "PRemove",
        Description = "Remove an object from the persistent Python session.",
        Category = "PythonExcelBridge")]
    public static object PRemove(
        [ExcelArgument(Name = "name", Description = "Name of the Python object.")] string name)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Error: object name is blank.";

            return PythonBridge.Remove(name);
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [ExcelFunction(
        Name = "PObjects",
        Description = "List objects in the persistent Python session with type and dimensions.",
        Category = "PythonExcelBridge")]
    public static object PObjects()
    {
        try
        {
            return PythonBridge.Objects();
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [ExcelFunction(
        Name = "PDescribe",
        Description = "Describe one object in the persistent Python session.",
        Category = "PythonExcelBridge")]
    public static object PDescribe(
        [ExcelArgument(Name = "name", Description = "Name of the Python object.")] string name)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Error: object name is blank.";

            return PythonBridge.Describe(name);
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [ExcelFunction(
    Name = "PPlotPath",
    Description = "Return the resolved Python plot output directory.",
    Category = "PythonExcelBridge")]
    public static object PPlotPath()
    {
        try
        {
            return PythonBridge.GetPlotDirectory();
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

}
