using System;
using ExcelDna.Integration;

namespace RExcelBridge;

public static class RFunctions
{
    [ExcelFunction(
        Name = "RPing",
        Description = "Ping the persistent R worker and return a simple response.",
        Category = "RExcelBridge")]
    public static object RPing()
    {
        try
        {
            return RBridge.Ping();
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [ExcelFunction(
        Name = "REval",
        Description = "Evaluate R code in the persistent R session.",
        Category = "RExcelBridge")]
    public static object REval(
        [ExcelArgument(Name = "code", Description = "R code to evaluate.")] string code)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(code))
                return "Error: code is blank.";

            return RBridge.Eval(code);
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
        Name = "RPlot",
        Description = "Render an R plot to a stable PNG file based on workbook, sheet, and cell, and return the file path.",
        Category = "RExcelBridge")]
    public static object RPlot(
        [ExcelArgument(Name = "code", Description = "R plotting code or a plotting function call, such as plot(1:10) or my_plot().")] string code,
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

            return RBridge.Plot(code, label, w, h);
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [ExcelFunction(
        Name = "RSource",
        Description = "Source an R script file into the persistent R session.",
        Category = "RExcelBridge")]
    public static object RSource(
        [ExcelArgument(Name = "file", Description = "Path to an R script file.")] string file)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(file))
                return "Error: file is blank.";

            return RBridge.Source(file);
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [ExcelFunction(
        Name = "RCall",
        Description = "Call an R function with up to 10 arguments.",
        Category = "RExcelBridge")]
    public static object RCall(
        [ExcelArgument(Name = "fun", Description = "Name of the R function.")] string fun,
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

            return RBridge.Call(fun, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10);
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [ExcelFunction(
        Name = "RSet",
        Description = "Assign an Excel value or range to an object in the persistent R session.",
        Category = "RExcelBridge")]
    public static object RSet(
        [ExcelArgument(Name = "name", Description = "Name of the R object.")] string name,
        [ExcelArgument(Name = "value", Description = "Excel value, vector, or range to assign.")] object value)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Error: object name is blank.";

            return RBridge.Set(name, value);
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [ExcelFunction(
        Name = "RGet",
        Description = "Return an object from the persistent R session to Excel.",
        Category = "RExcelBridge")]
    public static object RGet(
        [ExcelArgument(Name = "name", Description = "Name of the R object.")] string name)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Error: object name is blank.";

            return RBridge.Get(name);
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [ExcelFunction(
        Name = "RExists",
        Description = "Check whether an object exists in the persistent R session.",
        Category = "RExcelBridge")]
    public static object RExists(
        [ExcelArgument(Name = "name", Description = "Name of the R object.")] string name)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Error: object name is blank.";

            return RBridge.Exists(name);
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [ExcelFunction(
        Name = "RRemove",
        Description = "Remove an object from the persistent R session.",
        Category = "RExcelBridge")]
    public static object RRemove(
        [ExcelArgument(Name = "name", Description = "Name of the R object.")] string name)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Error: object name is blank.";

            return RBridge.Remove(name);
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [ExcelFunction(
        Name = "RObjects",
        Description = "List objects in the persistent R session with type and dimensions.",
        Category = "RExcelBridge")]
    public static object RObjects()
    {
        try
        {
            return RBridge.Objects();
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [ExcelFunction(
        Name = "RDescribe",
        Description = "Describe one object in the persistent R session.",
        Category = "RExcelBridge")]
    public static object RDescribe(
        [ExcelArgument(Name = "name", Description = "Name of the R object.")] string name)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Error: object name is blank.";

            return RBridge.Describe(name);
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [ExcelFunction(
    Name = "RPlotPath",
    Description = "Return the resolved plot output directory.",
    Category = "RExcelBridge")]
    public static object RPlotPath()
    {
        try
        {
            return RBridge.GetPlotDirectory();
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

}
