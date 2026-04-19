using System;
using System.Collections.Generic;
using System.Linq;
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

    private static double NormalizeTrigger(object triggerRange)
    {
        if (triggerRange is object[,] arr)
        {
            double sum = 0.0;

            int rows = arr.GetLength(0);
            int cols = arr.GetLength(1);

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    if (arr[i, j] is double d)
                        sum += d;
                }
            }

            return sum;
        }

        if (triggerRange is double d2)
            return d2;

        return 0.0;
    }

    [ExcelFunction(
        Name = "RPlot",
        Description = "Render an R plot to a PNG and return the file path.",
        Category = "RExcelBridge")]
    public static object RPlot(
        string code,
        object plotName,
        object width,
        object height,
        object triggerRange)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(code))
                return "Error: code is blank.";

            _ = NormalizeTrigger(triggerRange);

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
        Name = "RPlotDataNamed",
        Description = "Build an R data frame from named Excel ranges, run plotting code against df, and return the plot path.",
        Category = "RExcelBridge",
        IsVolatile = true)]
    public static object RPlotDataNamed(
        [ExcelArgument(Name = "code", Description = "R plotting code that uses data frame df.")] object code,
        [ExcelArgument(Name = "plot_name", Description = "Optional stable plot label.")] object plot_name,
        [ExcelArgument(Name = "width", Description = "PNG width in pixels.")] object width,
        [ExcelArgument(Name = "height", Description = "PNG height in pixels.")] object height,
        [ExcelArgument(Name = "recalc_key", Description = "Optional recalc trigger.")] object recalc_key,

        [ExcelArgument(Name = "name1")] object name1, [ExcelArgument(Name = "range1")] object range1,
        [ExcelArgument(Name = "name2")] object name2, [ExcelArgument(Name = "range2")] object range2,
        [ExcelArgument(Name = "name3")] object name3, [ExcelArgument(Name = "range3")] object range3,
        [ExcelArgument(Name = "name4")] object name4, [ExcelArgument(Name = "range4")] object range4,
        [ExcelArgument(Name = "name5")] object name5, [ExcelArgument(Name = "range5")] object range5,
        [ExcelArgument(Name = "name6")] object name6, [ExcelArgument(Name = "range6")] object range6,
        [ExcelArgument(Name = "name7")] object name7, [ExcelArgument(Name = "range7")] object range7,
        [ExcelArgument(Name = "name8")] object name8, [ExcelArgument(Name = "range8")] object range8,
        [ExcelArgument(Name = "name9")] object name9, [ExcelArgument(Name = "range9")] object range9,
        [ExcelArgument(Name = "name10")] object name10, [ExcelArgument(Name = "range10")] object range10)
    {
        try
        {
            _ = NormalizeTrigger(recalc_key);

            string rCode = ToText(code);
            if (string.IsNullOrWhiteSpace(rCode))
                return "Error: code is blank.";

            string plotTag = ToText(plot_name);
            if (string.IsNullOrWhiteSpace(plotTag))
                plotTag = "RPlotDataNamed";

            int pngWidth = ToIntOrDefault(width, 900);
            int pngHeight = ToIntOrDefault(height, 600);

            var pairs = new List<(string Name, object Range)>
            {
                (ToText(name1), range1),
                (ToText(name2), range2),
                (ToText(name3), range3),
                (ToText(name4), range4),
                (ToText(name5), range5),
                (ToText(name6), range6),
                (ToText(name7), range7),
                (ToText(name8), range8),
                (ToText(name9), range9),
                (ToText(name10), range10)
            };

            var cols = new List<(string Name, object Values, int Length)>();

            foreach (var pair in pairs)
            {
                bool hasName = !string.IsNullOrWhiteSpace(pair.Name);
                bool hasRange = HasUsableRange(pair.Range);

                if (!hasName && !hasRange)
                    continue;

                if (hasName && !hasRange)
                    return $"Error: range missing for column '{pair.Name}'.";

                if (!hasName && hasRange)
                    return "Error: a range was supplied without a column name.";

                object normalized = NormalizeRangeToColumnVector(pair.Range);
                int length = GetVectorLength(normalized);

                if (length == 0)
                    return $"Error: range for column '{pair.Name}' is empty.";

                cols.Add((pair.Name, normalized, length));
            }

            if (cols.Count == 0)
                return "Error: no named ranges were supplied.";

            int expectedLength = cols[0].Length;
            foreach (var col in cols)
            {
                if (col.Length != expectedLength)
                    return $"Error: column '{col.Name}' has {col.Length} values, expected {expectedLength}.";
            }

            var tempNames = new List<string>();
            for (int i = 0; i < cols.Count; i++)
            {
                string tempName = $".__rexcel_plot_col_{Guid.NewGuid():N}_{i + 1}";
                tempNames.Add(tempName);

                object setResult = RBridge.Set(tempName, cols[i].Values);
                if (setResult is string setError &&
                    (setError.StartsWith("Error", StringComparison.OrdinalIgnoreCase) ||
                     setError.StartsWith("R error:", StringComparison.OrdinalIgnoreCase)))
                    return setError;
            }

            string assignDfCode = BuildNamedDataFrameCode(tempNames, cols.Select(c => c.Name).ToList());
            object assignResult = RBridge.Eval(assignDfCode);
            if (assignResult is string assignError &&
                (assignError.StartsWith("Error", StringComparison.OrdinalIgnoreCase) ||
                 assignError.StartsWith("R error:", StringComparison.OrdinalIgnoreCase)))
                return assignError;

            return RBridge.Plot(rCode, plotTag, pngWidth, pngHeight);
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    private static string ToText(object value)
    {
        if (value is null || value is ExcelMissing || value is ExcelEmpty)
            return string.Empty;

        return Convert.ToString(value)?.Trim() ?? string.Empty;
    }

    private static bool HasUsableRange(object value)
    {
        if (value is null || value is ExcelMissing || value is ExcelEmpty)
            return false;

        if (value is string s)
            return !string.IsNullOrWhiteSpace(s);

        if (value is object[,] arr)
        {
            int rowMin = arr.GetLowerBound(0);
            int rowMax = arr.GetUpperBound(0);
            int colMin = arr.GetLowerBound(1);
            int colMax = arr.GetUpperBound(1);

            for (int r = rowMin; r <= rowMax; r++)
            {
                for (int c = colMin; c <= colMax; c++)
                {
                    object cell = arr[r, c];
                    if (cell is not null && cell is not ExcelMissing && cell is not ExcelEmpty)
                        return true;
                }
            }

            return false;
        }

        return true;
    }

    private static object NormalizeRangeToColumnVector(object value)
    {
        if (value is null || value is ExcelMissing || value is ExcelEmpty)
            return Array.Empty<object?>();

        if (value is object[,] arr)
        {
            int rowMin = arr.GetLowerBound(0);
            int rowMax = arr.GetUpperBound(0);
            int colMin = arr.GetLowerBound(1);
            int colMax = arr.GetUpperBound(1);
            int rows = rowMax - rowMin + 1;
            int cols = colMax - colMin + 1;

            if (rows == 1 && cols == 1)
                return new object?[] { NormalizePlotScalar(arr[rowMin, colMin]) };

            if (cols == 1)
            {
                var output = new object?[rows];
                for (int r = 0; r < rows; r++)
                    output[r] = NormalizePlotScalar(arr[rowMin + r, colMin]);
                return output;
            }

            if (rows == 1)
            {
                var output = new object?[cols];
                for (int c = 0; c < cols; c++)
                    output[c] = NormalizePlotScalar(arr[rowMin, colMin + c]);
                return output;
            }

            var flattened = new object?[rows * cols];
            int k = 0;
            for (int r = rowMin; r <= rowMax; r++)
            {
                for (int c = colMin; c <= colMax; c++)
                    flattened[k++] = NormalizePlotScalar(arr[r, c]);
            }

            return flattened;
        }

        return new object?[] { NormalizePlotScalar(value) };
    }

    private static object? NormalizePlotScalar(object? value)
    {
        if (value is null || value is ExcelEmpty || value is ExcelMissing || value is ExcelError)
            return null;

        return value switch
        {
            double d => (double.IsNaN(d) || double.IsInfinity(d)) ? null : d,
            float f => (float.IsNaN(f) || float.IsInfinity(f)) ? null : (double)f,
            int i => i,
            long l => l,
            short s => (int)s,
            decimal m => m,
            bool b => b,
            string s => s,
            DateTime dt => dt.ToString("o"),
            _ => value.ToString()
        };
    }

    private static int GetVectorLength(object value)
    {
        if (value is Array arr)
            return arr.Length;

        return value is null ? 0 : 1;
    }

    private static string BuildNamedDataFrameCode(IReadOnlyList<string> tempNames, IReadOnlyList<string> columnNames)
    {
        var parts = new List<string>();

        for (int i = 0; i < tempNames.Count; i++)
        {
            string tempName = tempNames[i];
            string colName = EscapeRString(columnNames[i]);
            parts.Add($"\"{colName}\" = get(\"{EscapeRString(tempName)}\", envir = .GlobalEnv)");
        }

        return "df <- data.frame(" + string.Join(", ", parts) + ", check.names = FALSE, stringsAsFactors = FALSE)";
    }

    private static string EscapeRString(string text)
    {
        return (text ?? string.Empty).Replace("\\", "\\\\").Replace("\"", "\\\"");
    }
}