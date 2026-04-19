using System;
using System.Collections.Generic;
using System.Linq;
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
        Name = "PPlot",
        Description = "Render a Python plot to a PNG and return the file path.",
        Category = "PythonExcelBridge")]
    public static object PPlot(
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
        [ExcelArgument(Name = "file", Description = "Path to a Python script file.")] string file)
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
        Name = "PPlotDataNamed",
        Description = "Assign named Excel ranges to Python objects, run plotting code, and return the plot path.",
        Category = "PythonExcelBridge",
        IsVolatile = true)]
    public static object PPlotDataNamed(
        [ExcelArgument(Name = "code", Description = "Python plotting code that uses named objects such as x and y.")] object code,
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
            string pyCode = ToText(code);
            if (string.IsNullOrWhiteSpace(pyCode))
                return "Error: code is blank.";

            string plotTag = ToText(plot_name);
            if (string.IsNullOrWhiteSpace(plotTag))
                plotTag = "PPlotDataNamed";

            int pngWidth = ToIntOrDefault(width, 900);
            int pngHeight = ToIntOrDefault(height, 600);

            object[] rawNames =
            {
                name1, name2, name3, name4, name5,
                name6, name7, name8, name9, name10
            };

            object[] rawRanges =
            {
                range1, range2, range3, range4, range5,
                range6, range7, range8, range9, range10
            };

            bool recalcProvided = LooksLikeRecalcKey(recalc_key);
            if (recalcProvided)
                _ = NormalizeTrigger(recalc_key);

            // Support both calling styles:
            // 1) with recalc_key:
            //    PPlotDataNamed(code, plot, w, h, recalc_key, "x", A1:A10, "y", B1:B10)
            // 2) without recalc_key:
            //    PPlotDataNamed(code, plot, w, h, "x", A1:A10, "y", B1:B10)
            //
            // If recalc_key is omitted, Excel shifts the pairs left, so:
            // recalc_key actually contains name1,
            // name1 contains range1,
            // name2 contains range2, etc.
            List<(string Name, object Range)> pairs = BuildPairs(recalc_key, rawNames, rawRanges, recalcProvided);

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

            foreach (var col in cols)
            {
                object setResult = PythonBridge.Set(col.Name, col.Values);
                if (IsBridgeError(setResult, out string setError))
                    return setError;
            }

            return PythonBridge.Plot(pyCode, plotTag, pngWidth, pngHeight);
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

    private static bool LooksLikeRecalcKey(object value)
    {
        if (value is null || value is ExcelMissing || value is ExcelEmpty)
            return false;

        if (value is object[,])
            return true;

        if (value is double or float or int or long or short or decimal or bool or DateTime)
            return true;

        return false;
    }

    private static List<(string Name, object Range)> BuildPairs(object recalcKey, object[] rawNames, object[] rawRanges, bool recalcProvided)
    {
        var pairs = new List<(string Name, object Range)>();

        if (recalcProvided)
        {
            for (int i = 0; i < rawNames.Length; i++)
                pairs.Add((ToText(rawNames[i]), rawRanges[i]));
            return pairs;
        }

        // Shift left by one pair because recalc_key is actually the first name.
        pairs.Add((ToText(recalcKey), rawNames[0]));

        for (int i = 1; i < rawNames.Length; i++)
            pairs.Add((ToText(rawRanges[i - 1]), rawNames[i]));

        return pairs;
    }

    private static bool IsBridgeError(object value, out string message)
    {
        if (value is string s &&
            (s.StartsWith("Error", StringComparison.OrdinalIgnoreCase) ||
             s.StartsWith("Python error:", StringComparison.OrdinalIgnoreCase)))
        {
            message = s;
            return true;
        }

        message = string.Empty;
        return false;
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
}