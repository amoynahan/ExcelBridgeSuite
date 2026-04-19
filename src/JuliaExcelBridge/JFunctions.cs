using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ExcelDna.Integration;

namespace JuliaExcelBridge;

public static class JFunctions
{
    [ExcelFunction(
        Name = "JPing",
        Description = "Ping the persistent Julia worker and return a simple response.",
        Category = "JuliaExcelBridge")]
    public static object JPing()
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
    public static object JEval(
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
            return Convert.ToInt32(value, CultureInfo.InvariantCulture);
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

            int rowMin = arr.GetLowerBound(0);
            int rowMax = arr.GetUpperBound(0);
            int colMin = arr.GetLowerBound(1);
            int colMax = arr.GetUpperBound(1);

            for (int r = rowMin; r <= rowMax; r++)
            {
                for (int c = colMin; c <= colMax; c++)
                {
                    if (arr[r, c] is double d)
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
        Name = "JPlot",
        Description = "Render a Julia plot to a PNG and return the file path.",
        Category = "JuliaExcelBridge")]
    public static object JPlot(
        [ExcelArgument(Name = "code", Description = "Julia plotting code or a plotting function call, such as plot(1:10) or my_plot().")] string code,
        [ExcelArgument(Name = "plot_name", Description = "Optional stable plot label used in the file name.")] object plotName,
        [ExcelArgument(Name = "width", Description = "PNG width in pixels.")] object width,
        [ExcelArgument(Name = "height", Description = "PNG height in pixels.")] object height,
        [ExcelArgument(Name = "trigger_range", Description = "Optional recalc trigger.")] object triggerRange)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(code))
                return "Error: code is blank.";

            _ = NormalizeTrigger(triggerRange);

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
        Description = "Source a Julia script file into the persistent Julia session.",
        Category = "JuliaExcelBridge")]
    public static object JSource(
        [ExcelArgument(Name = "file", Description = "Path to a Julia script file.")] string file)
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
    public static object JCall(
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
    public static object JSet(
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
    public static object JGet(
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
    public static object JExists(
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
    public static object JRemove(
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
    public static object JObjects()
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
    public static object JDescribe(
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
        Name = "JPlotDataNamed",
        Description = "Build Julia vectors from named Excel ranges, create df, and return the plot path.",
        Category = "JuliaExcelBridge",
        IsVolatile = true)]
    public static object JPlotDataNamed(
        [ExcelArgument(Name = "code", Description = "Julia plotting code that uses df and/or the supplied column names.")] object code,
        [ExcelArgument(Name = "plot_name", Description = "Optional stable plot label.")] object plotName,
        [ExcelArgument(Name = "width", Description = "PNG width in pixels.")] object width,
        [ExcelArgument(Name = "height", Description = "PNG height in pixels.")] object height,
        [ExcelArgument(Name = "recalc_key", Description = "Optional recalc trigger.")] object recalcKey,
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
            _ = NormalizeTrigger(recalcKey);

            string jCode = ToText(code);
            if (string.IsNullOrWhiteSpace(jCode))
                return "Error: code is blank.";

            string plotTag = ToText(plotName);
            if (string.IsNullOrWhiteSpace(plotTag))
                plotTag = "JPlotDataNamed";

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

            var cols = new List<(string OriginalName, string JuliaName, object?[] Values, int Length)>();
            var seen = new HashSet<string>(StringComparer.Ordinal);

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

                string juliaName = ToJuliaIdentifier(pair.Name);
                if (!seen.Add(juliaName))
                    return $"Error: duplicate column name '{pair.Name}'.";

                object?[] values = NormalizeRangeToColumnVector(pair.Range);
                if (values.Length == 0)
                    return $"Error: range for column '{pair.Name}' is empty.";

                cols.Add((pair.Name, juliaName, values, values.Length));
            }

            if (cols.Count == 0)
                return "Error: no named ranges were supplied.";

            int expectedLength = cols[0].Length;
            foreach (var col in cols)
            {
                if (col.Length != expectedLength)
                    return $"Error: column '{col.OriginalName}' has {col.Length} values, expected {expectedLength}.";
            }

            string setupCode = BuildNamedDfCode(cols);
            object assignResult = JBridge.Eval(setupCode);
            if (assignResult is string assignError && IsBridgeError(assignError))
                return assignError;

            return JBridge.Plot(jCode, plotTag, pngWidth, pngHeight);
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
    public static object JPlotPath()
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

    private static bool IsBridgeError(string message)
    {
        return !string.IsNullOrWhiteSpace(message) &&
               (message.StartsWith("Error", StringComparison.OrdinalIgnoreCase) ||
                message.StartsWith("Julia error:", StringComparison.OrdinalIgnoreCase));
    }

    private static string ToText(object value)
    {
        if (value is null || value is ExcelMissing || value is ExcelEmpty)
            return string.Empty;

        return Convert.ToString(value, CultureInfo.InvariantCulture)?.Trim() ?? string.Empty;
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

    private static object?[] NormalizeRangeToColumnVector(object value)
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
                return new[] { NormalizePlotScalar(arr[rowMin, colMin]) };

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

        return new[] { NormalizePlotScalar(value) };
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
            DateTime dt => dt,
            _ => value.ToString()
        };
    }

    private static string ToJuliaIdentifier(string name)
    {
        string trimmed = ToText(name);
        if (string.IsNullOrWhiteSpace(trimmed))
            return "col";

        string cleaned = Regex.Replace(trimmed, @"[^A-Za-z0-9_]", "_");
        if (string.IsNullOrWhiteSpace(cleaned))
            cleaned = "col";
        if (char.IsDigit(cleaned[0]))
            cleaned = "col_" + cleaned;
        return cleaned;
    }

    private static string BuildNamedDfCode(IReadOnlyList<(string OriginalName, string JuliaName, object?[] Values, int Length)> cols)
    {
        var sb = new StringBuilder();

        foreach (var col in cols)
        {
            sb.Append(col.JuliaName)
              .Append(" = ")
              .Append(BuildJuliaVectorLiteral(col.Values))
              .Append('\n');
        }

        sb.Append("df = (; ");
        for (int i = 0; i < cols.Count; i++)
        {
            if (i > 0)
                sb.Append(", ");

            sb.Append(cols[i].JuliaName)
              .Append(" = ")
              .Append(cols[i].JuliaName);
        }
        sb.Append(")\n");
        sb.Append("nothing");

        return sb.ToString();
    }

    private static string BuildJuliaVectorLiteral(IEnumerable<object?> values)
    {
        return "[" + string.Join(", ", values.Select(ToJuliaLiteral)) + "]";
    }

    private static string ToJuliaLiteral(object? value)
    {
        if (value is null)
            return "missing";

        return value switch
        {
            double d => d.ToString("R", CultureInfo.InvariantCulture),
            float f => ((double)f).ToString("R", CultureInfo.InvariantCulture),
            decimal m => m.ToString(CultureInfo.InvariantCulture),
            int i => i.ToString(CultureInfo.InvariantCulture),
            long l => l.ToString(CultureInfo.InvariantCulture),
            short s => s.ToString(CultureInfo.InvariantCulture),
            bool b => b ? "true" : "false",
            DateTime dt => $"\"{EscapeJuliaString(dt.ToString("o", CultureInfo.InvariantCulture))}\"",
            string s => $"\"{EscapeJuliaString(s)}\"",
            _ => $"\"{EscapeJuliaString(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty)}\""
        };
    }

    private static string EscapeJuliaString(string value)
    {
        return value
            .Replace("\\", "\\\\")
            .Replace("\"", "\\\"")
            .Replace("\r", "\\r")
            .Replace("\n", "\\n");
    }
}