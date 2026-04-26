from pathlib import Path
p=Path('/mnt/data/pywork')
# We'll replace PythonBridge.cs entirely with an enhanced version based on existing structure
cs = r'''using ExcelDna.Integration;
using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace PythonExcelBridge;

public static class PythonBridge
{
    private static Process? _proc;
    private static StreamWriter? _stdin;
    private static StreamReader? _stdout;
    private static StreamReader? _stderr;
    private static readonly object _lock = new();
    internal const string PythonObjectReferencePrefix = "__PYEXCEL_PYOBJ__:";

    public static string MakeObjectReference(string name)
    {
        return PythonObjectReferencePrefix + (name ?? string.Empty).Trim();
    }

    private static string AddInDir
    {
        get
        {
            string? xllPath = ExcelDnaUtil.XllPath;
            if (!string.IsNullOrWhiteSpace(xllPath))
            {
                string? dir = Path.GetDirectoryName(xllPath);
                if (!string.IsNullOrWhiteSpace(dir))
                    return dir;
            }
            return AppDomain.CurrentDomain.BaseDirectory;
        }
    }

    private static string WorkerPath => Path.Combine(AddInDir, "worker.py");
    private static string StartupPath => Path.Combine(AddInDir, "startup.py");
    private static string LogPath => Path.Combine(AddInDir, "python_worker_stderr.log");
    private static string StartupErrorLogPath => Path.Combine(AddInDir, "addin_startup_error.log");
    private static string PythonConfigPath => Path.Combine(AddInDir, "python-path.txt");
    private static string PlotConfigPath => Path.Combine(AddInDir, "plot-path.txt");

    public static void TryStart()
    {
        try { Start(); }
        catch (Exception ex) { SafeLogStartupError(ex); }
    }

    public static void Start()
    {
        lock (_lock)
        {
            if (_proc is { HasExited: false }) return;
            if (!File.Exists(WorkerPath)) throw new FileNotFoundException("worker.py was not found.", WorkerPath);
            if (!File.Exists(StartupPath)) throw new FileNotFoundException("startup.py was not found.", StartupPath);
            string pythonPath = GetPythonPath();
            var psi = new ProcessStartInfo
            {
                FileName = pythonPath,
                Arguments = $"\"{WorkerPath}\" \"{StartupPath}\"",
                UseShellExecute = false,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                WorkingDirectory = AddInDir
            };
            _proc = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start the Python worker process.");
            _stdin = _proc.StandardInput;
            _stdout = _proc.StandardOutput;
            _stderr = _proc.StandardError;
            BeginErrorCapture();
            object pingResult = Ping();
            if (pingResult is string s &&
                (s.StartsWith("Python error:", StringComparison.OrdinalIgnoreCase) || s.Contains("no response", StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException($"Python worker failed startup ping: {s}");
        }
    }

    public static void Stop()
    {
        lock (_lock)
        {
            try
            {
                _stdin?.Close();
                if (_proc is { HasExited: false }) _proc.Kill(entireProcessTree: true);
            }
            catch { }
            finally
            {
                _stdin?.Dispose(); _stdout?.Dispose(); _stderr?.Dispose(); _proc?.Dispose();
                _stdin = null; _stdout = null; _stderr = null; _proc = null;
            }
        }
    }

    public static object Ping() => Send(new Dictionary<string, object?> { ["id"] = Guid.NewGuid().ToString(), ["cmd"] = "ping" });
    public static object Source(string file) => Send(new Dictionary<string, object?> { ["id"] = Guid.NewGuid().ToString(), ["cmd"] = "source", ["file"] = file });
    public static object Eval(string code) => Send(new Dictionary<string, object?> { ["id"] = Guid.NewGuid().ToString(), ["cmd"] = "eval", ["code"] = code });

    public static object Call(string fun, params object?[] args)
    {
        return Send(new Dictionary<string, object?>
        {
            ["id"] = Guid.NewGuid().ToString(),
            ["cmd"] = "call",
            ["fun"] = fun,
            ["args"] = NormalizeArgs(args)
        });
    }

    public static object Set(string name, object value)
    {
        if (TryCreateNumericSetPayload(name, value, out Dictionary<string, object?>? fastPayload, out string? transferFile))
        {
            try { return Send(fastPayload!); }
            finally { TryDeleteFile(transferFile ?? string.Empty); }
        }
        return Send(new Dictionary<string, object?> { ["id"] = Guid.NewGuid().ToString(), ["cmd"] = "set", ["name"] = name, ["value"] = NormalizeExcelValue(value) });
    }

    public static object SetTable(string name, object value, bool hasHeaders = true)
    {
        if (TryCreateTableSetPayload(name, value, hasHeaders, out Dictionary<string, object?>? payload, out List<string>? transferFiles))
        {
            try { return Send(payload!); }
            finally
            {
                if (transferFiles is not null)
                    foreach (string file in transferFiles) TryDeleteFile(file);
            }
        }
        return "Error: PSetTable expects a rectangular Excel range.";
    }

    public static object Get(string name) => Send(new Dictionary<string, object?> { ["id"] = Guid.NewGuid().ToString(), ["cmd"] = "get", ["name"] = name });
    public static object GetNumeric(string name) => Send(new Dictionary<string, object?> { ["id"] = Guid.NewGuid().ToString(), ["cmd"] = "get_numeric", ["name"] = name });
    public static object GetTable(string name) => Send(new Dictionary<string, object?> { ["id"] = Guid.NewGuid().ToString(), ["cmd"] = "get_table", ["name"] = name });
    public static object LastTransfer() => Send(new Dictionary<string, object?> { ["id"] = Guid.NewGuid().ToString(), ["cmd"] = "last_transfer" });
    public static object Exists(string name) => Send(new Dictionary<string, object?> { ["id"] = Guid.NewGuid().ToString(), ["cmd"] = "exists", ["name"] = name });
    public static object Remove(string name) => Send(new Dictionary<string, object?> { ["id"] = Guid.NewGuid().ToString(), ["cmd"] = "remove", ["name"] = name });
    public static object Objects() => Send(new Dictionary<string, object?> { ["id"] = Guid.NewGuid().ToString(), ["cmd"] = "objects" });
    public static object Describe(string name) => Send(new Dictionary<string, object?> { ["id"] = Guid.NewGuid().ToString(), ["cmd"] = "describe", ["name"] = name });

    public static object Plot(string code, string? plotName = null, int width = 800, int height = 600, int res = 96)
    {
        string file = BuildPlotPath(plotName);
        return Send(new Dictionary<string, object?> { ["id"] = Guid.NewGuid().ToString(), ["cmd"] = "plot", ["code"] = code, ["file"] = file, ["width"] = width, ["height"] = height, ["res"] = res });
    }

    public static object InsertPlotFromSelectedCell()
    {
        try
        {
            dynamic? app = ExcelDnaUtil.Application;
            dynamic? target = app?.ActiveCell;
            if (target is null) return "Error: no active cell was found.";
            object rawValue = target.Value2;
            string? file = rawValue?.ToString();
            if (string.IsNullOrWhiteSpace(file)) return "Error: selected cell is blank. Select a cell containing the PNG path returned by PPlot.";
            return InsertPlotFile(file);
        }
        catch (Exception ex) { return $"Error: {ex.Message}"; }
    }

    public static object InsertPlotFile(string file)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(file)) return "Error: plot file path is blank.";
            string fullPath = Path.GetFullPath(Environment.ExpandEnvironmentVariables(file.Trim().Trim('"')));
            if (!File.Exists(fullPath)) return $"Error: plot file was not found: {fullPath}";
            dynamic? app = ExcelDnaUtil.Application;
            dynamic? sheet = app?.ActiveSheet;
            dynamic? target = app?.ActiveCell;
            if (sheet is null || target is null) return "Error: could not resolve the active worksheet or cell.";
            string cellAddress = "Cell"; string sheetName = "Sheet";
            try { cellAddress = target.Address[false, false]?.ToString() ?? "Cell"; } catch { }
            try { sheetName = sheet.Name?.ToString() ?? "Sheet"; } catch { }
            string shapeName = $"PPlot_{SanitizeFileComponent(sheetName)}_{SanitizeFileComponent(cellAddress)}";
            try { dynamic oldShape = sheet.Shapes.Item(shapeName); oldShape.Delete(); } catch { }
            float left = Convert.ToSingle(target.Left);
            float top = Convert.ToSingle(target.Top + target.Height + 6);
            dynamic shape = sheet.Shapes.AddPicture(fullPath, 0, -1, left, top, -1, -1);
            shape.Name = shapeName;
            try { shape.LockAspectRatio = -1; } catch { }
            return fullPath;
        }
        catch (Exception ex) { return $"Error: {ex.Message}"; }
    }

    public static string GetPlotDirectory()
    {
        string defaultDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "PythonExcelBridge", "PlotCache");
        if (File.Exists(PlotConfigPath))
        {
            string configured = File.ReadAllText(PlotConfigPath).Trim();
            if (!string.IsNullOrWhiteSpace(configured))
            {
                try { string candidate = ExpandConfiguredPath(configured); if (Directory.Exists(candidate)) return candidate; } catch { }
            }
        }
        Directory.CreateDirectory(defaultDir);
        return defaultDir;
    }

    private static string ExpandConfiguredPath(string configured) => Path.GetFullPath(Environment.ExpandEnvironmentVariables(configured.Trim().Trim('"')));

    private static string BuildPlotPath(string? plotName)
    {
        string plotDir = GetPlotDirectory();
        string workbook = SanitizeFileComponent(GetActiveWorkbookName());
        string sheet = SanitizeFileComponent(GetCallerSheetName());
        string cell = SanitizeFileComponent(GetCallerAddressA1());
        string label = string.IsNullOrWhiteSpace(plotName) ? "plot" : SanitizeFileComponent(plotName!);
        string fileName = $"{workbook}_{sheet}_{cell}_{label}.png";
        return Path.Combine(plotDir, fileName);
    }

    private static string GetActiveWorkbookName()
    {
        try { dynamic app = ExcelDnaUtil.Application; string? name = app?.ActiveWorkbook?.Name as string; if (!string.IsNullOrWhiteSpace(name)) return Path.GetFileNameWithoutExtension(name); } catch { }
        return "Workbook";
    }
    private static string GetCallerSheetName()
    {
        try { dynamic app = ExcelDnaUtil.Application; string? name = app?.ActiveSheet?.Name as string; if (!string.IsNullOrWhiteSpace(name)) return name; } catch { }
        return "Sheet";
    }
    private static string GetCallerAddressA1()
    {
        try
        {
            object caller = XlCall.Excel(XlCall.xlfCaller);
            if (caller is ExcelReference xr)
            {
                int row = xr.RowFirst + 1;
                int col = xr.ColumnFirst + 1;
                return $"{ColumnNumberToLetters(col)}{row}";
            }
        }
        catch { }
        return "Cell";
    }
    private static string ColumnNumberToLetters(int columnNumber)
    {
        var letters = string.Empty; int col = columnNumber;
        while (col > 0) { int rem = (col - 1) % 26; letters = (char)('A' + rem) + letters; col = (col - 1) / 26; }
        return letters;
    }
    private static string SanitizeFileComponent(string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return "item";
        string cleaned = Regex.Replace(value, @"[^A-Za-z0-9._-]+", "_").Trim('_', '.', ' ');
        return string.IsNullOrWhiteSpace(cleaned) ? "item" : cleaned;
    }

    private static object Send(Dictionary<string, object?> payload)
    {
        lock (_lock)
        {
            EnsureRunning();
            string json = JsonSerializer.Serialize(payload);
            _stdin!.WriteLine(json);
            _stdin.Flush();
            string? line = _stdout!.ReadLine();
            if (string.IsNullOrWhiteSpace(line)) return "Python worker returned no response.";
            using JsonDocument doc = JsonDocument.Parse(line);
            JsonElement root = doc.RootElement;
            bool ok = root.TryGetProperty("ok", out JsonElement okEl) && okEl.GetBoolean();
            if (!ok)
            {
                if (root.TryGetProperty("error", out JsonElement errEl)) return $"Python error: {errEl.GetString()}";
                return "Unknown Python error.";
            }
            if (!root.TryGetProperty("result", out JsonElement resultEl)) return ExcelEmpty.Value;
            return ConvertJsonToExcel(resultEl);
        }
    }

    private static void EnsureRunning() { if (_proc is null || _proc.HasExited) Start(); }
    private static void BeginErrorCapture()
    {
        string path = LogPath;
        Task.Run(() => { try { using var sw = new StreamWriter(path, append: true); while (_stderr is not null && !_stderr.EndOfStream) { string? line = _stderr.ReadLine(); if (!string.IsNullOrWhiteSpace(line)) { sw.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | {line}"); sw.Flush(); } } } catch { } });
    }
    private static void SafeLogStartupError(Exception ex)
    {
        try { File.AppendAllText(StartupErrorLogPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {ex}{Environment.NewLine}{Environment.NewLine}"); } catch { }
    }

    private static string GetPythonPath()
    {
        if (File.Exists(PythonConfigPath))
        {
            string configured = File.ReadAllText(PythonConfigPath).Trim();
            if (!string.IsNullOrWhiteSpace(configured) && File.Exists(configured)) return configured;
        }
        string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        string[] localCandidates =
        {
            Path.Combine(localAppData, "Microsoft", "WindowsApps", "python.exe"),
            Path.Combine(localAppData, "Programs", "Python", "Python313", "python.exe"),
            Path.Combine(localAppData, "Programs", "Python", "Python312", "python.exe"),
            Path.Combine(localAppData, "Programs", "Python", "Python311", "python.exe"),
            Path.Combine(localAppData, "Programs", "Python", "Python310", "python.exe"),
            Path.Combine(userProfile, "AppData", "Local", "Programs", "Python", "Python313", "python.exe"),
            Path.Combine(userProfile, "AppData", "Local", "Programs", "Python", "Python312", "python.exe"),
            Path.Combine(userProfile, "AppData", "Local", "Programs", "Python", "Python311", "python.exe"),
            Path.Combine(userProfile, "AppData", "Local", "Programs", "Python", "Python310", "python.exe"),
            Path.Combine(AddInDir, "Python", "python.exe"),
            Path.Combine(AddInDir, "runtime", "Python", "python.exe")
        };
        foreach (string candidate in localCandidates) if (File.Exists(candidate)) return candidate;
        string[] programRoots = { Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86) };
        foreach (string root in programRoots.Where(r => !string.IsNullOrWhiteSpace(r) && Directory.Exists(r)))
        {
            try
            {
                foreach (string dir in Directory.EnumerateDirectories(root, "Python*", SearchOption.TopDirectoryOnly).OrderByDescending(x => x))
                {
                    string candidate = Path.Combine(dir, "python.exe");
                    if (File.Exists(candidate)) return candidate;
                }
            }
            catch { }
        }
        return "python";
    }

    private static bool TryCreateTableSetPayload(string name, object value, bool hasHeaders, out Dictionary<string, object?>? payload, out List<string>? transferFiles)
    {
        payload = null;
        transferFiles = new List<string>();
        if (!TryExtractExcelRange(value, out object?[,]? cells, out int rows, out int cols)) return false;
        if (cells is null || rows <= 0 || cols <= 0) return false;
        int headerRows = hasHeaders ? 1 : 0;
        if (rows <= headerRows) return false;
        int dataRows = rows - headerRows;
        string[] names = new string[cols];
        for (int c = 0; c < cols; c++)
        {
            string? raw = hasHeaders ? NormalizeScalar(cells[0, c])?.ToString() : null;
            names[c] = string.IsNullOrWhiteSpace(raw) ? $"V{c + 1}" : raw!.Trim();
        }
        string dir = Path.Combine(Path.GetTempPath(), "PythonExcelBridgeTransfer");
        Directory.CreateDirectory(dir);
        var columns = new List<Dictionary<string, object?>>();
        for (int c = 0; c < cols; c++)
        {
            if (ColumnIsNumeric(cells, c, headerRows, rows))
            {
                var values = new double[dataRows];
                for (int r = headerRows; r < rows; r++)
                {
                    if (!TryConvertToDoubleOrNaN(cells[r, c], out double d, allowEmptyAsNaN: true)) d = double.NaN;
                    values[r - headerRows] = d;
                }
                string file = Path.Combine(dir, $"excel_table_col_{c + 1}_{Guid.NewGuid():N}.bin");
                byte[] bytes = new byte[checked(values.Length * sizeof(double))];
                Buffer.BlockCopy(values, 0, bytes, 0, bytes.Length);
                File.WriteAllBytes(file, bytes);
                transferFiles.Add(file);
                columns.Add(new Dictionary<string, object?> { ["name"] = names[c], ["type"] = "numeric", ["file"] = file.Replace('\\', '/'), ["na"] = "NaN" });
            }
            else if (ColumnIsLogical(cells, c, headerRows, rows))
            {
                var vals = new object?[dataRows];
                for (int r = headerRows; r < rows; r++) vals[r - headerRows] = TryConvertToNullableBool(cells[r, c], out bool? b) ? b : null;
                columns.Add(new Dictionary<string, object?> { ["name"] = names[c], ["type"] = "logical", ["values"] = vals });
            }
            else
            {
                var vals = new object?[dataRows];
                for (int r = headerRows; r < rows; r++) vals[r - headerRows] = NormalizeStringCell(cells[r, c]);
                columns.Add(new Dictionary<string, object?> { ["name"] = names[c], ["type"] = "character", ["values"] = vals });
            }
        }
        payload = new Dictionary<string, object?> { ["id"] = Guid.NewGuid().ToString(), ["cmd"] = "set_table", ["name"] = name, ["rows"] = dataRows, ["cols"] = cols, ["has_headers"] = hasHeaders, ["columns"] = columns };
        return true;
    }

    private static bool TryExtractExcelRange(object value, out object?[,]? cells, out int rows, out int cols)
    {
        cells = null; rows = 0; cols = 0;
        if (value is object[,] range)
        {
            int rowMin = range.GetLowerBound(0), rowMax = range.GetUpperBound(0), colMin = range.GetLowerBound(1), colMax = range.GetUpperBound(1);
            rows = rowMax - rowMin + 1; cols = colMax - colMin + 1;
            cells = new object?[rows, cols];
            for (int r = 0; r < rows; r++) for (int c = 0; c < cols; c++) cells[r, c] = range[rowMin + r, colMin + c];
            return true;
        }
        if (value is object[] vector)
        {
            rows = vector.Length; cols = 1; cells = new object?[rows, 1];
            for (int r = 0; r < rows; r++) cells[r, 0] = vector[r];
            return rows > 0;
        }
        if (value is null || value is ExcelMissing || value is ExcelEmpty || value is ExcelError) return false;
        rows = 1; cols = 1; cells = new object?[1, 1]; cells[0, 0] = value; return true;
    }

    private static bool ColumnIsNumeric(object?[,] cells, int c, int startRow, int rows)
    {
        bool sawValue = false;
        for (int r = startRow; r < rows; r++)
        {
            if (IsBlankExcelCell(cells[r, c])) continue;
            if (!TryConvertToDoubleOrNaN(cells[r, c], out _, allowEmptyAsNaN: true)) return false;
            sawValue = true;
        }
        return sawValue;
    }

    private static bool ColumnIsLogical(object?[,] cells, int c, int startRow, int rows)
    {
        bool sawValue = false;
        for (int r = startRow; r < rows; r++)
        {
            if (IsBlankExcelCell(cells[r, c])) continue;
            if (!TryConvertToNullableBool(cells[r, c], out _)) return false;
            sawValue = true;
        }
        return sawValue;
    }

    private static bool TryConvertToNullableBool(object? value, out bool? result)
    {
        result = null;
        if (IsBlankExcelCell(value)) return true;
        if (value is bool b) { result = b; return true; }
        if (value is string s)
        {
            string t = s.Trim();
            if (t.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || t.Equals("T", StringComparison.OrdinalIgnoreCase) || t.Equals("YES", StringComparison.OrdinalIgnoreCase)) { result = true; return true; }
            if (t.Equals("FALSE", StringComparison.OrdinalIgnoreCase) || t.Equals("F", StringComparison.OrdinalIgnoreCase) || t.Equals("NO", StringComparison.OrdinalIgnoreCase)) { result = false; return true; }
        }
        return false;
    }

    private static bool IsBlankExcelCell(object? value) => value is null || value is ExcelEmpty || value is ExcelMissing || (value is string s && string.IsNullOrWhiteSpace(s));

    private static object? NormalizeStringCell(object? value)
    {
        if (IsBlankExcelCell(value) || value is ExcelError) return null;
        return value switch
        {
            DateTime dt => dt.ToString("o", CultureInfo.InvariantCulture),
            double d => (double.IsNaN(d) || double.IsInfinity(d)) ? null : d.ToString("G17", CultureInfo.InvariantCulture),
            float f => (float.IsNaN(f) || float.IsInfinity(f)) ? null : ((double)f).ToString("G17", CultureInfo.InvariantCulture),
            decimal dec => dec.ToString(CultureInfo.InvariantCulture),
            _ => value.ToString()
        };
    }

    private static bool TryCreateNumericSetPayload(string name, object value, out Dictionary<string, object?>? payload, out string? transferFile)
    {
        payload = null; transferFile = null;
        if (!TryFlattenNumericExcelValue(value, out double[]? values, out int rows, out int cols)) return false;
        if (values is null || rows <= 0 || cols <= 0) return false;
        string dir = Path.Combine(Path.GetTempPath(), "PythonExcelBridgeTransfer");
        Directory.CreateDirectory(dir);
        transferFile = Path.Combine(dir, $"excel_numeric_{Guid.NewGuid():N}.bin");
        byte[] bytes = new byte[checked(values.Length * sizeof(double))];
        Buffer.BlockCopy(values, 0, bytes, 0, bytes.Length);
        File.WriteAllBytes(transferFile, bytes);
        payload = new Dictionary<string, object?> { ["id"] = Guid.NewGuid().ToString(), ["cmd"] = "set_numeric", ["name"] = name, ["file"] = transferFile.Replace('\\', '/'), ["rows"] = rows, ["cols"] = cols, ["storage_order"] = "row-major-double64", ["na"] = "NaN" };
        return true;
    }

    private static bool TryFlattenNumericExcelValue(object value, out double[]? values, out int rows, out int cols)
    {
        values = null; rows = 0; cols = 0;
        if (value is null || value is ExcelMissing || value is ExcelEmpty || value is ExcelError) return false;
        if (TryConvertToDoubleOrNaN(value, out double scalar, allowEmptyAsNaN: false)) { values = new[] { scalar }; rows = 1; cols = 1; return true; }
        if (value is object[] vector)
        {
            if (vector.Length == 0) return false;
            var tmp = new double[vector.Length]; bool sawAnyNumeric = false;
            for (int i = 0; i < vector.Length; i++)
            {
                if (!TryConvertToDoubleOrNaN(vector[i], out tmp[i], allowEmptyAsNaN: true)) return false;
                if (!double.IsNaN(tmp[i])) sawAnyNumeric = true;
            }
            if (!sawAnyNumeric) return false;
            values = tmp; rows = vector.Length; cols = 1; return true;
        }
        if (value is object[,] range)
        {
            int rowMin = range.GetLowerBound(0), rowMax = range.GetUpperBound(0), colMin = range.GetLowerBound(1), colMax = range.GetUpperBound(1);
            rows = rowMax - rowMin + 1; cols = colMax - colMin + 1;
            if (rows <= 0 || cols <= 0) return false;
            var tmp = new double[checked(rows * cols)]; bool sawAnyNumeric = false; int k = 0;
            for (int r = rowMin; r <= rowMax; r++)
            {
                for (int c = colMin; c <= colMax; c++)
                {
                    if (!TryConvertToDoubleOrNaN(range[r, c], out double d, allowEmptyAsNaN: true)) return false;
                    tmp[k++] = d;
                    if (!double.IsNaN(d)) sawAnyNumeric = true;
                }
            }
            if (!sawAnyNumeric) return false;
            values = tmp; return true;
        }
        return false;
    }

    private static bool TryConvertToDoubleOrNaN(object? value, out double result, bool allowEmptyAsNaN)
    {
        result = double.NaN;
        if (value is null || value is ExcelEmpty || value is ExcelMissing) return allowEmptyAsNaN;
        if (value is ExcelError) return false;
        switch (value)
        {
            case double d: result = (double.IsInfinity(d) || double.IsNaN(d)) ? double.NaN : d; return true;
            case float f: result = (float.IsInfinity(f) || float.IsNaN(f)) ? double.NaN : f; return true;
            case int i: result = i; return true;
            case long l: result = l; return true;
            case short sh: result = sh; return true;
            case decimal dec: result = (double)dec; return true;
            case DateTime dt: result = dt.ToOADate(); return true;
            default: return false;
        }
    }

    private static object[] NormalizeArgs(IEnumerable<object?> args)
    {
        var clean = new List<object>();
        foreach (object? arg in args)
        {
            if (arg is null || arg is ExcelMissing || arg is ExcelEmpty) continue;
            clean.Add(NormalizeExcelValue(arg));
        }
        return clean.ToArray();
    }

    private static object NormalizeExcelValue(object value)
    {
        if (value is object[] vector)
        {
            var output = new object?[vector.Length];
            for (int i = 0; i < vector.Length; i++) output[i] = NormalizeScalar(vector[i]);
            return output;
        }
        if (value is object[,] range)
        {
            int rowMin = range.GetLowerBound(0), rowMax = range.GetUpperBound(0), colMin = range.GetLowerBound(1), colMax = range.GetUpperBound(1);
            int rows = rowMax - rowMin + 1, cols = colMax - colMin + 1;
            if (rows == 1 && cols == 1) return NormalizeScalar(range[rowMin, colMin]) ?? "";
            if (cols == 1)
            {
                var output = new object?[rows];
                for (int r = 0; r < rows; r++) output[r] = NormalizeScalar(range[rowMin + r, colMin]);
                return output;
            }
            if (rows == 1)
            {
                var output = new object?[cols];
                for (int c = 0; c < cols; c++) output[c] = NormalizeScalar(range[rowMin, colMin + c]);
                return output;
            }
            var matrix = new object?[rows][];
            for (int r = 0; r < rows; r++)
            {
                matrix[r] = new object?[cols];
                for (int c = 0; c < cols; c++) matrix[r][c] = NormalizeScalar(range[rowMin + r, colMin + c]);
            }
            return matrix;
        }
        return NormalizeScalar(value) ?? "";
    }

    private static object? NormalizeScalar(object? value)
    {
        if (value is null || value is ExcelEmpty || value is ExcelMissing || value is ExcelError) return null;
        return value switch
        {
            double d => (double.IsNaN(d) || double.IsInfinity(d)) ? null : d,
            float f => (float.IsNaN(f) || float.IsInfinity(f)) ? null : (double)f,
            int i => i,
            long l => l,
            short s => (int)s,
            decimal m => m,
            bool b => b,
            string s => NormalizeStringArgument(s),
            DateTime dt => dt.ToString("o", CultureInfo.InvariantCulture),
            _ => value.ToString()
        };
    }

    private static object NormalizeStringArgument(string value)
    {
        if (value.StartsWith(PythonObjectReferencePrefix, StringComparison.Ordinal))
        {
            string name = value.Substring(PythonObjectReferencePrefix.Length).Trim();
            return new Dictionary<string, object?> { ["__pyexcel_arg_type"] = "pyobj", ["name"] = name };
        }
        return value;
    }

    private static object ConvertJsonToExcel(JsonElement el)
    {
        return el.ValueKind switch
        {
            JsonValueKind.Null => ExcelEmpty.Value,
            JsonValueKind.Number => el.TryGetInt64(out long l) ? l : el.GetDouble(),
            JsonValueKind.String => el.GetString() ?? "",
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.Array => ConvertJsonArrayToExcel(el),
            JsonValueKind.Object => ConvertJsonObjectToExcel(el),
            _ => el.ToString()
        };
    }

    private static object ConvertJsonObjectToExcel(JsonElement el)
    {
        if (el.TryGetProperty("__pyexcel_transfer_type", out JsonElement typeEl))
        {
            string? type = typeEl.GetString();
            if (string.Equals(type, "numeric", StringComparison.OrdinalIgnoreCase)) return ConvertNumericTransferToExcel(el);
            if (string.Equals(type, "table", StringComparison.OrdinalIgnoreCase)) return ConvertTableTransferToExcel(el);
        }
        return el.ToString();
    }

    private static object ConvertNumericTransferToExcel(JsonElement el)
    {
        string file = el.GetProperty("file").GetString() ?? string.Empty;
        int rows = el.GetProperty("rows").GetInt32();
        int cols = el.GetProperty("cols").GetInt32();
        if (rows < 0 || cols < 0) return "Error: invalid numeric transfer dimensions.";
        double[] values = ReadDoubleFile(file, checked(rows * cols));
        var output = new object[rows, cols];
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
            {
                double v = values[r * cols + c];
                output[r, c] = double.IsNaN(v) ? ExcelEmpty.Value : v;
            }
        TryDeleteFile(file);
        return output;
    }

    private static object ConvertTableTransferToExcel(JsonElement el)
    {
        int rows = el.GetProperty("rows").GetInt32();
        int cols = el.GetProperty("cols").GetInt32();
        bool includeHeaders = !el.TryGetProperty("include_headers", out JsonElement includeEl) || includeEl.GetBoolean();
        int rowOffset = includeHeaders ? 1 : 0;
        var output = new object[rows + rowOffset, cols];
        JsonElement columnsEl = el.GetProperty("columns");
        for (int c = 0; c < cols; c++)
        {
            JsonElement colEl = columnsEl[c];
            string name = colEl.GetProperty("name").GetString() ?? string.Empty;
            string type = colEl.GetProperty("type").GetString() ?? "character";
            if (includeHeaders) output[0, c] = name;
            if (string.Equals(type, "numeric", StringComparison.OrdinalIgnoreCase))
            {
                string file = colEl.GetProperty("file").GetString() ?? string.Empty;
                double[] values = ReadDoubleFile(file, rows);
                for (int r = 0; r < rows; r++) output[r + rowOffset, c] = double.IsNaN(values[r]) ? ExcelEmpty.Value : values[r];
                TryDeleteFile(file);
            }
            else
            {
                JsonElement valuesEl = colEl.GetProperty("values");
                for (int r = 0; r < rows; r++) output[r + rowOffset, c] = ConvertJsonToExcel(valuesEl[r]);
            }
        }
        return output;
    }

    private static double[] ReadDoubleFile(string file, int expectedCount)
    {
        if (string.IsNullOrWhiteSpace(file) || !File.Exists(file)) throw new FileNotFoundException("Python transfer file was not found.", file);
        long expectedBytes = (long)expectedCount * sizeof(double);
        var info = new FileInfo(file);
        if (info.Length < expectedBytes) throw new InvalidDataException($"Python transfer file is too short. Expected {expectedBytes} bytes, found {info.Length}.");
        byte[] bytes = File.ReadAllBytes(file);
        if (bytes.Length < expectedBytes) throw new InvalidDataException($"Python transfer file is too short. Expected {expectedBytes} bytes, read {bytes.Length}.");
        double[] values = new double[expectedCount];
        Buffer.BlockCopy(bytes, 0, values, 0, checked(expectedCount * sizeof(double)));
        return values;
    }

    private static void TryDeleteFile(string file)
    {
        try { if (!string.IsNullOrWhiteSpace(file) && File.Exists(file)) File.Delete(file); } catch { }
    }

    private static object ConvertJsonArrayToExcel(JsonElement el)
    {
        if (el.GetArrayLength() == 0) return ExcelEmpty.Value;
        bool allRows = true; int? cols = null;
        foreach (JsonElement row in el.EnumerateArray())
        {
            if (row.ValueKind != JsonValueKind.Array) { allRows = false; break; }
            int n = row.GetArrayLength();
            cols ??= n;
            if (cols.Value != n) { allRows = false; break; }
        }
        if (allRows && cols.HasValue) return ConvertJsonMatrixToExcel(el, el.GetArrayLength(), cols.Value);
        var arr = new object[el.GetArrayLength(), 1];
        for (int r = 0; r < el.GetArrayLength(); r++) arr[r, 0] = ConvertJsonToExcel(el[r]);
        return arr;
    }

    private static object ConvertJsonMatrixToExcel(JsonElement el, int rows, int cols)
    {
        var outArr = new object[rows, cols];
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++) outArr[r, c] = ConvertJsonToExcel(el[r][c]);
        return outArr;
    }
}
'''
(p/'PythonBridge.cs').write_text(cs, encoding='utf-8')
