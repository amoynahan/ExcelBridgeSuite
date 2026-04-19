using ExcelDna.Integration;
using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace JuliaExcelBridge;

public static class JBridge
{
    private static Process? _proc;
    private static StreamWriter? _stdin;
    private static StreamReader? _stdout;
    private static StreamReader? _stderr;
    private static readonly object _lock = new();

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

    private static string WorkerPath => Path.Combine(AddInDir, "worker.jl");
    private static string StartupPath => Path.Combine(AddInDir, "startup.jl");
    private static string LogPath => Path.Combine(AddInDir, "julia_worker_stderr.log");
    private static string StartupErrorLogPath => Path.Combine(AddInDir, "addin_startup_error.log");
    private static string JuliaConfigPath => Path.Combine(AddInDir, "julia-path.txt");
    private static string PlotConfigPath => Path.Combine(AddInDir, "plot-path.txt");

    public static void TryStart()
    {
        try
        {
            Start();
        }
        catch (Exception ex)
        {
            SafeLogStartupError(ex);
        }
    }

    public static void Start()
    {
        lock (_lock)
        {
            if (_proc is { HasExited: false })
                return;

            string workerPath = WorkerPath;
            string startupPath = StartupPath;

            if (!File.Exists(workerPath))
                throw new FileNotFoundException("worker.jl was not found.", workerPath);

            if (!File.Exists(startupPath))
                throw new FileNotFoundException("startup.jl was not found.", startupPath);

            string juliaPath = GetJuliaPath();

            var psi = new ProcessStartInfo
            {
                FileName = juliaPath,
                Arguments = $"--startup-file=no --color=no \"{workerPath}\" \"{startupPath}\"",
                UseShellExecute = false,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                WorkingDirectory = AddInDir
            };

            _proc = Process.Start(psi)
                ?? throw new InvalidOperationException("Failed to start the Julia worker process.");

            _stdin = _proc.StandardInput;
            _stdout = _proc.StandardOutput;
            _stderr = _proc.StandardError;

            BeginErrorCapture();

            object pingResult = Ping();
            if (pingResult is string s &&
                (s.StartsWith("Julia error:", StringComparison.OrdinalIgnoreCase) ||
                 s.Contains("no response", StringComparison.OrdinalIgnoreCase)))
            {
                throw new InvalidOperationException($"Julia worker failed startup ping: {s}");
            }
        }
    }

    public static void Stop()
    {
        lock (_lock)
        {
            try
            {
                _stdin?.Close();

                if (_proc is { HasExited: false })
                    _proc.Kill(entireProcessTree: true);
            }
            catch
            {
            }
            finally
            {
                _stdin?.Dispose();
                _stdout?.Dispose();
                _stderr?.Dispose();
                _proc?.Dispose();

                _stdin = null;
                _stdout = null;
                _stderr = null;
                _proc = null;
            }
        }
    }

    public static object Ping()
    {
        return Send(new Dictionary<string, object?>
        {
            ["id"] = Guid.NewGuid().ToString(),
            ["cmd"] = "ping"
        });
    }

    public static object Source(string file)
    {
        return Send(new Dictionary<string, object?>
        {
            ["id"] = Guid.NewGuid().ToString(),
            ["cmd"] = "source",
            ["file"] = file
        });
    }

    public static object Eval(string code)
    {
        return Send(new Dictionary<string, object?>
        {
            ["id"] = Guid.NewGuid().ToString(),
            ["cmd"] = "eval",
            ["code"] = code
        });
    }

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
        return Send(new Dictionary<string, object?>
        {
            ["id"] = Guid.NewGuid().ToString(),
            ["cmd"] = "set",
            ["name"] = name,
            ["value"] = NormalizeExcelValue(value)
        });
    }

    public static object Get(string name)
    {
        return Send(new Dictionary<string, object?>
        {
            ["id"] = Guid.NewGuid().ToString(),
            ["cmd"] = "get",
            ["name"] = name
        });
    }

    public static object Exists(string name)
    {
        return Send(new Dictionary<string, object?>
        {
            ["id"] = Guid.NewGuid().ToString(),
            ["cmd"] = "exists",
            ["name"] = name
        });
    }

    public static object Remove(string name)
    {
        return Send(new Dictionary<string, object?>
        {
            ["id"] = Guid.NewGuid().ToString(),
            ["cmd"] = "remove",
            ["name"] = name
        });
    }

    public static object Objects()
    {
        return Send(new Dictionary<string, object?>
        {
            ["id"] = Guid.NewGuid().ToString(),
            ["cmd"] = "objects"
        });
    }

    public static object Describe(string name)
    {
        return Send(new Dictionary<string, object?>
        {
            ["id"] = Guid.NewGuid().ToString(),
            ["cmd"] = "describe",
            ["name"] = name
        });
    }

    public static object Plot(string code, string? plotName = null, int width = 800, int height = 600, int res = 96)
    {
        string file = BuildPlotPath(plotName);

        return Send(new Dictionary<string, object?>
        {
            ["id"] = Guid.NewGuid().ToString(),
            ["cmd"] = "plot",
            ["code"] = code,
            ["file"] = file,
            ["width"] = width,
            ["height"] = height,
            ["res"] = res
        });
    }

    public static object InsertPlotFromSelectedCell()
    {
        try
        {
            dynamic? app = ExcelDnaUtil.Application;
            dynamic? target = app?.ActiveCell;

            if (target is null)
                return "Error: no active cell was found.";

            object rawValue = target.Value2;
            string? file = rawValue?.ToString();

            if (string.IsNullOrWhiteSpace(file))
                return "Error: selected cell is blank. Select a cell containing the PNG path returned by JPlot.";

            return InsertPlotFile(file);
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    public static object InsertPlotFile(string file)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(file))
                return "Error: plot file path is blank.";

            string fullPath = Path.GetFullPath(Environment.ExpandEnvironmentVariables(file.Trim().Trim('"')));

            if (!File.Exists(fullPath))
                return $"Error: plot file was not found: {fullPath}";

            dynamic? app = ExcelDnaUtil.Application;
            dynamic? sheet = app?.ActiveSheet;
            dynamic? target = app?.ActiveCell;

            if (sheet is null || target is null)
                return "Error: could not resolve the active worksheet or cell.";

            string cellAddress = "Cell";
            string sheetName = "Sheet";

            try
            {
                cellAddress = target.Address[false, false]?.ToString() ?? "Cell";
            }
            catch
            {
            }

            try
            {
                sheetName = sheet.Name?.ToString() ?? "Sheet";
            }
            catch
            {
            }

            string shapeName = $"JPlot_{SanitizeFileComponent(sheetName)}_{SanitizeFileComponent(cellAddress)}";

            try
            {
                dynamic oldShape = sheet.Shapes.Item(shapeName);
                oldShape.Delete();
            }
            catch
            {
            }

            float left = Convert.ToSingle(target.Left);
            float top = Convert.ToSingle(target.Top + target.Height + 6);

            dynamic shape = sheet.Shapes.AddPicture(fullPath, 0, -1, left, top, -1, -1);
            shape.Name = shapeName;

            try
            {
                shape.LockAspectRatio = -1;
            }
            catch
            {
            }

            return fullPath;
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    public static string GetPlotDirectory()
    {
        string defaultDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "JuliaExcelBridge",
            "PlotCache");

        if (File.Exists(PlotConfigPath))
        {
            string configured = File.ReadAllText(PlotConfigPath).Trim();

            if (!string.IsNullOrWhiteSpace(configured))
            {
                try
                {
                    string candidate = ExpandConfiguredPath(configured);
                    if (Directory.Exists(candidate))
                        return candidate;
                }
                catch
                {
                }
            }
        }

        Directory.CreateDirectory(defaultDir);
        return defaultDir;
    }

    private static string ExpandConfiguredPath(string configured)
    {
        string expanded = Environment.ExpandEnvironmentVariables(configured.Trim().Trim('"'));
        string fullPath = Path.GetFullPath(expanded);
        return fullPath;
    }

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
        try
        {
            dynamic app = ExcelDnaUtil.Application;
            string? name = app?.ActiveWorkbook?.Name as string;
            if (!string.IsNullOrWhiteSpace(name))
                return Path.GetFileNameWithoutExtension(name);
        }
        catch
        {
        }

        return "Workbook";
    }

    private static string GetCallerSheetName()
    {
        try
        {
            dynamic app = ExcelDnaUtil.Application;
            string? name = app?.ActiveSheet?.Name as string;
            if (!string.IsNullOrWhiteSpace(name))
                return name;
        }
        catch
        {
        }

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
        catch
        {
        }

        return "Cell";
    }

    private static string ColumnNumberToLetters(int columnNumber)
    {
        var letters = string.Empty;
        int col = columnNumber;
        while (col > 0)
        {
            int rem = (col - 1) % 26;
            letters = (char)('A' + rem) + letters;
            col = (col - 1) / 26;
        }
        return letters;
    }

    private static string SanitizeFileComponent(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return "item";

        string cleaned = Regex.Replace(value, @"[^A-Za-z0-9._-]+", "_");
        cleaned = cleaned.Trim('_', '.', ' ');
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

            if (string.IsNullOrWhiteSpace(line))
                return "Julia worker returned no response.";

            using JsonDocument doc = JsonDocument.Parse(line);
            JsonElement root = doc.RootElement;

            bool ok = root.TryGetProperty("ok", out JsonElement okEl) && okEl.GetBoolean();
            if (!ok)
            {
                if (root.TryGetProperty("error", out JsonElement errEl))
                    return $"Julia error: {errEl.GetString()}";

                return "Unknown Julia error.";
            }

            if (!root.TryGetProperty("result", out JsonElement resultEl))
                return ExcelEmpty.Value;

            return ConvertJsonToExcel(resultEl);
        }
    }

    private static void EnsureRunning()
    {
        if (_proc is null || _proc.HasExited)
            Start();
    }

    private static void BeginErrorCapture()
    {
        string path = LogPath;

        Task.Run(() =>
        {
            try
            {
                using var sw = new StreamWriter(path, append: true);

                while (_stderr is not null && !_stderr.EndOfStream)
                {
                    string? line = _stderr.ReadLine();
                    if (!string.IsNullOrWhiteSpace(line))
                    {
                        sw.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | {line}");
                        sw.Flush();
                    }
                }
            }
            catch
            {
            }
        });
    }

    private static void SafeLogStartupError(Exception ex)
    {
        try
        {
            File.AppendAllText(
                StartupErrorLogPath,
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {ex}{Environment.NewLine}{Environment.NewLine}");
        }
        catch
        {
        }
    }

    
private static string GetJuliaPath()
    {
        if (File.Exists(JuliaConfigPath))
        {
            string configured = File.ReadAllText(JuliaConfigPath).Trim();

            if (!string.IsNullOrWhiteSpace(configured) && File.Exists(configured))
                return configured;
        }

        string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        string[] candidates =
        {
            Path.Combine(localAppData, "Microsoft", "WindowsApps", "julia.exe"),
            Path.Combine(AddInDir, "Julia", "bin", "julia.exe"),
            Path.Combine(AddInDir, "runtime", "Julia", "bin", "julia.exe")
        };

        foreach (string candidate in candidates)
        {
            if (File.Exists(candidate))
                return candidate;
        }

        string[] programRoots =
        {
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)
        };

        foreach (string root in programRoots.Where(r => !string.IsNullOrWhiteSpace(r)))
        {
            foreach (string candidateDir in Directory.EnumerateDirectories(root, "Julia-*"))
            {
                string candidate = Path.Combine(candidateDir, "bin", "julia.exe");
                if (File.Exists(candidate))
                    return candidate;
            }
        }

        return "julia";
    }

    private static object[] NormalizeArgs(IEnumerable<object?> args)
    {
        var clean = new List<object>();

        foreach (object? arg in args)
        {
            if (arg is null || arg is ExcelMissing || arg is ExcelEmpty)
                continue;

            clean.Add(NormalizeExcelValue(arg));
        }

        return clean.ToArray();
    }

    private static object NormalizeExcelValue(object value)
    {
        if (value is object[,] range)
        {
            int rowMin = range.GetLowerBound(0);
            int rowMax = range.GetUpperBound(0);
            int colMin = range.GetLowerBound(1);
            int colMax = range.GetUpperBound(1);
            int rows = rowMax - rowMin + 1;
            int cols = colMax - colMin + 1;

            if (rows == 1 && cols == 1)
                return NormalizeScalar(range[rowMin, colMin]) ?? "";

            if (cols == 1)
            {
                var output = new object?[rows];
                for (int r = 0; r < rows; r++)
                    output[r] = NormalizeScalar(range[rowMin + r, colMin]);
                return output;
            }

            if (rows == 1)
            {
                var output = new object?[cols];
                for (int c = 0; c < cols; c++)
                    output[c] = NormalizeScalar(range[rowMin, colMin + c]);
                return output;
            }

            var matrix = new object?[rows][];
            for (int r = 0; r < rows; r++)
            {
                matrix[r] = new object?[cols];
                for (int c = 0; c < cols; c++)
                    matrix[r][c] = NormalizeScalar(range[rowMin + r, colMin + c]);
            }

            return matrix;
        }

        return NormalizeScalar(value) ?? "";
    }

    private static object? NormalizeScalar(object? value)
    {
        if (value is null || value is ExcelEmpty || value is ExcelMissing)
            return null;

        if (value is ExcelError)
            return null;

        return value switch
        {
            double d => d,
            float f => (double)f,
            int i => i,
            long l => l,
            short s => (int)s,
            bool b => b,
            string s => s,
            DateTime dt => dt.ToString("o", CultureInfo.InvariantCulture),
            _ => value.ToString()
        };
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
            JsonValueKind.Object => el.ToString(),
            _ => el.ToString()
        };
    }

    private static object ConvertJsonArrayToExcel(JsonElement el)
    {
        int rows = el.GetArrayLength();
        if (rows == 0)
            return ExcelEmpty.Value;

        bool firstIsArray = el[0].ValueKind == JsonValueKind.Array;

        if (!firstIsArray)
        {
            var arr = new object[rows, 1];
            for (int r = 0; r < rows; r++)
                arr[r, 0] = ConvertJsonToExcel(el[r]);
            return arr;
        }

        int cols = el[0].GetArrayLength();
        var outArr = new object[rows, cols];

        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                outArr[r, c] = ConvertJsonToExcel(el[r][c]);
            }
        }

        return outArr;
    }
}