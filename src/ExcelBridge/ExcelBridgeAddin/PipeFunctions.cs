using ExcelDna.Integration;
using System;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Reflection;
using System.Text;

public static class PipeFunctions
{
    private static readonly object SyncRoot = new();
    private static Process? WorkerProcess;
    private static string PipeName => "ExcelBridgePipe_" + Process.GetCurrentProcess().Id.ToString();

    private static string AddinDirectory
    {
        get
        {
            // Excel-DNA knows the loaded XLL path. This keeps the worker lookup
            // anchored to the same folder as the add-in, including the publish folder.
            string? xllPath = ExcelDnaUtil.XllPath;
            if (!string.IsNullOrWhiteSpace(xllPath))
            {
                string? xllDir = Path.GetDirectoryName(xllPath);
                if (!string.IsNullOrWhiteSpace(xllDir))
                    return xllDir;
            }

            string? location = Assembly.GetExecutingAssembly().Location;
            if (!string.IsNullOrWhiteSpace(location))
            {
                string? dir = Path.GetDirectoryName(location);
                if (!string.IsNullOrWhiteSpace(dir))
                    return dir;
            }

            return AppContext.BaseDirectory;
        }
    }

    private static string WorkerPath => Path.Combine(AddinDirectory, "ExcelBridgeWorker.exe");

    private static bool IsWorkerRunning()
    {
        lock (SyncRoot)
        {
            return WorkerProcess != null && !WorkerProcess.HasExited;
        }
    }

    private static string SendCommand(string command, int timeoutMs = 3000)
    {
        using NamedPipeClientStream client = new(".", PipeName, PipeDirection.InOut, PipeOptions.None);
        client.Connect(timeoutMs);

        using StreamReader reader = new(client, Encoding.UTF8, detectEncodingFromByteOrderMarks: false, bufferSize: 4096, leaveOpen: true);
        using StreamWriter writer = new(client, new UTF8Encoding(false), bufferSize: 4096, leaveOpen: true)
        {
            AutoFlush = true
        };

        writer.WriteLine(command);
        return reader.ReadLine() ?? string.Empty;
    }

    private static string EnsureWorkerStarted()
    {
        lock (SyncRoot)
        {
            if (WorkerProcess != null && !WorkerProcess.HasExited)
                return "Already running";

            if (!File.Exists(WorkerPath))
                return "ExcelBridgeWorker.exe not found beside the add-in: " + WorkerPath;

            ProcessStartInfo psi = new()
            {
                FileName = WorkerPath,
                Arguments = PipeName,
                UseShellExecute = false,
                CreateNoWindow = true,
                WorkingDirectory = AddinDirectory
            };

            WorkerProcess = Process.Start(psi);
            if (WorkerProcess == null)
                return "Failed to start ExcelBridgeWorker.exe.";
        }

        // Wait briefly until the pipe is accepting connections.
        DateTime deadline = DateTime.UtcNow.AddSeconds(5);
        Exception? lastError = null;
        while (DateTime.UtcNow < deadline)
        {
            try
            {
                string response = SendCommand("STATUS", 500);
                if (response.StartsWith("STATUS\tOK", StringComparison.OrdinalIgnoreCase))
                    return "OK: started worker. Pipe=" + PipeName + ".";
            }
            catch (Exception ex)
            {
                lastError = ex;
                System.Threading.Thread.Sleep(100);
            }
        }

        return "Worker process started, but the pipe did not respond. Last error: " + (lastError?.Message ?? "none");
    }

    [ExcelFunction(Name = "CPP_PIPE_START", Description = "Starts the external ExcelBridge worker process and opens a named-pipe endpoint.")]
    public static string CPP_PIPE_START()
    {
        try
        {
            return EnsureWorkerStarted();
        }
        catch (Exception ex)
        {
            return "CPP_PIPE_START failed: " + ex.Message;
        }
    }

    [ExcelFunction(Name = "CPP_PIPE_PING", Description = "Sends a test message to the external ExcelBridge worker over a named pipe.")]
    public static string CPP_PIPE_PING(string message = "hello")
    {
        try
        {
            string start = EnsureWorkerStarted();
            if (start.StartsWith("ExcelBridgeWorker.exe not found", StringComparison.OrdinalIgnoreCase) ||
                start.StartsWith("Failed", StringComparison.OrdinalIgnoreCase))
                return start;

            return SendCommand("PING\t" + (message ?? string.Empty));
        }
        catch (Exception ex)
        {
            return "CPP_PIPE_PING failed: " + ex.Message;
        }
    }

    [ExcelFunction(Name = "CPP_PIPE_STATUS", Description = "Returns the add-in's current named-pipe worker status.")]
    public static string CPP_PIPE_STATUS()
    {
        try
        {
            string processStatus = IsWorkerRunning() ? "tracked worker process is running" : "no tracked worker process";
            return "Pipe=" + PipeName + "; WorkerPath=" + WorkerPath + "; Status=" + processStatus;
        }
        catch (Exception ex)
        {
            return "CPP_PIPE_STATUS failed: " + ex.Message;
        }
    }

    [ExcelFunction(Name = "CPP_PIPE_STOP", Description = "Stops the external ExcelBridge worker process.")]
    public static string CPP_PIPE_STOP()
    {
        try
        {
            if (!IsWorkerRunning())
                return "No tracked worker process is running.";

            string response;
            try
            {
                response = SendCommand("STOP", 1000);
            }
            catch (Exception ex)
            {
                response = "STOP command failed: " + ex.Message;
            }

            lock (SyncRoot)
            {
                if (WorkerProcess != null && !WorkerProcess.HasExited)
                    WorkerProcess.WaitForExit(2000);

                if (WorkerProcess != null && !WorkerProcess.HasExited)
                    WorkerProcess.Kill(entireProcessTree: true);

                WorkerProcess?.Dispose();
                WorkerProcess = null;
            }

            return response;
        }
        catch (Exception ex)
        {
            return "CPP_PIPE_STOP failed: " + ex.Message;
        }
    }
    [ExcelFunction(Name = "CPP_STORE_TEXT", Description = "Stores a persistent text object in the worker.")]
    public static string CPP_STORE_TEXT(string name, string value)
    {
        try
        {
            EnsureWorkerStarted();
            return SendCommand("STORE_TEXT\t" + name + "\t" + value);
        }
        catch (Exception ex)
        {
            return "CPP_STORE_TEXT failed: " + ex.Message;
        }
    }

    [ExcelFunction(Name = "CPP_GET_TEXT", Description = "Gets a persistent text object from the worker.")]
    public static string CPP_GET_TEXT(string name)
    {
        try
        {
            EnsureWorkerStarted();
            return SendCommand("GET_TEXT\t" + name);
        }
        catch (Exception ex)
        {
            return "CPP_GET_TEXT failed: " + ex.Message;
        }
    }

    [ExcelFunction(Name = "CPP_OBJECTS", Description = "Lists persistent objects in the worker.")]
    public static string CPP_OBJECTS()
    {
        try
        {
            EnsureWorkerStarted();
            return SendCommand("OBJECTS");
        }
        catch (Exception ex)
        {
            return "CPP_OBJECTS failed: " + ex.Message;
        }
    }



    [ExcelFunction(Name = "CPP_PIPE_OBJECTS", Description = "Lists persistent objects in the worker. Alias for CPP_OBJECTS.")]
    public static string CPP_PIPE_OBJECTS()
    {
        return CPP_OBJECTS();
    }

    [ExcelFunction(Name = "CPP_REMOVE_OBJECT", Description = "Removes a persistent object from the worker.")]
    public static string CPP_REMOVE_OBJECT(string name)
    {
        try
        {
            EnsureWorkerStarted();
            return SendCommand("REMOVE_OBJECT\t" + name);
        }
        catch (Exception ex)
        {
            return "CPP_REMOVE_OBJECT failed: " + ex.Message;
        }
    }

    [ExcelFunction(Name = "CPP_CLEAR_OBJECTS", Description = "Clears all persistent objects from the worker.")]
    public static string CPP_CLEAR_OBJECTS()
    {
        try
        {
            EnsureWorkerStarted();
            return SendCommand("CLEAR_OBJECTS");
        }
        catch (Exception ex)
        {
            return "CPP_CLEAR_OBJECTS failed: " + ex.Message;
        }
    }


    [ExcelFunction(Name = "CPP_STORE_MATRIX", Description = "Stores a numeric matrix in the worker.")]
    public static string CPP_STORE_MATRIX(string name, object[,] values)
    {
        try
        {
            EnsureWorkerStarted();

            int rows = values.GetLength(0);
            int cols = values.GetLength(1);

            StringBuilder sb = new();

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    if (r > 0 || c > 0)
                        sb.Append(",");

                    sb.Append(Convert.ToString(values[r, c]) ?? "");
                }
            }

            return SendCommand($"STORE_MATRIX\t{name}\t{rows}\t{cols}\t{sb}");
        }
        catch (Exception ex)
        {
            return "CPP_STORE_MATRIX failed: " + ex.Message;
        }
    }

    [ExcelFunction(Name = "CPP_GET_MATRIX", Description = "Returns a stored numeric matrix.")]
    public static object[,] CPP_GET_MATRIX(string name)
    {
        try
        {
            EnsureWorkerStarted();

            string response = SendCommand("GET_MATRIX\t" + name);

            if (!response.StartsWith("OKMATRIX\t"))
                return new object[,] { { response } };

            string payload = response.Substring("OKMATRIX\t".Length);
            string[] parts = payload.Split('\t');

            int rows = int.Parse(parts[0]);
            int cols = int.Parse(parts[1]);

            string[] values = parts[2].Split(',');

            object[,] result = new object[rows, cols];

            int index = 0;
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    result[r, c] = values[index++];
                }
            }

            return result;
        }
        catch (Exception ex)
        {
            return new object[,] { { "CPP_GET_MATRIX failed: " + ex.Message } };
        }
    }

    [ExcelFunction(Name = "CPP_STORE_NUMERIC_MATRIX", Description = "Stores a numeric matrix using numeric serialization.")]
    public static string CPP_STORE_NUMERIC_MATRIX(string name, object[,] values)
    {
        try
        {
            EnsureWorkerStarted();

            int rows = values.GetLength(0);
            int cols = values.GetLength(1);

            StringBuilder sb = new();

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    if (r > 0 || c > 0)
                        sb.Append(",");

                    double d = Convert.ToDouble(values[r, c]);
                    sb.Append(d.ToString(System.Globalization.CultureInfo.InvariantCulture));
                }
            }

            return SendCommand($"STORE_NUMERIC_MATRIX\t{name}\t{rows}\t{cols}\t{sb}");
        }
        catch (Exception ex)
        {
            return "CPP_STORE_NUMERIC_MATRIX failed: " + ex.Message;
        }
    }

    [ExcelFunction(Name = "CPP_GET_NUMERIC_MATRIX", Description = "Returns a stored numeric matrix.")]
    public static object[,] CPP_GET_NUMERIC_MATRIX(string name)
    {
        try
        {
            EnsureWorkerStarted();

            string response = SendCommand("GET_NUMERIC_MATRIX\t" + name);

            if (!response.StartsWith("OKNUMERICMATRIX\t"))
                return new object[,] { { response } };

            string payload = response.Substring("OKNUMERICMATRIX\t".Length);
            string[] parts = payload.Split('\t');

            int rows = int.Parse(parts[0]);
            int cols = int.Parse(parts[1]);

            string[] values = parts[2].Split(',');

            object[,] result = new object[rows, cols];

            int index = 0;
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    result[r, c] = double.Parse(values[index++], System.Globalization.CultureInfo.InvariantCulture);
                }
            }

            return result;
        }
        catch (Exception ex)
        {
            return new object[,] { { "CPP_GET_NUMERIC_MATRIX failed: " + ex.Message } };
        }
    }


    [ExcelFunction(Name = "CPP_MATRIX_INFO", Description = "Returns information about a numeric matrix.")]
    public static string CPP_MATRIX_INFO(string name)
    {
        try
        {
            EnsureWorkerStarted();
            return SendCommand("MATRIX_INFO\t" + name);
        }
        catch (Exception ex)
        {
            return "CPP_MATRIX_INFO failed: " + ex.Message;
        }
    }

    [ExcelFunction(Name = "CPP_MATRIX_TRANSPOSE", Description = "Creates a transposed numeric matrix.")]
    public static string CPP_MATRIX_TRANSPOSE(string targetName, string sourceName)
    {
        try
        {
            EnsureWorkerStarted();
            return SendCommand("MATRIX_TRANSPOSE\t" + targetName + "\t" + sourceName);
        }
        catch (Exception ex)
        {
            return "CPP_MATRIX_TRANSPOSE failed: " + ex.Message;
        }
    }

    [ExcelFunction(Name = "CPP_MATRIX_MULTIPLY", Description = "Multiplies two numeric matrices.")]
    public static string CPP_MATRIX_MULTIPLY(string targetName, string leftName, string rightName)
    {
        try
        {
            EnsureWorkerStarted();
            return SendCommand("MATRIX_MULTIPLY\t" + targetName + "\t" + leftName + "\t" + rightName);
        }
        catch (Exception ex)
        {
            return "CPP_MATRIX_MULTIPLY failed: " + ex.Message;
        }
    }


}
