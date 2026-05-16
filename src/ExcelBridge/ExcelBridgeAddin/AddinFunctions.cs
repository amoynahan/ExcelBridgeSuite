using ExcelDna.Integration;
using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;

public static class AddinFunctions
{
    // Native C++ function exported from NativeMath.dll.
    // This name must match the exported C++ function name.
    [DllImport("NativeMath.dll", CallingConvention = CallingConvention.Cdecl)]
    private static extern int MatrixRoundTrip(
        double[] input,
        int rows,
        int cols,
        double[] output
    );

    [DllImport("NativeMath.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
    private static extern int CppMatrixStore(
        string name,
        double[] input,
        int rows,
        int cols
    );

    [DllImport("NativeMath.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
    private static extern int CppLLTCreate(
        string matrixName,
        StringBuilder outHandle,
        int outHandleCapacity
    );

    [DllImport("NativeMath.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
    private static extern int CppLLTDim(
        string handle,
        out int n
    );

    [DllImport("NativeMath.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
    private static extern int CppLLTGetL(
        string handle,
        double[] output,
        int rows,
        int cols
    );

    [DllImport("NativeMath.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
    private static extern int CppLLTGetU(
        string handle,
        double[] output,
        int rows,
        int cols
    );

    [DllImport("NativeMath.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
    private static extern int CppLLTSolve(
        string handle,
        double[] rhs,
        int rhsRows,
        int rhsCols,
        double[] output
    );

    [DllImport("NativeMath.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
    private static extern int CppLLTRCond(
        string handle,
        out double value
    );

    [DllImport("NativeMath.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
    private static extern int CppLLTReconstruct(
        string handle,
        double[] output,
        int rows,
        int cols
    );

    [DllImport("NativeMath.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
    private static extern int CppListObjects(
        StringBuilder buffer,
        int capacity
    );

    [DllImport("NativeMath.dll", CallingConvention = CallingConvention.Cdecl)]
    private static extern int CppClearStore();

    [DllImport("NativeMath.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
    private static extern int CppMatrixGet(
        string name,
        double[] output,
        int rows,
        int cols
    );

    [DllImport("NativeMath.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
    private static extern int CppMatrixDim(
        string name,
        out int rows,
        out int cols
    );

    [DllImport("NativeMath.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
    private static extern int CppDeleteObject(
        string name
    );

    [DllImport("NativeMath.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi, EntryPoint = "CppLLTCreateAs")]
    private static extern int CppLLTCreateAsNative(
        string handleName,
        string matrixName
    );



    private static object[,] Error(string message) => new object[,] { { message } };

    private static string NativeError(string functionName, int rc)
    {
        string detail = rc switch
        {
            1 => "invalid argument",
            2 => "object not found",
            3 => "dimension mismatch or non-square matrix",
            4 => "Eigen decomposition/solve failed; matrix may not be positive definite",
            20 => "invalid output buffer",
            21 => "output buffer too small",
            99 => "native C++ exception",
            _ => "unknown error"
        };

        return $"{functionName} failed: rc={rc} ({detail})";
    }

    private static bool TryToFlatDoubleArray(object[,] input, out double[] flat, out int rows, out int cols, out string error)
    {
        flat = Array.Empty<double>();
        rows = 0;
        cols = 0;
        error = string.Empty;

        if (input == null)
        {
            error = "Input range is null.";
            return false;
        }

        rows = input.GetLength(0);
        cols = input.GetLength(1);

        if (rows == 0 || cols == 0)
        {
            error = "Input range is empty.";
            return false;
        }

        flat = new double[rows * cols];

        try
        {
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    object value = input[r, c];

                    if (value == null || value is ExcelEmpty || value is ExcelMissing)
                    {
                        error = "Input range contains blank cells.";
                        return false;
                    }

                    if (value is double d)
                    {
                        flat[r * cols + c] = d;
                    }
                    else
                    {
                        flat[r * cols + c] = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                    }
                }
            }

            return true;
        }
        catch
        {
            error = "Input range must contain only numeric values.";
            return false;
        }
    }

    private static double[] ToFlatDoubleArray(double[,] input, out int rows, out int cols)
    {
        rows = input.GetLength(0);
        cols = input.GetLength(1);

        double[] flat = new double[rows * cols];
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                flat[r * cols + c] = input[r, c];
            }
        }

        return flat;
    }

    private static object[,] ToObjectMatrix(double[] flat, int rows, int cols)
    {
        object[,] result = new object[rows, cols];

        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                result[r, c] = flat[r * cols + c];
            }
        }

        return result;
    }

    private static object[,] RunMatrixRoundTrip(double[] flat, int rows, int cols)
    {
        double[] output = new double[flat.Length];

        try
        {
            int rc = MatrixRoundTrip(flat, rows, cols, output);
            if (rc != 0)
                return Error(NativeError("MatrixRoundTrip", rc));
        }
        catch (DllNotFoundException)
        {
            return Error("NativeMath.dll not found. Rebuild solution and confirm it is copied next to ExcelBridgeAddin.dll.");
        }
        catch (BadImageFormatException)
        {
            return Error("NativeMath.dll architecture mismatch. Build both projects as x64.");
        }
        catch (EntryPointNotFoundException)
        {
            return Error("Native function MatrixRoundTrip was not found in NativeMath.dll.");
        }

        return ToObjectMatrix(output, rows, cols);
    }

    // Safe/default version.
    // Excel users type: =MatrixRoundTrip(...)
    [ExcelFunction(
        Name = "MatrixRoundTrip",
        Description = "Safe version. Passes a numeric Excel range to native C++ and returns the result."
    )]
    public static object[,] MatrixRoundTripExcel(object[,] input)
    {
        if (!TryToFlatDoubleArray(input, out double[] flat, out int rows, out int cols, out string error))
            return Error(error);

        return RunMatrixRoundTrip(flat, rows, cols);
    }

    // Fast version.
    // Excel users type: =MatrixRoundTripFast(...)
    // Best for clean numeric-only matrices.
    [ExcelFunction(
        Name = "MatrixRoundTripFast",
        Description = "Fast version. Requires a clean numeric-only matrix."
    )]
    public static object[,] MatrixRoundTripFastExcel(double[,] input)
    {
        if (input == null)
            return Error("Input range is null.");

        int rows;
        int cols;
        double[] flat = ToFlatDoubleArray(input, out rows, out cols);

        if (rows == 0 || cols == 0)
            return Error("Input range is empty.");

        return RunMatrixRoundTrip(flat, rows, cols);
    }

    [ExcelFunction(
        Name = "CPP_MATRIX_STORE",
        Description = "Stores a numeric matrix in persistent native C++ memory under the supplied name."
    )]
    public static object CPP_MATRIX_STORE(string name, object[,] input)
    {
        if (string.IsNullOrWhiteSpace(name))
            return "Matrix name is required.";

        if (!TryToFlatDoubleArray(input, out double[] flat, out int rows, out int cols, out string error))
            return error;

        try
        {
            int rc = CppMatrixStore(name.Trim(), flat, rows, cols);
            if (rc != 0)
                return NativeError("CppMatrixStore", rc);

            return name.Trim();
        }
        catch (Exception ex) when (ex is DllNotFoundException || ex is BadImageFormatException || ex is EntryPointNotFoundException)
        {
            return ex.Message;
        }
    }

    [ExcelFunction(
        Name = "CPP_LLT_CREATE",
        Description = "Creates a persistent Eigen LLT Cholesky decomposition from a stored matrix and returns an LLT handle."
    )]
    public static object CPP_LLT_CREATE(string matrixName)
    {
        if (string.IsNullOrWhiteSpace(matrixName))
            return "Matrix name is required.";

        try
        {
            const int capacity = 256;
            StringBuilder handle = new(capacity);
            int rc = CppLLTCreate(matrixName.Trim(), handle, capacity);
            if (rc != 0)
                return NativeError("CppLLTCreate", rc);

            return handle.ToString();
        }
        catch (Exception ex) when (ex is DllNotFoundException || ex is BadImageFormatException || ex is EntryPointNotFoundException)
        {
            return ex.Message;
        }
    }

    private static bool TryGetLLTDim(string handle, out int n, out object[,] error)
    {
        n = 0;
        error = Error("Unknown error.");

        if (string.IsNullOrWhiteSpace(handle))
        {
            error = Error("LLT handle is required.");
            return false;
        }

        int rc = CppLLTDim(handle.Trim(), out n);
        if (rc != 0)
        {
            error = Error(NativeError("CppLLTDim", rc));
            return false;
        }

        return true;
    }

    [ExcelFunction(
        Name = "CPP_LLT_L",
        Description = "Returns the lower triangular L factor from a persistent Eigen LLT object."
    )]
    public static object[,] CPP_LLT_L(string handle)
    {
        try
        {
            if (!TryGetLLTDim(handle, out int n, out object[,] error))
                return error;

            double[] output = new double[n * n];
            int rc = CppLLTGetL(handle.Trim(), output, n, n);
            if (rc != 0)
                return Error(NativeError("CppLLTGetL", rc));

            return ToObjectMatrix(output, n, n);
        }
        catch (Exception ex) when (ex is DllNotFoundException || ex is BadImageFormatException || ex is EntryPointNotFoundException)
        {
            return Error(ex.Message);
        }
    }

    [ExcelFunction(
        Name = "CPP_LLT_U",
        Description = "Returns the upper triangular U factor from a persistent Eigen LLT object."
    )]
    public static object[,] CPP_LLT_U(string handle)
    {
        try
        {
            if (!TryGetLLTDim(handle, out int n, out object[,] error))
                return error;

            double[] output = new double[n * n];
            int rc = CppLLTGetU(handle.Trim(), output, n, n);
            if (rc != 0)
                return Error(NativeError("CppLLTGetU", rc));

            return ToObjectMatrix(output, n, n);
        }
        catch (Exception ex) when (ex is DllNotFoundException || ex is BadImageFormatException || ex is EntryPointNotFoundException)
        {
            return Error(ex.Message);
        }
    }

    [ExcelFunction(
        Name = "CPP_LLT_SOLVE",
        Description = "Solves A*X=B using a persistent Eigen LLT object created from A."
    )]



    public static object[,] CPP_LLT_SOLVE(string handle, object[,] rhs)
    {
        try
        {
            if (!TryGetLLTDim(handle, out int n, out object[,] dimError))
                return dimError;

            if (!TryToFlatDoubleArray(rhs, out double[] flatRhs, out int rhsRows, out int rhsCols, out string error))
                return Error(error);

            if (rhsRows != n)
                return Error($"Right-hand side row count must equal {n}.");

            double[] output = new double[rhsRows * rhsCols];
            int rc = CppLLTSolve(handle.Trim(), flatRhs, rhsRows, rhsCols, output);
            if (rc != 0)
                return Error(NativeError("CppLLTSolve", rc));

            return ToObjectMatrix(output, rhsRows, rhsCols);
        }
        catch (Exception ex) when (ex is DllNotFoundException || ex is BadImageFormatException || ex is EntryPointNotFoundException)
        {
            return Error(ex.Message);
        }
    }

    [ExcelFunction(
        Name = "CPP_LLT_RCOND",
        Description = "Returns Eigen's estimated reciprocal condition number for a persistent LLT object."
    )]
    public static object CPP_LLT_RCOND(string handle)
    {
        if (string.IsNullOrWhiteSpace(handle))
            return "LLT handle is required.";

        try
        {
            int rc = CppLLTRCond(handle.Trim(), out double value);
            if (rc != 0)
                return NativeError("CppLLTRCond", rc);

            return value;
        }
        catch (Exception ex) when (ex is DllNotFoundException || ex is BadImageFormatException || ex is EntryPointNotFoundException)
        {
            return ex.Message;
        }
    }

    [ExcelFunction(
        Name = "CPP_LLT_RECONSTRUCT",
        Description = "Reconstructs the original matrix from a persistent Eigen LLT object."
    )]
    public static object[,] CPP_LLT_RECONSTRUCT(string handle)
    {
        try
        {
            if (!TryGetLLTDim(handle, out int n, out object[,] error))
                return error;

            double[] output = new double[n * n];
            int rc = CppLLTReconstruct(handle.Trim(), output, n, n);
            if (rc != 0)
                return Error(NativeError("CppLLTReconstruct", rc));

            return ToObjectMatrix(output, n, n);
        }
        catch (Exception ex) when (ex is DllNotFoundException || ex is BadImageFormatException || ex is EntryPointNotFoundException)
        {
            return Error(ex.Message);
        }
    }

    [ExcelFunction(
        Name = "CPP_NATIVE_OBJECTS",
        Description = "Lists matrices and LLT decompositions currently stored in persistent native C++ memory."
    )]
    public static object[,] CPP_NATIVE_OBJECTS()
    {
        try
        {
            const int capacity = 32768;
            StringBuilder buffer = new(capacity);
            int rc = CppListObjects(buffer, capacity);
            if (rc != 0)
                return Error(NativeError("CppListObjects", rc));

            string[] lines = buffer.ToString().Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length == 0)
                return Error("No output returned.");

            object[,] result = new object[lines.Length, 4];
            for (int r = 0; r < lines.Length; r++)
            {
                string[] parts = lines[r].TrimEnd('\r').Split('\t');
                for (int c = 0; c < 4; c++)
                {
                    result[r, c] = c < parts.Length ? parts[c] : string.Empty;
                }
            }

            return result;
        }
        catch (Exception ex) when (ex is DllNotFoundException || ex is BadImageFormatException || ex is EntryPointNotFoundException)
        {
            return Error(ex.Message);
        }
    }



    [ExcelFunction(
        Name = "CPP_CLEAR",
        Description = "Clears all persistent native C++ matrix and LLT objects."
    )]
    public static object CPP_CLEAR()
    {
        try
        {
            int rc = CppClearStore();
            if (rc != 0)
                return NativeError("CppClearStore", rc);

            return "Persistent native object store cleared.";
        }
        catch (Exception ex) when (ex is DllNotFoundException || ex is BadImageFormatException || ex is EntryPointNotFoundException)
        {
            return ex.Message;
        }
    }

    [ExcelFunction(Name = "CPP_MATRIX_GET", Description = "Retrieve a stored persistent native C++ matrix.")]
    public static object CppMatrixGet(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            return Error("Matrix name is required.");

        int rc = CppMatrixDim(name.Trim(), out int rows, out int cols);
        if (rc != 0)
            return Error(NativeError("CppMatrixDim", rc));

        double[] output = new double[rows * cols];

        rc = CppMatrixGet(name.Trim(), output, rows, cols);
        if (rc != 0)
            return Error(NativeError("CppMatrixGet", rc));

        object[,] result = new object[rows, cols];

        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                result[r, c] = output[r * cols + c];
            }
        }

        return result;
    }

    [ExcelFunction(Name = "CPP_OBJECT_DELETE", Description = "Delete a stored persistent native C++ object by handle/name.")]
    public static string CppObjectDelete(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            return "Object name/handle is required.";

        int rc = CppDeleteObject(name.Trim());
        return rc == 0 ? "Deleted" : NativeError("CppObjectDelete", rc);
    }

    [ExcelFunction(Name = "CPP_LLT_CREATE_AS", Description = "Create a named persistent LLT/Cholesky object from a stored matrix.")]
    public static string CppLLTCreateAs(string handleName, string matrixName)
    {
        if (string.IsNullOrWhiteSpace(handleName))
            return "LLT handle name is required.";

        if (string.IsNullOrWhiteSpace(matrixName))
            return "Matrix name is required.";

        int rc = CppLLTCreateAsNative(handleName.Trim(), matrixName.Trim());
        return rc == 0 ? handleName.Trim() : NativeError("CppLLTCreateAs", rc);
    }

}