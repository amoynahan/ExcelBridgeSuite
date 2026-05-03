using ExcelDna.Integration;
using System;
using System.Runtime.InteropServices;

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

    private static object[,] RunMatrixRoundTrip(double[] flat, int rows, int cols)
    {
        double[] output = new double[flat.Length];

        try
        {
            int rc = MatrixRoundTrip(flat, rows, cols, output);
            if (rc != 0)
                return new object[,] { { $"Native error: {rc}" } };
        }
        catch (DllNotFoundException)
        {
            return new object[,] { { "NativeMath.dll not found. Rebuild solution and confirm it is copied next to ExcelBridgeAddin.dll." } };
        }
        catch (BadImageFormatException)
        {
            return new object[,] { { "NativeMath.dll architecture mismatch. Build both projects as x64." } };
        }
        catch (EntryPointNotFoundException)
        {
            return new object[,] { { "Native function MatrixRoundTrip was not found in NativeMath.dll." } };
        }

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

    // Safe/default version.
    // Excel users type: =MatrixRoundTrip(...)
    [ExcelFunction(
        Name = "MatrixRoundTrip",
        Description = "Safe version. Passes a numeric Excel range to native C++ and returns the result."
    )]
    public static object[,] MatrixRoundTripExcel(object[,] input)
    {
        if (input == null)
            return new object[,] { { ExcelError.ExcelErrorNull } };

        int rows = input.GetLength(0);
        int cols = input.GetLength(1);

        if (rows == 0 || cols == 0)
            return new object[,] { { ExcelError.ExcelErrorValue } };

        double[] flat = new double[rows * cols];

        try
        {
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    object value = input[r, c];

                    if (value == null || value is ExcelEmpty || value is ExcelMissing)
                        return new object[,] { { ExcelError.ExcelErrorValue } };

                    if (value is double d)
                        flat[r * cols + c] = d;
                    else
                        flat[r * cols + c] = Convert.ToDouble(value);
                }
            }

            return RunMatrixRoundTrip(flat, rows, cols);
        }
        catch
        {
            return new object[,] { { ExcelError.ExcelErrorValue } };
        }
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
            return new object[,] { { ExcelError.ExcelErrorNull } };

        int rows = input.GetLength(0);
        int cols = input.GetLength(1);

        if (rows == 0 || cols == 0)
            return new object[,] { { ExcelError.ExcelErrorValue } };

        double[] flat = new double[rows * cols];

        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                flat[r * cols + c] = input[r, c];
            }
        }

        return RunMatrixRoundTrip(flat, rows, cols);
    }
}
