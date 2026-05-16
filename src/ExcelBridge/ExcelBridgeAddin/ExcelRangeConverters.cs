using System;
using ExcelDna.Integration;

internal static class ExcelRangeConverters
{
    // Excel range object[,] -> row-major double[] (reusable)
    internal static double[] RangeToRowMajor(object[,] range, out int rows, out int cols)
    {
        rows = range.GetLength(0);
        cols = range.GetLength(1);

        var flat = new double[rows * cols];

        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
            {
                object v = range[r, c];
                if (v == null || v is ExcelEmpty)
                    throw new ArgumentException("Range contains empty cells.");

                flat[r * cols + c] = Convert.ToDouble(v); // row-major
            }

        return flat;
    }

    // row-major double[] -> double[,] (reusable for UDF return/spill)
    internal static double[,] RowMajorToDoubleArray(double[] flat, int rows, int cols)
    {
        var arr = new double[rows, cols];
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
                arr[r, c] = flat[r * cols + c];

        return arr;
    }
}
