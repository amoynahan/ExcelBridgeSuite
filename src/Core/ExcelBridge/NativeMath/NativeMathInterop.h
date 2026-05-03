#pragma once
#include <Eigen/Dense>

namespace NativeMathInterop
{
    using MatRM =
        Eigen::Matrix<double, Eigen::Dynamic, Eigen::Dynamic, Eigen::RowMajor>;

    inline Eigen::Map<const MatRM> AsMatConst(const double* data, int rows, int cols)
    {
        return Eigen::Map<const MatRM>(data, rows, cols);
    }

    inline Eigen::Map<MatRM> AsMat(double* data, int rows, int cols)
    {
        return Eigen::Map<MatRM>(data, rows, cols);
    }
}
