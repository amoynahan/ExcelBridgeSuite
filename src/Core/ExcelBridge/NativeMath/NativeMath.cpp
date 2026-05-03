#include <cstring>

#define EXPORT extern "C" __declspec(dllexport)

EXPORT int __cdecl MatrixRoundTrip(
    const double* input,
    int rows,
    int cols,
    double* output)
{
    if (!input || !output || rows <= 0 || cols <= 0)
        return 1;

    const size_t n = static_cast<size_t>(rows) * static_cast<size_t>(cols);

    std::memcpy(output, input, n * sizeof(double));

    return 0;
}
