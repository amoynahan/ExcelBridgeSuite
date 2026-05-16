#include <cstring>
#include <string>
#include <unordered_map>
#include <memory>
#include <mutex>
#include <sstream>
#include <atomic>

#include <Eigen/Dense>
#include <Eigen/Cholesky>

#define EXPORT extern "C" __declspec(dllexport)

namespace
{
    using MatRM = Eigen::Matrix<double, Eigen::Dynamic, Eigen::Dynamic, Eigen::RowMajor>;

    struct MatrixEntry
    {
        MatRM value;
    };

    struct LLTEntry
    {
        Eigen::LLT<MatRM> llt;
        int n = 0;
    };

    std::unordered_map<std::string, std::shared_ptr<MatrixEntry>> g_matrices;
    std::unordered_map<std::string, std::shared_ptr<LLTEntry>> g_llts;
    std::mutex g_storeMutex;
    std::atomic<unsigned long long> g_nextLltId{ 1 };

    bool IsBlank(const char* s)
    {
        return s == nullptr || *s == '\0';
    }

    int CopyStringToBuffer(const std::string& value, char* buffer, int capacity)
    {
        if (!buffer || capacity <= 0)
            return 20;

        if (static_cast<int>(value.size()) >= capacity)
        {
            std::memcpy(buffer, value.c_str(), static_cast<size_t>(capacity - 1));
            buffer[capacity - 1] = '\0';
            return 21;
        }

        std::memcpy(buffer, value.c_str(), value.size() + 1);
        return 0;
    }
}

// Existing round-trip test. Copies a row-major matrix from input to output.
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

// Stores a matrix in native C++ memory under a user-supplied name.
// The matrix persists until Excel unloads the add-in/DLL or CppClearStore is called.
EXPORT int __cdecl CppMatrixStore(
    const char* name,
    const double* input,
    int rows,
    int cols)
{
    if (IsBlank(name) || !input || rows <= 0 || cols <= 0)
        return 1;

    try
    {
        Eigen::Map<const MatRM> mapped(input, rows, cols);

        auto entry = std::make_shared<MatrixEntry>();
        entry->value = mapped;

        std::lock_guard<std::mutex> lock(g_storeMutex);
        g_matrices[std::string(name)] = entry;

        return 0;
    }
    catch (...)
    {
        return 99;
    }
}

// Creates and stores a persistent Eigen LLT Cholesky decomposition from a stored matrix.
// Returns a handle such as LLT_1 in outHandle.
EXPORT int __cdecl CppLLTCreate(
    const char* matrixName,
    char* outHandle,
    int outHandleCapacity)
{
    if (IsBlank(matrixName) || !outHandle || outHandleCapacity <= 0)
        return 1;

    try
    {
        std::shared_ptr<MatrixEntry> matrixEntry;
        {
            std::lock_guard<std::mutex> lock(g_storeMutex);
            auto it = g_matrices.find(std::string(matrixName));
            if (it == g_matrices.end())
                return 2;
            matrixEntry = it->second;
        }

        const MatRM& A = matrixEntry->value;
        if (A.rows() <= 0 || A.cols() <= 0 || A.rows() != A.cols())
            return 3;

        Eigen::LLT<MatRM> llt(A);
        if (llt.info() != Eigen::Success)
            return 4;

        auto entry = std::make_shared<LLTEntry>();
        entry->llt = std::move(llt);
        entry->n = static_cast<int>(A.rows());

        std::ostringstream handle;
        handle << "LLT_" << g_nextLltId.fetch_add(1);
        const std::string handleText = handle.str();

        {
            std::lock_guard<std::mutex> lock(g_storeMutex);
            g_llts[handleText] = entry;
        }

        return CopyStringToBuffer(handleText, outHandle, outHandleCapacity);
    }
    catch (...)
    {
        return 99;
    }
}

// Returns the dimension n of a stored LLT decomposition for an n-by-n source matrix.
EXPORT int __cdecl CppLLTDim(
    const char* handle,
    int* n)
{
    if (IsBlank(handle) || !n)
        return 1;

    std::lock_guard<std::mutex> lock(g_storeMutex);
    auto it = g_llts.find(std::string(handle));
    if (it == g_llts.end())
        return 2;

    *n = it->second->n;
    return 0;
}

// Returns the lower triangular Cholesky factor L.
EXPORT int __cdecl CppLLTGetL(
    const char* handle,
    double* output,
    int rows,
    int cols)
{
    if (IsBlank(handle) || !output || rows <= 0 || cols <= 0 || rows != cols)
        return 1;

    try
    {
        std::shared_ptr<LLTEntry> entry;
        {
            std::lock_guard<std::mutex> lock(g_storeMutex);
            auto it = g_llts.find(std::string(handle));
            if (it == g_llts.end())
                return 2;
            entry = it->second;
        }

        if (rows != entry->n || cols != entry->n)
            return 3;

        MatRM L = entry->llt.matrixL();
        std::memcpy(output, L.data(), static_cast<size_t>(rows) * static_cast<size_t>(cols) * sizeof(double));
        return 0;
    }
    catch (...)
    {
        return 99;
    }
}

// Returns the upper triangular Cholesky factor U.
EXPORT int __cdecl CppLLTGetU(
    const char* handle,
    double* output,
    int rows,
    int cols)
{
    if (IsBlank(handle) || !output || rows <= 0 || cols <= 0 || rows != cols)
        return 1;

    try
    {
        std::shared_ptr<LLTEntry> entry;
        {
            std::lock_guard<std::mutex> lock(g_storeMutex);
            auto it = g_llts.find(std::string(handle));
            if (it == g_llts.end())
                return 2;
            entry = it->second;
        }

        if (rows != entry->n || cols != entry->n)
            return 3;

        MatRM U = entry->llt.matrixU();
        std::memcpy(output, U.data(), static_cast<size_t>(rows) * static_cast<size_t>(cols) * sizeof(double));
        return 0;
    }
    catch (...)
    {
        return 99;
    }
}

// Solves A * X = B using a stored LLT decomposition of A.
EXPORT int __cdecl CppLLTSolve(
    const char* handle,
    const double* rhs,
    int rhsRows,
    int rhsCols,
    double* output)
{
    if (IsBlank(handle) || !rhs || !output || rhsRows <= 0 || rhsCols <= 0)
        return 1;

    try
    {
        std::shared_ptr<LLTEntry> entry;
        {
            std::lock_guard<std::mutex> lock(g_storeMutex);
            auto it = g_llts.find(std::string(handle));
            if (it == g_llts.end())
                return 2;
            entry = it->second;
        }

        if (rhsRows != entry->n)
            return 3;

        Eigen::Map<const MatRM> B(rhs, rhsRows, rhsCols);
        MatRM X = entry->llt.solve(B);

        if (entry->llt.info() != Eigen::Success)
            return 4;

        std::memcpy(output, X.data(), static_cast<size_t>(rhsRows) * static_cast<size_t>(rhsCols) * sizeof(double));
        return 0;
    }
    catch (...)
    {
        return 99;
    }
}

// Returns Eigen's estimated reciprocal condition number for the stored LLT.
EXPORT int __cdecl CppLLTRCond(
    const char* handle,
    double* value)
{
    if (IsBlank(handle) || !value)
        return 1;

    try
    {
        std::shared_ptr<LLTEntry> entry;
        {
            std::lock_guard<std::mutex> lock(g_storeMutex);
            auto it = g_llts.find(std::string(handle));
            if (it == g_llts.end())
                return 2;
            entry = it->second;
        }

        *value = entry->llt.rcond();
        return 0;
    }
    catch (...)
    {
        return 99;
    }
}

// Reconstructs A from the stored LLT decomposition.
EXPORT int __cdecl CppLLTReconstruct(
    const char* handle,
    double* output,
    int rows,
    int cols)
{
    if (IsBlank(handle) || !output || rows <= 0 || cols <= 0 || rows != cols)
        return 1;

    try
    {
        std::shared_ptr<LLTEntry> entry;
        {
            std::lock_guard<std::mutex> lock(g_storeMutex);
            auto it = g_llts.find(std::string(handle));
            if (it == g_llts.end())
                return 2;
            entry = it->second;
        }

        if (rows != entry->n || cols != entry->n)
            return 3;

        MatRM A = entry->llt.reconstructedMatrix();
        std::memcpy(output, A.data(), static_cast<size_t>(rows) * static_cast<size_t>(cols) * sizeof(double));
        return 0;
    }
    catch (...)
    {
        return 99;
    }
}

// Lists persistent native objects as tab-delimited text: Type<TAB>Name<TAB>Rows<TAB>Cols.
EXPORT int __cdecl CppListObjects(
    char* buffer,
    int capacity)
{
    try
    {
        std::ostringstream ss;
        ss << "Type\tName\tRows\tCols\n";

        std::lock_guard<std::mutex> lock(g_storeMutex);

        for (const auto& kv : g_matrices)
        {
            ss << "Matrix\t" << kv.first << "\t"
                << kv.second->value.rows() << "\t"
                << kv.second->value.cols() << "\n";
        }

        for (const auto& kv : g_llts)
        {
            ss << "LLT\t" << kv.first << "\t"
                << kv.second->n << "\t"
                << kv.second->n << "\n";
        }

        return CopyStringToBuffer(ss.str(), buffer, capacity);
    }
    catch (...)
    {
        return 99;
    }
}

// Clears all persistent native objects.
EXPORT int __cdecl CppClearStore()
{
    try
    {
        std::lock_guard<std::mutex> lock(g_storeMutex);
        g_matrices.clear();
        g_llts.clear();
        return 0;
    }
    catch (...)
    {
        return 99;
    }
}

// Retrieves a stored matrix by name.
EXPORT int __cdecl CppMatrixGet(
    const char* name,
    double* output,
    int rows,
    int cols)
{
    if (IsBlank(name) || !output || rows <= 0 || cols <= 0)
        return 1;

    try
    {
        std::shared_ptr<MatrixEntry> entry;
        {
            std::lock_guard<std::mutex> lock(g_storeMutex);
            auto it = g_matrices.find(std::string(name));
            if (it == g_matrices.end())
                return 2;
            entry = it->second;
        }

        if (entry->value.rows() != rows || entry->value.cols() != cols)
            return 3;

        std::memcpy(output, entry->value.data(),
            static_cast<size_t>(rows) * static_cast<size_t>(cols) * sizeof(double));

        return 0;
    }
    catch (...)
    {
        return 99;
    }
}

EXPORT int __cdecl CppMatrixDim(
    const char* name,
    int* rows,
    int* cols)
{
    if (IsBlank(name) || !rows || !cols)
        return 1;

    std::lock_guard<std::mutex> lock(g_storeMutex);
    auto it = g_matrices.find(std::string(name));
    if (it == g_matrices.end())
        return 2;

    *rows = static_cast<int>(it->second->value.rows());
    *cols = static_cast<int>(it->second->value.cols());
    return 0;
}

EXPORT int __cdecl CppDeleteObject(const char* name)
{
    if (IsBlank(name))
        return 1;

    std::lock_guard<std::mutex> lock(g_storeMutex);

    if (g_matrices.erase(std::string(name)) > 0)
        return 0;

    if (g_llts.erase(std::string(name)) > 0)
        return 0;

    return 2;
}

EXPORT int __cdecl CppLLTCreateAs(
    const char* handleName,
    const char* matrixName)
{
    if (IsBlank(handleName) || IsBlank(matrixName))
        return 1;

    try
    {
        std::shared_ptr<MatrixEntry> matrixEntry;
        {
            std::lock_guard<std::mutex> lock(g_storeMutex);
            auto it = g_matrices.find(std::string(matrixName));
            if (it == g_matrices.end())
                return 2;
            matrixEntry = it->second;
        }

        const MatRM& A = matrixEntry->value;

        Eigen::LLT<MatRM> llt(A);
        if (llt.info() != Eigen::Success)
            return 4;

        auto entry = std::make_shared<LLTEntry>();
        entry->llt = std::move(llt);
        entry->n = static_cast<int>(A.rows());

        {
            std::lock_guard<std::mutex> lock(g_storeMutex);
            g_llts[std::string(handleName)] = entry;
        }

        return 0;
    }
    catch (...)
    {
        return 99;
    }

}

// Detailed listing of persistent native objects as tab-delimited text.
// This is an alias-style function for Excel-facing CPP_OBJECTS_DETAIL().
EXPORT int __cdecl CppObjectsDetail(
    char* buffer,
    int capacity)
{
    try
    {
        std::ostringstream ss;
        ss << "Handle\tType\tRows\tCols\n";

        std::lock_guard<std::mutex> lock(g_storeMutex);

        for (const auto& kv : g_matrices)
        {
            ss << kv.first << "\tMatrix\t"
                << kv.second->value.rows() << "\t"
                << kv.second->value.cols() << "\n";
        }

        for (const auto& kv : g_llts)
        {
            ss << kv.first << "\tLLT\t"
                << kv.second->n << "\t"
                << kv.second->n << "\n";
        }

        return CopyStringToBuffer(ss.str(), buffer, capacity);
    }
    catch (...)
    {
        return 99;
    }
}