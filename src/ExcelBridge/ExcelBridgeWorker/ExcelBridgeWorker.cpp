#include <windows.h>

#include <iostream>
#include <sstream>
#include <string>
#include <unordered_map>
#include <vector>

namespace
{
    std::unordered_map<std::string, std::string> TextObjects;
    std::unordered_map<std::string, std::string> MatrixObjects;
    std::unordered_map<std::string, std::vector<double>> NumericMatrixValues;
    std::unordered_map<std::string, std::pair<int,int>> NumericMatrixDims;

    std::string FullPipeName(const std::string& pipeName)
    {
        return "\\\\.\\pipe\\" + pipeName;
    }

    std::string TrimLineEndings(std::string value)
    {
        while (!value.empty() && (value.back() == '\r' || value.back() == '\n'))
            value.pop_back();
        return value;
    }

    bool StartsWith(const std::string& value, const std::string& prefix)
    {
        return value.rfind(prefix, 0) == 0;
    }

    std::vector<std::string> SplitTabs(const std::string& value)
    {
        std::vector<std::string> parts;
        std::stringstream ss(value);
        std::string item;

        while (std::getline(ss, item, '\t'))
            parts.push_back(item);

        return parts;
    }

    std::string HandleCommand(const std::string& command, bool& shouldStop)
    {
        if (command == "STATUS")
            return "STATUS\tOK\n";

        if (command == "STOP")
        {
            shouldStop = true;
            return "OK: worker stopped.\n";
        }

        if (StartsWith(command, "PING"))
        {
            std::string payload;
            if (command.size() > 4 && command[4] == '\t')
                payload = command.substr(5);

            if (payload.empty())
                return "OK: pipe is working.\n";

            return "OK: " + payload + ".\n";
        }

        auto parts = SplitTabs(command);

        if (!parts.empty() && parts[0] == "STORE_TEXT")
        {
            if (parts.size() < 3)
                return "ERROR: STORE_TEXT requires name and value.\n";

            TextObjects[parts[1]] = parts[2];
            return "OK: stored object '" + parts[1] + "'.\n";
        }

        if (!parts.empty() && parts[0] == "GET_TEXT")
        {
            if (parts.size() < 2)
                return "ERROR: GET_TEXT requires object name.\n";

            auto it = TextObjects.find(parts[1]);
            if (it == TextObjects.end())
                return "ERROR: object not found.\n";

            return "OK: " + it->second + "\n";
        }


        if (!parts.empty() && parts[0] == "STORE_MATRIX")
        {
            if (parts.size() < 5)
                return "ERROR: STORE_MATRIX requires name, rows, cols, and data.\n";

            MatrixObjects[parts[1]] = parts[2] + "\t" + parts[3] + "\t" + parts[4];
            return "OK: stored matrix '" + parts[1] + "'.\n";
        }

        if (!parts.empty() && parts[0] == "GET_MATRIX")
        {
            if (parts.size() < 2)
                return "ERROR: GET_MATRIX requires object name.\n";

            auto it = MatrixObjects.find(parts[1]);
            if (it == MatrixObjects.end())
                return "ERROR: matrix not found.\n";

            return "OKMATRIX\t" + it->second + "\n";
        }

        
        if (!parts.empty() && parts[0] == "STORE_NUMERIC_MATRIX")
        {
            if (parts.size() < 5)
                return "ERROR: STORE_NUMERIC_MATRIX requires name, rows, cols, and data.\n";

            int rows = std::stoi(parts[2]);
            int cols = std::stoi(parts[3]);

            std::vector<double> values;
            std::stringstream ss(parts[4]);
            std::string item;

            while (std::getline(ss, item, ','))
            {
                values.push_back(std::stod(item));
            }

            NumericMatrixValues[parts[1]] = values;
            NumericMatrixDims[parts[1]] = { rows, cols };

            return "OK: stored numeric matrix '" + parts[1] + "'.\n";
        }

        if (!parts.empty() && parts[0] == "GET_NUMERIC_MATRIX")
        {
            if (parts.size() < 2)
                return "ERROR: GET_NUMERIC_MATRIX requires object name.\n";

            auto vit = NumericMatrixValues.find(parts[1]);
            auto dit = NumericMatrixDims.find(parts[1]);

            if (vit == NumericMatrixValues.end() || dit == NumericMatrixDims.end())
                return "ERROR: numeric matrix not found.\n";

            std::string payload = std::to_string(dit->second.first) + "\t" + std::to_string(dit->second.second) + "\t";

            for (size_t i = 0; i < vit->second.size(); ++i)
            {
                if (i > 0)
                    payload += ",";

                payload += std::to_string(vit->second[i]);
            }

            return "OKNUMERICMATRIX\t" + payload + "\n";
        }



        if (!parts.empty() && parts[0] == "MATRIX_INFO")
        {
            if (parts.size() < 2)
                return "ERROR: MATRIX_INFO requires object name.\n";

            auto dit = NumericMatrixDims.find(parts[1]);
            if (dit == NumericMatrixDims.end())
                return "ERROR: numeric matrix not found.\n";

            return "OK: " + parts[1] + " = [" + std::to_string(dit->second.first) + "x" + std::to_string(dit->second.second) + "] numeric matrix.\n";
        }

        if (!parts.empty() && parts[0] == "MATRIX_TRANSPOSE")
        {
            if (parts.size() < 3)
                return "ERROR: MATRIX_TRANSPOSE requires target and source names.\n";

            auto vit = NumericMatrixValues.find(parts[2]);
            auto dit = NumericMatrixDims.find(parts[2]);

            if (vit == NumericMatrixValues.end() || dit == NumericMatrixDims.end())
                return "ERROR: source numeric matrix not found.\n";

            int rows = dit->second.first;
            int cols = dit->second.second;

            std::vector<double> out(rows * cols);

            for (int r = 0; r < rows; ++r)
            {
                for (int c = 0; c < cols; ++c)
                {
                    out[c * rows + r] = vit->second[r * cols + c];
                }
            }

            NumericMatrixValues[parts[1]] = out;
            NumericMatrixDims[parts[1]] = { cols, rows };

            return "OK: transposed matrix '" + parts[2] + "' into '" + parts[1] + "'.\n";
        }

        if (!parts.empty() && parts[0] == "MATRIX_MULTIPLY")
        {
            if (parts.size() < 4)
                return "ERROR: MATRIX_MULTIPLY requires target, left, and right names.\n";

            auto av = NumericMatrixValues.find(parts[2]);
            auto ad = NumericMatrixDims.find(parts[2]);
            auto bv = NumericMatrixValues.find(parts[3]);
            auto bd = NumericMatrixDims.find(parts[3]);

            if (av == NumericMatrixValues.end() || ad == NumericMatrixDims.end())
                return "ERROR: left matrix not found.\n";

            if (bv == NumericMatrixValues.end() || bd == NumericMatrixDims.end())
                return "ERROR: right matrix not found.\n";

            int aRows = ad->second.first;
            int aCols = ad->second.second;
            int bRows = bd->second.first;
            int bCols = bd->second.second;

            if (aCols != bRows)
                return "ERROR: incompatible matrix dimensions.\n";

            std::vector<double> out(aRows * bCols, 0.0);

            for (int r = 0; r < aRows; ++r)
            {
                for (int c = 0; c < bCols; ++c)
                {
                    double sum = 0.0;

                    for (int k = 0; k < aCols; ++k)
                    {
                        sum += av->second[r * aCols + k] * bv->second[k * bCols + c];
                    }

                    out[r * bCols + c] = sum;
                }
            }

            NumericMatrixValues[parts[1]] = out;
            NumericMatrixDims[parts[1]] = { aRows, bCols };

            return "OK: multiplied '" + parts[2] + "' x '" + parts[3] + "' into '" + parts[1] + "'.\n";
        }

        if (!parts.empty() && parts[0] == "REMOVE_OBJECT")
        {
            if (parts.size() < 2)
                return "ERROR: REMOVE_OBJECT requires object name.\n";

            size_t removed = TextObjects.erase(parts[1]);
            removed += MatrixObjects.erase(parts[1]);
            removed += NumericMatrixValues.erase(parts[1]);
            NumericMatrixDims.erase(parts[1]);

            if (removed == 0)
                return "ERROR: object not found: '" + parts[1] + "'.\n";

            return "OK: removed object '" + parts[1] + "'.\n";
        }

        if (command == "CLEAR_OBJECTS")
        {
            TextObjects.clear();
            MatrixObjects.clear();
            NumericMatrixValues.clear();
            NumericMatrixDims.clear();
            return "OK: cleared all objects.\n";
        }

        if (command == "OBJECTS")
        {
            if (TextObjects.empty() && MatrixObjects.empty() && NumericMatrixValues.empty())
                return "OK: no objects stored.\n";

            std::string result = "OK: objects = ";

            bool first = true;
            for (const auto& kv : TextObjects)
            {
                if (!first)
                    result += ", ";

                result += kv.first + " [text]";
                first = false;
            }

            for (const auto& kv : MatrixObjects)
            {
                if (!first)
                    result += ", ";

                std::vector<std::string> matrixParts = SplitTabs(kv.second);
                if (matrixParts.size() >= 2)
                    result += kv.first + " [matrix " + matrixParts[0] + "x" + matrixParts[1] + "]";
                else
                    result += kv.first + " [matrix]";

                first = false;
            }

            for (const auto& kv : NumericMatrixDims)
            {
                if (!first)
                    result += ", ";
                result += kv.first + " [numeric matrix " + std::to_string(kv.second.first) + "x" + std::to_string(kv.second.second) + "]";
                first = false;
            }

            result += ".\n";
            return result;
        }

        return "ERROR: unknown command.\n";
    }
}

int main(int argc, char* argv[])
{
    if (argc < 2 || argv[1] == nullptr || argv[1][0] == '\0')
    {
        std::cerr << "Usage: ExcelBridgeWorker.exe <pipe-name>" << std::endl;
        return 1;
    }

    const std::string pipePath = FullPipeName(argv[1]);
    bool shouldStop = false;

    while (!shouldStop)
    {
        HANDLE pipe = CreateNamedPipeA(
            pipePath.c_str(),
            PIPE_ACCESS_DUPLEX,
            PIPE_TYPE_MESSAGE | PIPE_READMODE_MESSAGE | PIPE_WAIT,
            PIPE_UNLIMITED_INSTANCES,
            4096,
            4096,
            0,
            nullptr);

        if (pipe == INVALID_HANDLE_VALUE)
            return 2;

        BOOL connected = ConnectNamedPipe(pipe, nullptr) ? TRUE : (GetLastError() == ERROR_PIPE_CONNECTED);
        if (!connected)
        {
            CloseHandle(pipe);
            continue;
        }

        char buffer[4096] = {};
        DWORD bytesRead = 0;
        BOOL ok = ReadFile(pipe, buffer, sizeof(buffer) - 1, &bytesRead, nullptr);

        if (ok && bytesRead > 0)
        {
            buffer[bytesRead] = '\0';

            std::string command = TrimLineEndings(buffer);
            std::string response = HandleCommand(command, shouldStop);

            DWORD bytesWritten = 0;
            WriteFile(pipe, response.c_str(), static_cast<DWORD>(response.size()), &bytesWritten, nullptr);
            FlushFileBuffers(pipe);
        }

        DisconnectNamedPipe(pipe);
        CloseHandle(pipe);
    }

    return 0;
}
