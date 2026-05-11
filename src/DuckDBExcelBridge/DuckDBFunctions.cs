using System;
using System.Globalization;
using System.IO;
using System.Text;
using ExcelDna.Integration;
using DuckDB.NET.Data;

namespace DuckDBExcelBridge
{
    public static class DuckDBFunctions
    {
        private static DuckDBConnection? _conn;
        private static string _dbPath = ":memory:";
        private static readonly object _lock = new();

        private static DuckDBConnection GetConnection()
        {
            lock (_lock)
            {
                if (_conn == null)
                {
                    _conn = new DuckDBConnection($"Data Source={_dbPath}");
                    _conn.Open();
                }

                return _conn;
            }
        }

        private static void ResetConnection()
        {
            lock (_lock)
            {
                try
                {
                    _conn?.Close();
                    _conn?.Dispose();
                }
                catch
                {
                    // Ignore cleanup errors
                }

                _conn = null;
            }
        }

        private static string FormatDuckDbError(DuckDBException ex)
        {
            // Do not reset the connection for ordinary SQL/DuckDB errors.
            // This preserves in-memory tables after bad SQL, bad casts, etc.
            return "DUCKDB ERROR | " + ex.Message;
        }

        private static string FormatFatalError(Exception ex)
        {
            // Use this for unexpected .NET/runtime errors only.
            // With :memory:, resetting the connection necessarily loses tables.
            return "ERROR | " + ex.Message;
        }

        private static string EscapeSqlString(string value)
        {
            return value.Replace("'", "''");
        }

        private static string QuoteIdentifier(string name)
        {
            return "\"" + name.Replace("\"", "\"\"") + "\"";
        }

        private static string ToSqlLiteral(object value)
        {
            if (value == null ||
                value is ExcelEmpty ||
                value is ExcelMissing ||
                value is ExcelError)
                return "NULL";

            if (value is string s)
                return $"'{EscapeSqlString(s)}'";

            if (value is DateTime dt)
                return $"TIMESTAMP '{dt:yyyy-MM-dd HH:mm:ss}'";

            if (value is bool b)
                return b ? "TRUE" : "FALSE";

            if (value is double d)
                return double.IsNaN(d) || double.IsInfinity(d)
                    ? "NULL"
                    : d.ToString(CultureInfo.InvariantCulture);

            if (value is float f)
                return float.IsNaN(f) || float.IsInfinity(f)
                    ? "NULL"
                    : f.ToString(CultureInfo.InvariantCulture);

            if (value is int or long or short or decimal)
                return Convert.ToString(value, CultureInfo.InvariantCulture) ?? "NULL";

            return $"'{EscapeSqlString(value.ToString() ?? "")}'";
        }

        private static string ToCsvField(object value)
        {
            if (value == null ||
                value is ExcelEmpty ||
                value is ExcelMissing ||
                value is ExcelError)
                return "";

            string text;

            if (value is DateTime dt)
                text = dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            else if (value is double d)
                text = d.ToString(CultureInfo.InvariantCulture);
            else if (value is float f)
                text = f.ToString(CultureInfo.InvariantCulture);
            else if (value is bool b)
                text = b ? "TRUE" : "FALSE";
            else
                text = Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";

            bool mustQuote =
                text.Contains(",") ||
                text.Contains("\"") ||
                text.Contains("\r") ||
                text.Contains("\n");

            if (mustQuote)
                text = "\"" + text.Replace("\"", "\"\"") + "\"";

            return text;
        }

        private static object[,] ReaderToArray(DuckDBDataReader reader, bool includeHeaders)
        {
            int cols = reader.FieldCount;
            var rows = new System.Collections.Generic.List<object[]>();

            if (includeHeaders)
            {
                var header = new object[cols];
                for (int c = 0; c < cols; c++)
                    header[c] = reader.GetName(c);
                rows.Add(header);
            }

            while (reader.Read())
            {
                var row = new object[cols];

                for (int c = 0; c < cols; c++)
                {
                    var value = reader.GetValue(c);
                    row[c] = value == DBNull.Value ? "" : value;
                }

                rows.Add(row);
            }

            var output = new object[rows.Count, cols];

            for (int r = 0; r < rows.Count; r++)
                for (int c = 0; c < cols; c++)
                    output[r, c] = rows[r][c];

            return output;
        }


        [ExcelFunction(Name = "DuckOpen", Description = "Opens a DuckDB database file, or :memory: for an in-memory database.")]
        public static string DuckOpen(string path)
        {
            try
            {
                lock (_lock)
                {
                    _conn?.Dispose();
                    _conn = null;

                    _dbPath = string.IsNullOrWhiteSpace(path) ? ":memory:" : path.Trim();
                    _conn = new DuckDBConnection($"Data Source={_dbPath}");
                    _conn.Open();
                }

                return $"OK | DuckDB opened {_dbPath}";
            }
            catch (DuckDBException ex)
            {
                return FormatDuckDbError(ex);
            }
            catch (Exception ex)
            {
                return FormatFatalError(ex);
            }
        }

        [ExcelFunction(Name = "DuckClose", Description = "Closes the current DuckDB connection.")]
        public static string DuckClose()
        {
            try
            {
                lock (_lock)
                {
                    _conn?.Dispose();
                    _conn = null;
                }

                return "OK | DuckDB closed";
            }
            catch (DuckDBException ex)
            {
                return FormatDuckDbError(ex);
            }
            catch (Exception ex)
            {
                return FormatFatalError(ex);
            }
        }

        [ExcelFunction(Name = "DuckDBPath", Description = "Returns the current DuckDB database path.")]
        public static string DuckDBPath()
        {
            return _dbPath;
        }

        [ExcelFunction(Name = "DuckSetTable", Description = "Creates or replaces a DuckDB table from an Excel range. Alias for DuckImportRangeBulk.")]
        public static object DuckSetTable(string tableName, object[,] data, bool firstRowHeaders)
        {
            return DuckImportRangeBulk(tableName, data, firstRowHeaders);
        }

        [ExcelFunction(Name = "DuckListTables", Description = "Lists DuckDB tables and views. Alias for DuckTables.")]
        public static object DuckListTables()
        {
            return DuckTables();
        }

        [ExcelFunction(Name = "DuckDropTable", Description = "Drops a DuckDB table or view if it exists. Alias for DuckDrop.")]
        public static object DuckDropTable(string tableName)
        {
            return DuckDrop(tableName);
        }

        [ExcelFunction(Name = "DuckPing", Description = "Tests DuckDBExcelBridge connectivity.")]
        public static string DuckPing()
        {
            return "OK | DuckDBExcelBridge";
        }

        [ExcelFunction(Name = "DuckVersion", Description = "Returns the DuckDB engine version.")]
        public static object DuckVersion()
        {
            return DuckSQL("SELECT version()");
        }

        [ExcelFunction(Name = "DuckSQL", Description = "Executes a DuckDB scalar SQL query.")]
        public static object DuckSQL(string sql)
        {
            try
            {
                using var cmd = GetConnection().CreateCommand();
                cmd.CommandText = sql;

                var result = cmd.ExecuteScalar();
                return result == null || result == DBNull.Value ? "" : result;
            }
            catch (DuckDBException ex)
            {
                return FormatDuckDbError(ex);
            }
            catch (Exception ex)
            {
                return FormatFatalError(ex);
            }
        }

        [ExcelFunction(Name = "DuckExec", Description = "Executes a DuckDB SQL command and returns OK.")]
        public static object DuckExec(string sql)
        {
            try
            {
                using var cmd = GetConnection().CreateCommand();
                cmd.CommandText = sql;
                cmd.ExecuteNonQuery();

                return "OK";
            }
            catch (DuckDBException ex)
            {
                return FormatDuckDbError(ex);
            }
            catch (Exception ex)
            {
                return FormatFatalError(ex);
            }
        }

        [ExcelFunction(Name = "DuckQuery", Description = "Executes a DuckDB query and returns a table with headers.")]
        public static object DuckQuery(string sql)
        {
            try
            {
                using var cmd = GetConnection().CreateCommand();
                cmd.CommandText = sql;

                using var reader = (DuckDBDataReader)cmd.ExecuteReader();
                return ReaderToArray(reader, includeHeaders: true);
            }
            catch (DuckDBException ex)
            {
                return FormatDuckDbError(ex);
            }
            catch (Exception ex)
            {
                return FormatFatalError(ex);
            }
        }

        [ExcelFunction(Name = "DuckQueryNoHeaders", Description = "Executes a DuckDB query and returns a table without headers.")]
        public static object DuckQueryNoHeaders(string sql)
        {
            try
            {
                using var cmd = GetConnection().CreateCommand();
                cmd.CommandText = sql;

                using var reader = (DuckDBDataReader)cmd.ExecuteReader();
                return ReaderToArray(reader, includeHeaders: false);
            }
            catch (DuckDBException ex)
            {
                return FormatDuckDbError(ex);
            }
            catch (Exception ex)
            {
                return FormatFatalError(ex);
            }
        }

        [ExcelFunction(Name = "DuckQueryNumeric", Description = "Executes a DuckDB query and returns a numeric matrix.")]
        public static object DuckQueryNumeric(string sql)
        {
            try
            {
                using var cmd = GetConnection().CreateCommand();
                cmd.CommandText = sql;

                using var reader = cmd.ExecuteReader();

                int cols = reader.FieldCount;
                var rows = new System.Collections.Generic.List<double[]>();

                while (reader.Read())
                {
                    var row = new double[cols];

                    for (int c = 0; c < cols; c++)
                    {
                        var value = reader.GetValue(c);
                        row[c] = value == DBNull.Value ? double.NaN : Convert.ToDouble(value);
                    }

                    rows.Add(row);
                }

                var output = new double[rows.Count, cols];

                for (int r = 0; r < rows.Count; r++)
                    for (int c = 0; c < cols; c++)
                        output[r, c] = rows[r][c];

                return output;
            }
            catch (DuckDBException ex)
            {
                return FormatDuckDbError(ex);
            }
            catch (Exception ex)
            {
                return FormatFatalError(ex);
            }
        }

        [ExcelFunction(Name = "DuckReset", Description = "Resets DuckDB back to a new in-memory database.")]
        public static string DuckReset()
        {
            lock (_lock)
            {
                _conn?.Dispose();
                _conn = null;
                _dbPath = ":memory:";
            }

            return "OK | DuckDB reset to :memory:";
        }

        [ExcelFunction(Name = "DuckImportRange", Description = "Imports an Excel range into a DuckDB table using SQL VALUES.")]
        public static object DuckImportRange(string tableName, object[,] data, bool firstRowHeaders)
        {
            try
            {
                int rowCount = data.GetLength(0);
                int colCount = data.GetLength(1);

                if (rowCount == 0 || colCount == 0)
                    return "No data";

                int startRow = firstRowHeaders ? 1 : 0;

                var columnNames = new string[colCount];

                for (int c = 0; c < colCount; c++)
                {
                    if (firstRowHeaders)
                    {
                        var rawName = data[0, c]?.ToString();
                        columnNames[c] = string.IsNullOrWhiteSpace(rawName)
                            ? $"col{c + 1}"
                            : rawName.Trim();
                    }
                    else
                    {
                        columnNames[c] = $"col{c + 1}";
                    }
                }

                var sql = new StringBuilder();
                sql.Append($"CREATE OR REPLACE TABLE {QuoteIdentifier(tableName)} AS SELECT * FROM (VALUES ");

                for (int r = startRow; r < rowCount; r++)
                {
                    if (r > startRow)
                        sql.Append(", ");

                    sql.Append("(");

                    for (int c = 0; c < colCount; c++)
                    {
                        if (c > 0)
                            sql.Append(", ");

                        sql.Append(ToSqlLiteral(data[r, c]));
                    }

                    sql.Append(")");
                }

                sql.Append(") AS t(");

                for (int c = 0; c < colCount; c++)
                {
                    if (c > 0)
                        sql.Append(", ");

                    sql.Append(QuoteIdentifier(columnNames[c]));
                }

                sql.Append(")");

                using var cmd = GetConnection().CreateCommand();
                cmd.CommandText = sql.ToString();
                cmd.ExecuteNonQuery();

                return $"OK | Imported {rowCount - startRow} rows x {colCount} columns into {tableName}";
            }
            catch (DuckDBException ex)
            {
                return FormatDuckDbError(ex);
            }
            catch (Exception ex)
            {
                return FormatFatalError(ex);
            }
        }

        [ExcelFunction(Name = "DuckImportRangeBulk", Description = "Bulk imports an Excel range into a DuckDB table using a temporary CSV file.")]
        public static object DuckImportRangeBulk(string tableName, object[,] data, bool firstRowHeaders)
        {
            string? tempFile = null;

            try
            {
                int rowCount = data.GetLength(0);
                int colCount = data.GetLength(1);

                if (rowCount == 0 || colCount == 0)
                    return "No data";

                tempFile = Path.Combine(
                    Path.GetTempPath(),
                    "duck_excel_" + Guid.NewGuid().ToString("N") + ".csv"
                );

                using (var writer = new StreamWriter(tempFile, false, new UTF8Encoding(false)))
                {
                    if (!firstRowHeaders)
                    {
                        for (int c = 0; c < colCount; c++)
                        {
                            if (c > 0) writer.Write(",");
                            writer.Write($"col{c + 1}");
                        }

                        writer.WriteLine();
                    }

                    for (int r = 0; r < rowCount; r++)
                    {
                        for (int c = 0; c < colCount; c++)
                        {
                            if (c > 0) writer.Write(",");
                            writer.Write(ToCsvField(data[r, c]));
                        }

                        writer.WriteLine();
                    }
                }

                string sql =
                    $"CREATE OR REPLACE TABLE {QuoteIdentifier(tableName)} AS " +
                    $"SELECT * FROM read_csv_auto('{EscapeSqlString(tempFile)}', header=true)";

                using var cmd = GetConnection().CreateCommand();
                cmd.CommandText = sql;
                cmd.ExecuteNonQuery();

                int importedRows = firstRowHeaders ? rowCount - 1 : rowCount;

                return $"OK | Bulk imported {importedRows} rows x {colCount} columns into {tableName}";
            }
            catch (DuckDBException ex)
            {
                return FormatDuckDbError(ex);
            }
            catch (Exception ex)
            {
                return FormatFatalError(ex);
            }
            finally
            {
                try
                {
                    if (tempFile != null && File.Exists(tempFile))
                        File.Delete(tempFile);
                }
                catch
                {
                    // Ignore temp file cleanup errors
                }
            }
        }


        // -----------------------------
        // Phase II convenience helpers
        // -----------------------------

        private static int NormalizeLimit(int maxRows, int defaultRows = 20)
        {
            if (maxRows <= 0)
                return defaultRows;

            return Math.Min(maxRows, 1000000);
        }

        [ExcelFunction(Name = "DuckLoadCsv", Description = "Loads a CSV file into a DuckDB table.")]
        public static object DuckLoadCsv(string tableName, string filePath)
        {
            try
            {
                string sql =
                    $"CREATE OR REPLACE TABLE {QuoteIdentifier(tableName)} AS " +
                    $"SELECT * FROM read_csv_auto('{EscapeSqlString(filePath)}', header=true)";

                using var cmd = GetConnection().CreateCommand();
                cmd.CommandText = sql;
                cmd.ExecuteNonQuery();

                return $"OK | Loaded CSV into {tableName}";
            }
            catch (DuckDBException ex)
            {
                return FormatDuckDbError(ex);
            }
            catch (Exception ex)
            {
                return FormatFatalError(ex);
            }
        }

        [ExcelFunction(Name = "DuckLoadParquet", Description = "Loads a Parquet file into a DuckDB table.")]
        public static object DuckLoadParquet(string tableName, string filePath)
        {
            try
            {
                string sql =
                    $"CREATE OR REPLACE TABLE {QuoteIdentifier(tableName)} AS " +
                    $"SELECT * FROM read_parquet('{EscapeSqlString(filePath)}')";

                using var cmd = GetConnection().CreateCommand();
                cmd.CommandText = sql;
                cmd.ExecuteNonQuery();

                return $"OK | Loaded Parquet into {tableName}";
            }
            catch (DuckDBException ex)
            {
                return FormatDuckDbError(ex);
            }
            catch (Exception ex)
            {
                return FormatFatalError(ex);
            }
        }

        [ExcelFunction(Name = "DuckRegisterParquet", Description = "Registers a Parquet file as a DuckDB view.")]
        public static object DuckRegisterParquet(string viewName, string filePath)
        {
            try
            {
                string sql =
                    $"CREATE OR REPLACE VIEW {QuoteIdentifier(viewName)} AS " +
                    $"SELECT * FROM read_parquet('{EscapeSqlString(filePath)}')";

                using var cmd = GetConnection().CreateCommand();
                cmd.CommandText = sql;
                cmd.ExecuteNonQuery();

                return $"OK | Parquet registered as {viewName}";
            }
            catch (DuckDBException ex)
            {
                return FormatDuckDbError(ex);
            }
            catch (Exception ex)
            {
                return FormatFatalError(ex);
            }
        }

        [ExcelFunction(Name = "DuckPreview", Description = "Returns the first N rows from a DuckDB table or view.")]
        public static object DuckPreview(string objectName, int maxRows)
        {
            int limit = NormalizeLimit(maxRows);
            return DuckQuery($"SELECT * FROM {QuoteIdentifier(objectName)} LIMIT {limit}");
        }

        [ExcelFunction(Name = "DuckHead", Description = "Returns the first N rows from a DuckDB table or view. Alias for DuckPreview.")]
        public static object DuckHead(string objectName, int maxRows)
        {
            return DuckPreview(objectName, maxRows);
        }

        [ExcelFunction(Name = "DuckPreviewSql", Description = "Returns the first N rows from a SQL query.")]
        public static object DuckPreviewSql(string sql, int maxRows)
        {
            int limit = NormalizeLimit(maxRows);
            return DuckQuery($"SELECT * FROM ({sql}) AS q LIMIT {limit}");
        }

        [ExcelFunction(Name = "DuckQueryRange", Description = "Queries an existing DuckDB table or view with an optional WHERE clause.")]
        public static object DuckQueryRange(string tableName, string whereClause)
        {
            string where = string.IsNullOrWhiteSpace(whereClause)
                ? ""
                : " WHERE " + whereClause.Trim();

            return DuckQuery($"SELECT * FROM {QuoteIdentifier(tableName)}{where}");
        }

        [ExcelFunction(Name = "DuckSQLRange", Description = "Runs a scalar SELECT expression against an existing DuckDB table or view with an optional WHERE clause.")]
        public static object DuckSQLRange(string tableName, string selectExpression, string whereClause)
        {
            string where = string.IsNullOrWhiteSpace(whereClause)
                ? ""
                : " WHERE " + whereClause.Trim();

            return DuckSQL($"SELECT {selectExpression} FROM {QuoteIdentifier(tableName)}{where}");
        }

        [ExcelFunction(Name = "DuckQueryExcelRange", Description = "Imports an Excel range as a named table, then runs a DuckDB query against it.")]
        public static object DuckQueryExcelRange(string tableName, object[,] data, bool firstRowHeaders, string sql)
        {
            var importResult = DuckSetTable(tableName, data, firstRowHeaders);

            if (importResult is string message && !message.StartsWith("OK", StringComparison.OrdinalIgnoreCase))
                return message;

            return DuckQuery(sql);
        }

        [ExcelFunction(Name = "DuckSQLExcelRange", Description = "Imports an Excel range as a named table, then runs a scalar DuckDB SQL query against it.")]
        public static object DuckSQLExcelRange(string tableName, object[,] data, bool firstRowHeaders, string sql)
        {
            var importResult = DuckSetTable(tableName, data, firstRowHeaders);

            if (importResult is string message && !message.StartsWith("OK", StringComparison.OrdinalIgnoreCase))
                return message;

            return DuckSQL(sql);
        }

        [ExcelFunction(Name = "DuckExportParquet", Description = "Exports a DuckDB query result to a Parquet file.")]
        public static object DuckExportParquet(string sql, string filePath)
        {
            try
            {
                string copySql = $"COPY ({sql}) TO '{EscapeSqlString(filePath)}' (FORMAT PARQUET)";

                using var cmd = GetConnection().CreateCommand();
                cmd.CommandText = copySql;
                cmd.ExecuteNonQuery();

                return $"OK | Exported query result to {filePath}";
            }
            catch (DuckDBException ex)
            {
                return FormatDuckDbError(ex);
            }
            catch (Exception ex)
            {
                return FormatFatalError(ex);
            }
        }

        [ExcelFunction(Name = "DuckRegisterCsv", Description = "Registers a CSV file as a DuckDB view.")]
        public static object DuckRegisterCsv(string viewName, string filePath)
        {
            try
            {
                string sql =
                    $"CREATE OR REPLACE VIEW {QuoteIdentifier(viewName)} AS " +
                    $"SELECT * FROM read_csv_auto('{EscapeSqlString(filePath)}')";

                using var cmd = GetConnection().CreateCommand();
                cmd.CommandText = sql;
                cmd.ExecuteNonQuery();

                return $"OK | CSV registered as {viewName}";
            }
            catch (DuckDBException ex)
            {
                return FormatDuckDbError(ex);
            }
            catch (Exception ex)
            {
                return FormatFatalError(ex);
            }
        }

        [ExcelFunction(Name = "DuckReadCsv", Description = "Reads a CSV file and returns rows to Excel.")]
        public static object DuckReadCsv(string filePath, int maxRows)
        {
            string limit = maxRows > 0 ? $" LIMIT {maxRows}" : "";
            return DuckQuery($"SELECT * FROM read_csv_auto('{EscapeSqlString(filePath)}'){limit}");
        }

        [ExcelFunction(Name = "DuckReadParquet", Description = "Reads a Parquet file and returns rows to Excel.")]
        public static object DuckReadParquet(string filePath, int maxRows)
        {
            string limit = maxRows > 0 ? $" LIMIT {maxRows}" : "";
            return DuckQuery($"SELECT * FROM read_parquet('{EscapeSqlString(filePath)}'){limit}");
        }

        [ExcelFunction(Name = "DuckExportCsv", Description = "Exports a DuckDB query result to a CSV file.")]
        public static object DuckExportCsv(string sql, string filePath)
        {
            try
            {
                string copySql =
                    $"COPY ({sql}) TO '{EscapeSqlString(filePath)}' " +
                    "(HEADER, DELIMITER ',')";

                using var cmd = GetConnection().CreateCommand();
                cmd.CommandText = copySql;
                cmd.ExecuteNonQuery();

                return $"OK | Exported query result to {filePath}";
            }
            catch (DuckDBException ex)
            {
                return FormatDuckDbError(ex);
            }
            catch (Exception ex)
            {
                return FormatFatalError(ex);
            }
        }

        [ExcelFunction(Name = "DuckTables", Description = "Lists DuckDB tables and views.")]
        public static object DuckTables()
        {
            return DuckQuery(
                "SELECT table_schema, table_name, table_type " +
                "FROM information_schema.tables " +
                "ORDER BY table_schema, table_name"
            );
        }

        [ExcelFunction(Name = "DuckColumns", Description = "Lists columns for a DuckDB table or view.")]
        public static object DuckColumns(string tableName)
        {
            return DuckQuery(
                "SELECT column_name, data_type, ordinal_position " +
                "FROM information_schema.columns " +
                $"WHERE table_name = '{EscapeSqlString(tableName)}' " +
                "ORDER BY ordinal_position"
            );
        }

        [ExcelFunction(Name = "DuckDescribe", Description = "Describes a DuckDB table or view.")]
        public static object DuckDescribe(string tableName)
        {
            return DuckQuery($"DESCRIBE {QuoteIdentifier(tableName)}");
        }

        [ExcelFunction(Name = "DuckCountRows", Description = "Counts rows in a DuckDB table or view.")]
        public static object DuckCountRows(string tableName)
        {
            return DuckSQL($"SELECT COUNT(*) FROM {QuoteIdentifier(tableName)}");
        }



        // -----------------------------
        // Phase III workflow helpers
        // Persistent databases, cached query tables, and range/external-file joins
        // -----------------------------

        private static void EnsureParentDirectoryExists(string filePath)
        {
            var directory = Path.GetDirectoryName(filePath);

            if (!string.IsNullOrWhiteSpace(directory) && !Directory.Exists(directory))
                Directory.CreateDirectory(directory);
        }

        [ExcelFunction(Name = "DuckOpenFile", Description = "Opens a persistent DuckDB database file and creates the parent folder if needed.")]
        public static string DuckOpenFile(string filePath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(filePath))
                    return "ERROR | Database file path cannot be blank";

                EnsureParentDirectoryExists(filePath);
                return DuckOpen(filePath);
            }
            catch (DuckDBException ex)
            {
                return FormatDuckDbError(ex);
            }
            catch (Exception ex)
            {
                return FormatFatalError(ex);
            }
        }

        [ExcelFunction(Name = "DuckOpenTempFile", Description = "Opens a persistent DuckDB database file in the temp folder.")]
        public static string DuckOpenTempFile(string databaseName)
        {
            try
            {
                string name = string.IsNullOrWhiteSpace(databaseName)
                    ? "excelbridge.duckdb"
                    : databaseName.Trim();

                if (!name.EndsWith(".duckdb", StringComparison.OrdinalIgnoreCase))
                    name += ".duckdb";

                string path = Path.Combine(Path.GetTempPath(), name);
                return DuckOpenFile(path);
            }
            catch (DuckDBException ex)
            {
                return FormatDuckDbError(ex);
            }
            catch (Exception ex)
            {
                return FormatFatalError(ex);
            }
        }

        [ExcelFunction(Name = "DuckCheckpoint", Description = "Checkpoints the current persistent DuckDB database.")]
        public static object DuckCheckpoint()
        {
            return DuckExec("CHECKPOINT");
        }

        [ExcelFunction(Name = "DuckDatabaseSize", Description = "Returns the current DuckDB database file size in bytes, or blank for :memory:.")]
        public static object DuckDatabaseSize()
        {
            try
            {
                if (string.Equals(_dbPath, ":memory:", StringComparison.OrdinalIgnoreCase))
                    return "";

                if (!File.Exists(_dbPath))
                    return "";

                return new FileInfo(_dbPath).Length;
            }
            catch (Exception ex)
            {
                return FormatFatalError(ex);
            }
        }

        [ExcelFunction(Name = "DuckSetTableFast", Description = "Fast bulk import alias for DuckSetTable/DuckImportRangeBulk.")]
        public static object DuckSetTableFast(string tableName, object[,] data, bool firstRowHeaders)
        {
            return DuckImportRangeBulk(tableName, data, firstRowHeaders);
        }

        [ExcelFunction(Name = "DuckCacheQuery", Description = "Creates or replaces a cached DuckDB table from a SQL query.")]
        public static object DuckCacheQuery(string tableName, string sql)
        {
            try
            {
                string createSql =
                    $"CREATE OR REPLACE TABLE {QuoteIdentifier(tableName)} AS " +
                    $"SELECT * FROM ({sql}) AS q";

                using var cmd = GetConnection().CreateCommand();
                cmd.CommandText = createSql;
                cmd.ExecuteNonQuery();

                return $"OK | Cached query as {tableName}";
            }
            catch (DuckDBException ex)
            {
                return FormatDuckDbError(ex);
            }
            catch (Exception ex)
            {
                return FormatFatalError(ex);
            }
        }

        [ExcelFunction(Name = "DuckAppendQuery", Description = "Appends a SQL query result into an existing DuckDB table.")]
        public static object DuckAppendQuery(string tableName, string sql)
        {
            try
            {
                string appendSql =
                    $"INSERT INTO {QuoteIdentifier(tableName)} " +
                    $"SELECT * FROM ({sql}) AS q";

                using var cmd = GetConnection().CreateCommand();
                cmd.CommandText = appendSql;
                cmd.ExecuteNonQuery();

                return $"OK | Appended query result into {tableName}";
            }
            catch (DuckDBException ex)
            {
                return FormatDuckDbError(ex);
            }
            catch (Exception ex)
            {
                return FormatFatalError(ex);
            }
        }

        [ExcelFunction(Name = "DuckLoadCsvOptions", Description = "Loads a CSV file into a DuckDB table with options.")]
        public static object DuckLoadCsvOptions(string tableName, string filePath, object header, string delimiter)
        {
            try
            {
                bool hasHeader = Convert.ToBoolean(header, CultureInfo.InvariantCulture);

                string sql =
                    $"CREATE OR REPLACE TABLE {QuoteIdentifier(tableName)} AS " +
                    $"SELECT * FROM read_csv_auto(" +
                    $"'{EscapeSqlString(filePath)}', " +
                    $"header={(hasHeader ? "true" : "false")}, " +
                    $"delim='{EscapeSqlString(delimiter)}')";

                using var cmd = GetConnection().CreateCommand();
                cmd.CommandText = sql;
                cmd.ExecuteNonQuery();

                return $"OK | Loaded CSV into {tableName}";
            }
            catch (Exception ex)
            {
                return "ERROR | " + ex.Message;
            }
        }

        [ExcelFunction(Name = "DuckRegisterCsvOptions", Description = "Registers a CSV file as a DuckDB view with an explicit header option.")]
        public static object DuckRegisterCsvOptions(string viewName, string filePath, bool header)
        {
            try
            {
                string sql =
                    $"CREATE OR REPLACE VIEW {QuoteIdentifier(viewName)} AS " +
                    $"SELECT * FROM read_csv_auto('{EscapeSqlString(filePath)}', header={header.ToString().ToLowerInvariant()})";

                using var cmd = GetConnection().CreateCommand();
                cmd.CommandText = sql;
                cmd.ExecuteNonQuery();

                return $"OK | CSV registered as {viewName}";
            }
            catch (DuckDBException ex)
            {
                return FormatDuckDbError(ex);
            }
            catch (Exception ex)
            {
                return FormatFatalError(ex);
            }
        }

        [ExcelFunction(Name = "DuckJoinRangeCsv", Description = "Imports an Excel range, registers a CSV view, then runs a join/query SQL statement.")]
        public static object DuckJoinRangeCsv(string rangeTableName, object[,] data, bool firstRowHeaders, string csvViewName, string csvFilePath, string sql)
        {
            var importResult = DuckSetTableFast(rangeTableName, data, firstRowHeaders);

            if (importResult is string importMessage && !importMessage.StartsWith("OK", StringComparison.OrdinalIgnoreCase))
                return importMessage;

            var registerResult = DuckRegisterCsv(csvViewName, csvFilePath);

            if (registerResult is string registerMessage && !registerMessage.StartsWith("OK", StringComparison.OrdinalIgnoreCase))
                return registerMessage;

            return DuckQuery(sql);
        }

        [ExcelFunction(Name = "DuckJoinRangeParquet", Description = "Imports an Excel range, registers a Parquet view, then runs a join/query SQL statement.")]
        public static object DuckJoinRangeParquet(string rangeTableName, object[,] data, bool firstRowHeaders, string parquetViewName, string parquetFilePath, string sql)
        {
            var importResult = DuckSetTableFast(rangeTableName, data, firstRowHeaders);

            if (importResult is string importMessage && !importMessage.StartsWith("OK", StringComparison.OrdinalIgnoreCase))
                return importMessage;

            var registerResult = DuckRegisterParquet(parquetViewName, parquetFilePath);

            if (registerResult is string registerMessage && !registerMessage.StartsWith("OK", StringComparison.OrdinalIgnoreCase))
                return registerMessage;

            return DuckQuery(sql);
        }

        [ExcelFunction(Name = "DuckJoinRangeCsvScalar", Description = "Imports an Excel range, registers a CSV view, then runs a scalar SQL statement.")]
        public static object DuckJoinRangeCsvScalar(string rangeTableName, object[,] data, bool firstRowHeaders, string csvViewName, string csvFilePath, string sql)
        {
            var importResult = DuckSetTableFast(rangeTableName, data, firstRowHeaders);

            if (importResult is string importMessage && !importMessage.StartsWith("OK", StringComparison.OrdinalIgnoreCase))
                return importMessage;

            var registerResult = DuckRegisterCsv(csvViewName, csvFilePath);

            if (registerResult is string registerMessage && !registerMessage.StartsWith("OK", StringComparison.OrdinalIgnoreCase))
                return registerMessage;

            return DuckSQL(sql);
        }

        [ExcelFunction(Name = "DuckJoinRangeParquetScalar", Description = "Imports an Excel range, registers a Parquet view, then runs a scalar SQL statement.")]
        public static object DuckJoinRangeParquetScalar(string rangeTableName, object[,] data, bool firstRowHeaders, string parquetViewName, string parquetFilePath, string sql)
        {
            var importResult = DuckSetTableFast(rangeTableName, data, firstRowHeaders);

            if (importResult is string importMessage && !importMessage.StartsWith("OK", StringComparison.OrdinalIgnoreCase))
                return importMessage;

            var registerResult = DuckRegisterParquet(parquetViewName, parquetFilePath);

            if (registerResult is string registerMessage && !registerMessage.StartsWith("OK", StringComparison.OrdinalIgnoreCase))
                return registerMessage;

            return DuckSQL(sql);
        }

        [ExcelFunction(Name = "DuckDrop", Description = "Drops a DuckDB table or view if it exists.")]
        public static object DuckDrop(string objectName)
        {
            try
            {
                using var cmd = GetConnection().CreateCommand();

                cmd.CommandText = $"DROP TABLE IF EXISTS {QuoteIdentifier(objectName)}";
                cmd.ExecuteNonQuery();

                cmd.CommandText = $"DROP VIEW IF EXISTS {QuoteIdentifier(objectName)}";
                cmd.ExecuteNonQuery();

                return $"OK | Dropped {objectName} if it existed";
            }
            catch (DuckDBException ex)
            {
                return FormatDuckDbError(ex);
            }
            catch (Exception ex)
            {
                return FormatFatalError(ex);
            }
        }
    }
}