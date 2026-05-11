# DuckDBExcelBridge User Guide

## Overview

DuckDBExcelBridge is an Excel-DNA add-in that allows Excel to interact directly with DuckDB using fast in-process SQL queries.

The bridge is designed for:

- Fast analytical queries
- Large datasets
- Local Parquet and CSV analysis
- Excel-based analytics workflows
- Lightweight embedded SQL processing
- HEOR and scientific data workflows

DuckDB runs inside the Excel process and does not require a separate database server.

---

## Installation

### Requirements

- Windows 10 or later
- Microsoft Excel 64-bit
- .NET 8 Runtime
- DuckDBExcelBridge publish folder

### Publish Folder Contents

Typical publish folder:

```text
DuckDBExcelBridge64.xll
DuckDBExcelBridge.dna
DuckDBExcelBridge.dll
DuckDB.NET.Data.dll
ExcelDna.Integration.dll
runtimes\
```

### Loading the Add-In

1. Open Excel.
2. Go to:

```text
File -> Options -> Add-ins
```

3. Select:

```text
Excel Add-ins -> Go
```

4. Browse to:

```text
DuckDBExcelBridge64.xll
```

5. Enable the add-in.

---

## Quick Start

### Verify the Bridge

In Excel:

```excel
=DuckPing()
```

Expected result:

```text
DuckDBExcelBridge OK
```

---

## Core Concepts

DuckDBExcelBridge maintains an in-memory DuckDB connection.

You can:

- Execute SQL
- Create tables
- Query Excel ranges
- Import and export data
- Query Parquet files
- Query CSV files
- Return tables directly into Excel

---

## Basic SQL Execution

### Execute SQL

```excel
=DuckExec("CREATE TABLE test(id INTEGER, value DOUBLE)")
```

Returns:

```text
OK
```

---

## Inserting Data

```excel
=DuckExec("INSERT INTO test VALUES (1, 10.5)")
```

```excel
=DuckExec("INSERT INTO test VALUES (2, 20.0)")
```

---

## Querying Data

```excel
=DuckQuery("SELECT * FROM test")
```

Results spill into Excel dynamically.

---

## Working with Excel Ranges

Suppose range `A1:C5` contains:

```text
ID    NAME    SCORE
1     John    10
2     Mary    20
```

Import the range:

```excel
=DuckImportTable("people", A1:C5, TRUE)
```

Arguments:

| Argument | Description |
|---|---|
| table name | Destination table |
| Excel range | Source data |
| TRUE/FALSE | Whether first row contains headers |

---

## Query Imported Excel Tables

```excel
=DuckQuery("SELECT * FROM people")
```

---

## Export DuckDB Tables Back to Excel

```excel
=DuckTable("people")
```

Returns the entire table into Excel.

---

## Querying CSV Files

### Query CSV Directly

```excel
=DuckQuery("SELECT * FROM read_csv_auto('C:/data/file.csv')")
```

Example with filtering:

```excel
=DuckQuery("SELECT * FROM read_csv_auto('C:/data/file.csv') WHERE age > 65")
```

---

## Querying Parquet Files

### Query Parquet Directly

```excel
=DuckQuery("SELECT * FROM 'C:/data/file.parquet'")
```

Example aggregation:

```excel
=DuckQuery("SELECT gender, COUNT(*) FROM 'C:/data/file.parquet' GROUP BY gender")
```

---

## Joining Files and Excel Tables

```excel
=DuckQuery("SELECT a.*, b.score FROM people a LEFT JOIN 'C:/data/scores.parquet' b ON a.id = b.id")
```

---

## Creating Persistent DuckDB Databases

### Connect to Database File

```excel
=DuckOpen("C:/duckdb/mydatabase.duckdb")
```

This creates or opens a persistent database.

---

## Using In-Memory Databases

```excel
=DuckOpen(":memory:")
```

This database exists only during the Excel session.

---

## Table Management

### List Tables

```excel
=DuckQuery("SHOW TABLES")
```

### Describe Table

```excel
=DuckQuery("DESCRIBE people")
```

### Drop Table

```excel
=DuckExec("DROP TABLE people")
```

---

## Analytical Examples

### Aggregation

```excel
=DuckQuery("SELECT gender, AVG(cost) FROM claims GROUP BY gender")
```

### Window Functions

```excel
=DuckQuery("SELECT *, ROW_NUMBER() OVER(PARTITION BY patient ORDER BY date) rn FROM claims")
```

### Cohort Counts

```excel
=DuckQuery("SELECT year(indexdate), COUNT(*) FROM cohort GROUP BY 1")
```

---

## Performance Tips

### Prefer Bulk Transfers

Import/export large blocks rather than cell-by-cell operations.

Good:

```excel
=DuckImportTable("claims", A1:Z100000, TRUE)
```

Avoid many small transfers.

### Use Parquet for Large Datasets

Parquet is usually faster and more compact than CSV.

Preferred:

```sql
SELECT * FROM 'data.parquet'
```

### Keep Queries Set-Based

Good:

```sql
SELECT gender, COUNT(*)
FROM claims
GROUP BY gender
```

Avoid procedural row-by-row logic when possible.

---

## Error Handling

### Table Does Not Exist

```text
Catalog Error: Table not found
```

Verify the table name.

### File Not Found

```text
IO Error
```

Verify file path spelling.

### Excel Spill Errors

Ensure destination cells are empty.

---

## Example Workflow

### Step 1: Open Database

```excel
=DuckOpen("C:/duckdb/heor.duckdb")
```

### Step 2: Import Excel Data

```excel
=DuckImportTable("patients", A1:G5000, TRUE)
```

### Step 3: Query Results

```excel
=DuckQuery("SELECT gender, COUNT(*) FROM patients GROUP BY gender")
```

### Step 4: Save Results

```excel
=DuckExec("CREATE TABLE summary AS SELECT gender, COUNT(*) cnt FROM patients GROUP BY gender")
```

---

## Suggested Use Cases

DuckDBExcelBridge works especially well for:

- HEOR analytics
- Claims analysis
- Epidemiology workflows
- Local analytical pipelines
- Simulation outputs
- Clinical trial summaries
- Rapid SQL prototyping
- Large Excel datasets

---

## Architecture

```text
Excel
  ↓
Excel-DNA Add-in
  ↓
DuckDBExcelBridge
  ↓
DuckDB Engine
  ↓
Parquet / CSV / DuckDB Files
```

---

## Future Enhancements

Potential future enhancements include:

- Arrow integration
- Async/background queries
- Named cached datasets
- SQL using workbook range references
- Reporting helpers
- Persistent query sessions
- Parallel query execution

---

## Troubleshooting

### Add-In Does Not Load

Verify:

- 64-bit Excel
- .NET 8 installed
- All DLLs present in publish folder

### Excel Crashes

Check:

- Invalid native DLL references
- Missing runtime files
- Incorrect platform target

Recommended build:

```text
PlatformTarget = x64
TargetFramework = net8.0-windows
```

---

## Notes

DuckDBExcelBridge is optimized for analytical workloads rather than transactional processing.

Best performance occurs when:

- Working with large block transfers
- Using Parquet
- Keeping operations set-based
- Avoiding excessive Excel recalculation

---

## Additional Resources

DuckDB:

```text
https://duckdb.org
```

DuckDB Documentation:

```text
https://duckdb.org/docs/
```

Excel-DNA:

```text
https://excel-dna.net
```
