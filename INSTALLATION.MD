# ExcelBridgeSuite Installation Guide

ExcelBridgeSuite provides three Excel add-ins:

- RExcelBridge (R integration)  
- JuliaExcelBridge (Julia integration)  
- PythonExcelBridge (Python integration)  

You can install any or all of them using one of the methods below.

---

## Installation Options

- Option 1: Use precompiled add-ins (recommended)  
- Option 2: Build from source using Visual Studio  

---

## Option 1: Install Precompiled Add-ins (Recommended)

1. Go to the Releases section of this repository  
2. Download the latest release  
3. Extract the contents to a local folder  

Each bridge (R, Julia, Python) includes a `publish` folder containing all required files.

---

## Publish Folder Structure (Important)

Each bridge has a folder like:

RExcelBridge\bin\Release\net8.0-windows\publish  

This folder contains everything needed to run the add-in.

Typical contents:

- `<BridgeName>64.xll`  
  The Excel add-in. This is the file you load into Excel.

- `<BridgeName>.dll`  
  The compiled .NET assembly containing the bridge logic.

- `<BridgeName>.dna`  
  Excel-DNA configuration file that tells Excel how to load the add-in.

- `*.deps.json`, `*.runtimeconfig.json`  
  .NET runtime configuration and dependency files.

- `ExcelDna.Integration.dll` (and related files)  
  Excel-DNA runtime components required by the add-in.

- `worker.R`, `startup.R` (R bridge only)  
  Scripts used to execute R code.

- `rscript-path.txt`  
  Path configuration file for locating `Rscript.exe`.

- `plot-path.txt`  
  Optional configuration file for controlling where plots are saved.

All files must remain together in this folder.  
Do not move or copy only the `.xll` file.

---

## rscript-path.txt (Required for RExcelBridge)

This file tells the add-in where to find `Rscript.exe`.

### Setup

1. Open `rscript-path.txt` in a text editor  
2. Enter the full path to your Rscript executable  

Example: C:\Program Files\R\R-4.3.2\bin\Rscript.exe


3. Save the file  

### Notes

- The path must be a full absolute path  
- Do not use just `Rscript`  
- If the path is incorrect or missing, RExcelBridge will not run  
- If you install a new version of R, update this file  

---

## plot-path.txt (Optional)

This file controls where plot images (PNG files) are written.

### Setup

1. Open `plot-path.txt`  
2. Enter a folder path where you want plots saved  

Example: bin\Release\net8.0-windows\publish


Then follow the same steps above to attach the `.xll` file in Excel.

---

## Notes

- Always run the add-in from the publish folder  
- Do not separate files from the `.xll`  
- If Excel does not load the add-in:
  - Uncheck it  
  - Restart Excel  
  - Reattach the add-in  
- When updating:
  - Replace all files in the publish folder  

---

## Next Step

See the Usage section for:

- Running code from Excel  
- Passing data between Excel and R, Julia, and Python  
- Generating plots  
- Using bridge functions  




