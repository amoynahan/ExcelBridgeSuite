# ExcelBridgeSuite Installation Guide

ExcelBridgeSuite provides three Excel add-ins:

- RExcelBridge (R integration)  
- JuliaExcelBridge (Julia integration)  
- PythonExcelBridge (Python integration)  

You can install them in two ways:

- Option 1: Use precompiled add-ins (recommended)  
- Option 2: Build from source using Visual Studio  

---

## Option 1: Install Precompiled Add-ins (Recommended)

1. Go to the Releases section of this repository  
2. Download the latest release zip file  
3. Extract the contents to a local folder (do not run from the zip)  

Each bridge includes a `publish` folder containing all required files.

---

## What’s in the publish folder

Each bridge has a folder like:

RExcelBridge\bin\Release\net8.0-windows\publish  
JuliaExcelBridge\...  
PythonExcelBridge\...  

Typical contents:

- `<BridgeName>64.xll`  
  The Excel add-in file. This is what you load into Excel.

- `<BridgeName>.dll`  
  The compiled .NET assembly containing the bridge logic.

- `<BridgeName>.dna`  
  Excel-DNA configuration file that tells Excel how to load the add-in.

- `*.deps.json`, `*.runtimeconfig.json`  
  .NET runtime dependency and configuration files required for execution.

- `ExcelDna.Integration.dll` (and related files)  
  Excel-DNA runtime components.

- `worker.R`, `startup.R` (R bridge only)  
  Scripts used to execute R code and manage the session.

- `plot-path.txt` (optional)  
  Overrides the default directory where plots are saved.

All files must remain in the same folder.  
Do not move the `.xll` file by itself.

---

## Attach the Add-in in Excel

1. Open Excel  
2. Go to File → Options → Add-ins  
3. At the bottom:
   - Manage: Excel Add-ins  
   - Click Go  
4. Click Browse  
5. Navigate to the publish folder  
6. Select one of:
   - RExcelBridge64.xll  
   - JuliaExcelBridge64.xll  
   - PythonExcelBridge64.xll  
7. Click OK  
8. Ensure the add-in is checked  

Excel will load the bridge.

If prompted about security:
- Click Enable  
- Or unblock the files via file properties if needed  

---

## Option 2: Build from Source (Visual Studio)

Use this option if you want to modify or extend the bridges.

Requirements:

- Visual Studio 2022  
- .NET 8.0 (Windows)  
- Excel-DNA (already included in the project)  

Steps:

1. Clone the repository  
2. Open the solution file in Visual Studio  
3. Select a project:
   - RExcelBridge  
   - JuliaExcelBridge  
   - PythonExcelBridge  
4. Set configuration:
   - Release  
   - x64 / Any CPU (as configured)  
5. Build the project  

The compiled add-in will be created in:

bin\Release\net8.0-windows\publish  

Then follow the same steps above to attach the `.xll` file in Excel.

---

## Notes

- Always run the add-in from the publish folder  
- Do not separate files from the `.xll`  
- If Excel does not load the add-in:
  - Uncheck it, restart Excel, and reattach  
- When updating versions:
  - Replace the entire publish folder contents  

---

## Next Step

See the Usage section for:

- Running code from Excel  
- Passing data between Excel and R / Julia / Python  
- Generating plots  
- Using bridge functions  
