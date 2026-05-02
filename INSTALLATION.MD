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

## Load the Excel Add-in

1. Open Excel  
2. Click **File**  
3. Click **Options**  
4. In the Excel Options window, click **Add-ins**  
5. At the bottom, in the **Manage** dropdown, select **Excel Add-ins** and click **Go**  
6. In the Add-ins window, click **Browse**  
7. Navigate to your publish directory:  
   `RExcelBridge/bin/Release/net8.0-windows/publish`  
8. Select the file:  
   `RExcelBridge64-packed.xll`  
9. Click **OK**  

The add-in should now be loaded and available in Excel.

---

## Trust and Security Settings (Required for .xll Add-ins)

If Excel blocks the add-in or it does not load, follow one of the options below:

### Option 1: Unblock the file (recommended)

1. Close Excel  
2. In File Explorer, go to your publish folder  
3. Right-click `RExcelBridge64-packed.xll`  
4. Click **Properties**  
5. At the bottom, if you see a security message, check **Unblock**  
6. Click **Apply**, then **OK**  

---

### Option 2: Add the folder as a trusted location

1. Open Excel  
2. Click **File**  
3. Click **Options**  
4. Go to **Trust Center**  
5. Click **Trust Center Settings**  
6. Click **Trusted Locations**  
7. Click **Add new location**  
8. Browse to your publish folder:  
   `RExcelBridge/bin/Release/net8.0-windows/publish`  
9. Click **OK**  

---

### Option 3: Enable add-ins if disabled

1. Go to **File → Options → Add-ins**  
2. In **Manage**, select **Disabled Items** → click **Go**  
3. If the add-in appears, enable it  

---

After completing one of the above:

- Restart Excel  
- Re-load the add-in if needed (via **Browse**)  

---

**Note:**  
For broader distribution, consider digitally signing the `.xll` so users do not need to adjust security settings.

See the Usage section for:

- Running code from Excel  
- Passing data between Excel and R, Julia, and Python  
- Generating plots  
- Using bridge functions  




