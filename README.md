# ExcelBridgeSuite

Note: The site is the testing phase. I will remove this note when I think it is ready.

ExcelBridgeSuite is a multi-language Excel add-in framework built with Excel-DNA.

Current bridges:
- RExcelBridge
- PythonExcelBridge
- JuliaExcelBridge

A C++ bridge is planned as a future addition.

## What this repo does

These add-ins are designed to let Excel work with external languages and runtimes through a consistent bridge-style approach.

Depending on the bridge, the add-ins can be used to:
- run code from Excel
- exchange data between Excel and the target language
- support helper functions for integration workflows
- support plotting and other extensions

## Repository structure

- `src/RExcelBridge` - R bridge
- `src/PythonExcelBridge` - Python bridge
- `src/JuliaExcelBridge` - Julia bridge

## Getting started

You have two options.

### Option 1: Use the release files
Download the latest packaged release files from the Releases section of this repository.

This is the easiest way to get started.

### Option 2: Build in Visual Studio
You can also build the projects yourself in Visual Studio.

General steps:
1. Clone or download the repository
2. Open the relevant project in Visual Studio
3. Restore packages if needed
4. Build the project
5. Load the generated `.xll` add-in in Excel

## Requirements

- Windows
- 64-bit Excel
- Visual Studio for local builds
- The corresponding runtime installed locally:
  - R for RExcelBridge
  - Python for PythonExcelBridge
  - Julia for JuliaExcelBridge

## Configuration

Some bridges currently use local path configuration text files for runtime or plot settings.

These may need to be updated for your local environment.

A future improvement will be to reduce or eliminate manual path configuration.

## Releases

Packaged builds are available in the Releases section.

Each bridge is distributed separately so you can download only what you need.

## Current status

The R, Python, and Julia bridge versions are posted.

Examples and additional documentation will be added soon.

## Notes

This repository is under active development.
