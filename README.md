# ExcelBridgeSuite

ExcelBridgeSuite is a multi-language Excel add-in framework built with Excel-DNA.

It is designed to extend Excel with external runtimes and high-performance native code, including integration with C++ libraries.

## Current bridges

- RExcelBridge  
- PythonExcelBridge  
- JuliaExcelBridge  
- Core/ExcelBridge (C++ bridge)

## What this repo does

ExcelBridgeSuite enables Excel to work with external languages and runtimes through a consistent bridge-style architecture.

Each bridge allows Excel to:
- run code from external runtimes
- exchange data between Excel and the target language
- support helper functions for integration workflows
- support plotting and other extensions

Core/ExcelBridge is the native C++ bridge. It enables high-performance data transfer and computation, and allows direct integration of C++ libraries into Excel workflows.

## Repository structure

- core/ExcelBridge - C++ bridge and native performance layer  
- src/RExcelBridge - R bridge  
- src/PythonExcelBridge - Python bridge  
- src/JuliaExcelBridge - Julia bridge  

## Getting started

You have two options.

### Option 1: Use the release files
Download the latest packaged release files from the Releases section of this repository.

This allows you to load the add-in in Excel and explore example functions.

The primary value of ExcelBridgeSuite is the ability to extend Excel with custom functionality, including integrating C++ libraries through the Core/ExcelBridge layer. The release files provide a starting point, but most users will want to build and customize.

### Option 2: Build in Visual Studio
You can build the projects yourself in Visual Studio to customize or extend functionality.

This is the recommended approach if you want to:
- integrate custom C++ libraries  
- add new functions to Excel  
- modify or extend existing bridges  
- build performance-critical workflows  

General steps:
1. Clone or download the repository  
2. Open the relevant project in Visual Studio  
3. Restore packages if needed  
4. Build the project  
5. Load the generated .xll add-in in Excel  

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

The R, Python, and Julia bridges are available.

The Core/ExcelBridge C++ bridge is under active development and will expand support for native extensions and performance-critical functionality.

Examples and additional documentation will be added.

## Notes

This repository is under active development.
