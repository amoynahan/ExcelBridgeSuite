# Core/ExcelBridge

Core/ExcelBridge is the native C++ bridge for ExcelBridgeSuite.

It provides high-performance functionality for Excel by enabling direct integration of C++ code and libraries through Excel-DNA.

## Purpose

The goal of Core/ExcelBridge is to extend Excel beyond scripting languages by allowing:

- direct execution of C++ functions from Excel
- high-performance numerical computation
- efficient memory handling for large datasets
- integration with existing C++ libraries

This makes it possible to build production-grade, performance-critical workflows directly in Excel.

## Key capabilities

- Native C++ functions exposed to Excel
- Fast data transfer between Excel and C++
- Support for matrix and table operations
- Foundation for performance used by other bridges

## How it fits in ExcelBridgeSuite

Core/ExcelBridge acts as both:
- a standalone C++ bridge
- a performance layer that other bridges can build on

Other bridges (R, Python, Julia) focus on flexibility and ecosystem integration, while Core/ExcelBridge focuses on speed and control.

## Example use cases

- Numerical algorithms (linear algebra, optimization)
- Large dataset processing
- Custom financial or scientific models
- Wrapping existing C++ libraries for Excel use

## Getting started

1. Open the Core/ExcelBridge project in Visual Studio
2. Build the project
3. Load the generated .xll file in Excel
4. Call exposed functions directly from Excel cells

## Next steps

Future documentation will include:
- creating your first C++ Excel function
- passing arrays and tables between Excel and C++
- integrating third-party C++ libraries

## Notes

This component is under active development and will continue to expand in functionality and performance.

