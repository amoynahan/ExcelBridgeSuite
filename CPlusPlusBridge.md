# ExcelBridge

A reusable high-performance C++ computational framework for Excel built with Excel-DNA, .NET, and native C++.

---

# Overview

Core/ExcelBridge is a lightweight framework for building modern Excel add-ins that integrate directly with native C++ libraries.

The project is intentionally designed as a reusable computational shell rather than a finished analytics product.

The primary value of the framework is the ability to:

- Integrate advanced native C++ libraries into Excel
- Maintain persistent computational state
- Transfer matrices efficiently
- Build scalable computational workflows
- Extend functionality incrementally over time

The framework separates:

- Excel integration
- Managed interop
- Native computation
- Persistent object management

This architecture makes it possible to evolve the project into highly specialized computational platforms without redesigning the entire stack.

---

# Core Idea

The framework itself is intentionally minimal.

The real value comes from linking advanced computational libraries such as:

- Eigen
- QuantLib
- Boost
- CGAL
- DuckDB
- Intel MKL
- OpenBLAS


The project provides the infrastructure needed to expose those libraries cleanly inside Excel.

---

# Design Goals

| Goal | Description |
|---|---|
| Reusable foundation | Build once and extend incrementally |
| Native performance | Heavy computation occurs in C++ |
| Persistent state | Objects survive between Excel calls |
| Fast transfer | Efficient matrix and array movement |
| Extensible architecture | Easy integration of new libraries |
| Modern tooling | Visual Studio 2022 + .NET 8 |
| Excel-first workflow | Simple deployment and usability |

---

# High-Level Architecture

```text
Excel
  ↓
Excel-DNA Add-In (.NET / C#)
  ↓
P/Invoke Layer
  ↓
Native C++ DLL
  ↓
Computational Libraries
```

---

# Technology Stack

| Layer | Technology |
|---|---|
| Excel Integration | Excel-DNA |
| Managed Layer | C# / .NET 8 |
| Native Layer | Modern C++ |
| Interop | P/Invoke |
| Packaging | ExcelDnaPack |
| Build System | Visual Studio 2022 |

---

# Current Project Structure

```text
Core/
├── ExcelBridge/
│   ├── Excel-DNA add-in
│   ├── C# wrapper layer
│   ├── Excel-facing functions
│   └── object registry interface
│
├── NativeMath/
│   ├── native C++ computation layer
│   ├── matrix utilities
│   ├── object store
│   └── exported DLL functions
│
└── publish/
    ├── *.xll
    ├── Native DLLs
    └── deployment files
```

---

# Build Philosophy

The build process intentionally creates a lightweight computational shell rather than a finished analytics platform.

The framework establishes:

- Excel integration
- Native DLL loading
- Managed/native interoperability
- Matrix transfer infrastructure
- Persistent object infrastructure
- Reusable extension patterns

The expectation is that developers will extend the framework by linking specialized computational libraries into the native C++ layer.

---

# Why This Architecture Matters

Traditional Excel integrations often hit limitations:

| Approach | Limitation |
|---|---|
| VBA | Slow numerical performance |
| COM Automation | Complex deployment |
| Python/R Bridges | External runtime dependencies |
| Traditional XLLs | Difficult extensibility |
| Pure .NET | Limited native library access |

This project attempts to combine:

- Excel usability
- Native C++ performance
- Persistent computational state
- Modern tooling
- Incremental extensibility

into a reusable framework.

---

# Key Architectural Concepts

## 1. Excel as the Front End

Excel acts primarily as:

- A UI layer
- A calculation trigger
- A reporting environment
- A lightweight workflow manager

The heavy computation occurs in native C++.

---

## 2. Native C++ as the Compute Engine

The native layer handles:

- Matrix operations
- Numerical algorithms
- Persistent objects
- Session-style computation
- Future scientific/quantitative workflows

This allows the framework to scale far beyond traditional VBA solutions.

---

## 3. Persistent Native Objects

One major design goal is persistent computation.

Instead of recreating objects on every recalculation:

```text
Excel Function
    → Create Native Object
    → Store Handle
    → Reuse Later
```

This is especially important for:

- QuantLib pricing engines
- Optimization models
- Monte Carlo simulations
- Scientific workflows
- Geometry engines

---

## 4. Efficient Matrix Transfer

The framework avoids slow cell-by-cell transfer patterns.

Instead, it transfers:

- Flat contiguous arrays
- Row-major matrices
- Bulk memory blocks

Conceptually similar to:

```vba
Range.Value = VariantArray
```

rather than looping through cells individually.

This becomes critical for:

- Large matrices
- Statistical modeling
- Simulation output
- Quantitative finance
- Scientific computing

---

# Current Features

## Connectivity

| Function | Purpose |
|---|---|
| `CPP_PING()` | Test native connectivity |
| `CPP_VERSION()` | Return framework version |
| `CPP_STATUS()` | Diagnostic status |

---

## Matrix Operations

| Function | Purpose |
|---|---|
| `CPP_TRANSPOSE()` | Matrix transpose |
| `CPP_MATRIX_ROUNDTRIP()` | Validate matrix transfer |
| `CPP_IDENTITY()` | Identity matrix generation |

---

## Object Store

| Function | Purpose |
|---|---|
| `CPP_OBJECT_CREATE()` | Create persistent object |
| `CPP_OBJECTS()` | List stored objects |
| `CPP_OBJECT_DETAIL()` | Inspect object metadata |

The object system is intentionally generic so future libraries can integrate cleanly.

---

# Intended Library Integrations

The framework is specifically designed for integration with advanced native libraries.

---

# Numerical & Linear Algebra

| Library | Purpose |
|---|---|
| Eigen | Dense/sparse linear algebra |
| Armadillo | MATLAB-style matrix workflows |
| Blaze | High-performance vectorized math |
| Intel MKL | Optimized numerical kernels |
| OpenBLAS | BLAS/LAPACK operations |

Potential uses:

- Regression
- PCA
- Matrix decompositions
- Optimization
- Statistical modeling

---

# Quantitative Finance

| Library | Purpose |
|---|---|
| QuantLib | Derivatives pricing and risk |
| Premia | Financial option pricing |
| Boost.Math | Financial/statistical functions |

Potential uses:

- Yield curves
- Swaption models
- Greeks
- Monte Carlo pricing
- Risk engines

A major motivation for the framework is enabling persistent QuantLib-style objects inside Excel.

---

# Computational Geometry & Spatial Analysis

| Library | Purpose |
|---|---|
| CGAL | Computational geometry |
| GEOS | Spatial geometry engine |
| GDAL | Geospatial processing |
| Boost.Geometry | Geometry algorithms |

Potential uses:

- Mesh generation
- Polygon operations
- Mapping workflows
- Scientific geometry models

---

# Scientific & Engineering Libraries

| Library | Purpose |
|---|---|
| Sundials | Differential equation solvers |
| PETSc | Parallel scientific computing |
| Trilinos | Large-scale numerical methods |
| deal.II | Finite element analysis |
| OpenFOAM | Computational fluid dynamics |

Potential uses:

- PDE solvers
- Engineering simulations
- Physics models
- Environmental modeling

---

# Machine Learning & AI

| Library | Purpose |
|---|---|
| XGBoost | Gradient boosting |
| ONNX Runtime | ML inference |
| TensorRT | GPU inference acceleration |
| dlib | Classical machine learning |
| mlpack | C++ machine learning |

Potential uses:

- Embedded prediction models
- Real-time scoring
- Native inference engines

---

# Data & Analytics

| Library | Purpose |
|---|---|
| DuckDB | Embedded analytics database |
| Apache Arrow | Columnar in-memory data |
| Parquet C++ | High-performance storage |
| SQLite | Lightweight persistence |
| HDF5 | Scientific datasets |

Potential uses:

- Large local analytics
- Persistent object storage
- Embedded databases
- Fast table interchange

---

# Parallel & High-Performance Computing

| Library | Purpose |
|---|---|
| OpenMP | Shared-memory parallelism |
| Intel TBB | Task-based parallelism |
| CUDA | GPU acceleration |
| OpenCL | Cross-platform GPU compute |
| MPI | Distributed computing |

Potential uses:

- Monte Carlo simulation
- Parallel optimization
- GPU acceleration
- Distributed compute grids

---

# Serialization & Infrastructure

| Library | Purpose |
|---|---|
| Boost.Serialization | Object persistence |
| Protobuf | Structured messaging |
| FlatBuffers | Zero-copy serialization |
| ZeroMQ | Messaging infrastructure |
| gRPC | Remote compute APIs |

Potential uses:

- Persistent object stores
- Async execution
- Worker communication
- Distributed compute services

---

# Build Requirements

| Component | Version |
|---|---|
| Visual Studio | 2022 |
| .NET SDK | 8.0 |
| Excel-DNA | 1.9+ |
| Excel | 64-bit recommended |
| Windows SDK | Current |

---

# Build Instructions

## Open Solution

```text
Core.sln
```

---

## Build Release x64

```text
Build → Release → x64
```

---

## Publish Output

The packed add-in appears under:

```text
bin/Release/net8.0-windows/win-x64/publish/
```

Typical output:

```text
ExcelBridge-AddIn64-packed.xll
NativeMath.dll
```

---

# Installation

## Step 1 — Open Excel Add-In Manager

```text
File → Options → Add-ins
```

---

## Step 2 — Browse to XLL

```text
Manage Excel Add-ins → Go → Browse
```

Select:

```text
ExcelBridge-AddIn64-packed.xll
```

---

## Step 3 — Trust the Add-In

Excel may prompt:

```text
Enable Add-In
```

or:

```text
Trust Publisher
```

Accept the prompts.

---

# Basic Usage

## Test Connectivity

In Excel:

```excel
=CPP_PING()
```

Expected result:

```text
Native bridge active
```

---

## Matrix Round Trip

Select a 3×3 range and enter:

```excel
=CPP_MATRIX_ROUNDTRIP(A1:C3)
```

This validates:

- Excel → C#
- C# → Native C++
- Native C++ → C#
- C# → Excel

---

## Matrix Transpose

```excel
=CPP_TRANSPOSE(A1:C3)
```

---

# Future Architectural Directions

Potential future enhancements include:

---

# Persistent Worker Process

Future architecture may evolve toward:

```text
Excel
  ↓
Lightweight Add-In
  ↓
Persistent C++ Worker
```

Possible communication mechanisms:

- Named pipes
- Localhost TCP
- Shared memory

Potential benefits:

- Long-running jobs
- Background execution
- Persistent QuantLib sessions
- Cached models
- Distributed compute

---

# Asynchronous Computation

Potential future workflow:

```text
Excel Function
    ↓
Submit Job
    ↓
Background Compute
    ↓
Store Result
    ↓
Refresh Excel
```

This becomes increasingly important for:

- Monte Carlo
- Calibration
- Optimization
- Large simulations

---

# Development Philosophy

The intended workflow is:

1. Build the baseline framework
2. Verify native connectivity
3. Add specialized libraries
4. Expose functionality to Excel
5. Add persistent objects
6. Add optimized workflows
7. Expand incrementally over time

The repository is intended to demonstrate:

- Modern Excel-DNA architecture
- Native C++ integration
- Persistent computational workflows
- High-performance Excel extensions
- Reusable interop patterns

rather than a single fixed end-user application.

---

# Recommended Next Steps

A reasonable evolution path is:

1. Stabilize matrix/object infrastructure
2. Add Eigen integration
3. Add serialization support
4. Add async/background jobs
5. Add QuantLib examples
6. Add persistent worker process
7. Add distributed execution

---

# License

Add your preferred license here:

- MIT
- Apache 2.0
- GPL
- Proprietary

---
