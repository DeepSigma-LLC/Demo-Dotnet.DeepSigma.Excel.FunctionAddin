# Demo-Dotnet.DeepSigma.Excel.FunctionAddin

A demo solution by [DeepSigma](https://github.com/DeepSigma-LLC) that shows how to build, test, and deploy a custom **Excel Function Add-In** using [Excel-DNA](https://excel-dna.net/) and **.NET 8**.

---

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Solution Structure](#solution-structure)
- [Prerequisites](#prerequisites)
- [Getting Started](#getting-started)
  - [Clone the Repository](#clone-the-repository)
  - [Build the Solution](#build-the-solution)
  - [Run the Add-In in Excel](#run-the-add-in-in-excel)
- [Custom Excel Functions](#custom-excel-functions)
  - [AddTwo](#addtwo)
  - [SayHello](#sayhello)
  - [ReturnArray](#returnarray)
  - [MixedOutput](#mixedoutput)
- [IntelliSense Support](#intellisense-support)
- [Testing](#testing)
- [Key NuGet Packages](#key-nuget-packages)
- [How It Works](#how-it-works)
- [License](#license)

---

## Overview

This project demonstrates how to create a **custom Excel Function Add-In** using the Excel-DNA framework with C# and .NET 8. Excel-DNA allows .NET code to be hosted inside Microsoft Excel, exposing static C# methods as native Excel worksheet functions — no COM interop or VSTO required.

The solution also includes a separate **unit test project** (xUnit) that tests the function logic in isolation, completely independent of Excel.

---

## Features

- ✅ Custom Excel worksheet functions written in C#
- ✅ Functions grouped under a `DeepSigma` category in Excel's function wizard
- ✅ IntelliSense descriptions and argument metadata surfaced in Excel
- ✅ Support for returning scalar values, strings, and 2D arrays
- ✅ Demonstrates `ExcelError` and `ExcelEmpty` return values
- ✅ Add-in lifecycle management (`AutoOpen` / `AutoClose`)
- ✅ Unit tests with xUnit that run without requiring Excel

---

## Solution Structure

```
Demo-Dotnet.DeepSigma.Excel.FunctionAddin/
│
├── Demo-DeepSigma.Excel.FunctionAddin/          # Main add-in project
│   ├── AddIn.cs                                 # Add-in entry point (IExcelAddIn)
│   ├── MyExcelFunctions.cs                      # Custom Excel functions
│   ├── DeepSigma.Excel.FunctionAddin.csproj     # Project file
│   └── Properties/
│       └── launchSettings.json                  # Debug launch settings (opens Excel)
│
├── DeepSigma.Excel.FunctionAddin.Test/          # Unit test project
│   ├── MyExcelFunctionsTest.cs                  # xUnit tests for the functions
│   └── DeepSigma.Excel.FunctionAddin.Test.csproj
│
├── README.md
└── LICENSE
```

---

## Prerequisites

| Requirement | Version |
|---|---|
| [.NET SDK](https://dotnet.microsoft.com/en-us/download/dotnet/8.0) | 8.0 or later |
| [Visual Studio](https://visualstudio.microsoft.com/) | 2022 or later (Community edition is fine) |
| Microsoft Excel | 2010 or later (32-bit or 64-bit) |

> **Note:** The project targets `net8.0-windows` and requires a Windows environment with Excel installed to run the add-in. The unit tests can be run on any machine with .NET 8.

---

## Getting Started

### Clone the Repository

```bash
git clone https://github.com/DeepSigma-LLC/Demo-Dotnet.DeepSigma.Excel.FunctionAddin.git
cd Demo-Dotnet.DeepSigma.Excel.FunctionAddin
```

### Build the Solution

Open the solution in Visual Studio and build it (**Build → Build Solution**, or `Ctrl+Shift+B`), or use the .NET CLI:

```bash
dotnet build
```

The Excel-DNA build step automatically packages the output into two `.xll` files in the `bin\Debug\net8.0-windows\` folder:

| File | Architecture |
|---|---|
| `Demo-DeepSigma.Excel.FunctionAddin-AddIn.xll` | 32-bit Excel |
| `Demo-DeepSigma.Excel.FunctionAddin-AddIn64.xll` | 64-bit Excel |

### Run the Add-In in Excel

**Option 1 — Debug from Visual Studio (recommended):**

Press **F5**. The `launchSettings.json` is configured to launch Excel with the `.xll` add-in loaded automatically. This allows breakpoints and debugging in real time.

**Option 2 — Load manually in Excel:**

1. Open Excel.
2. Go to **File → Options → Add-ins**.
3. In the **Manage** dropdown, select **Excel Add-ins** and click **Go...**.
4. Click **Browse** and navigate to the `bin\Debug\net8.0-windows\` output folder.
5. Select `Demo-DeepSigma.Excel.FunctionAddin-AddIn64.xll` (for 64-bit Excel) and click **OK**.

The custom functions will now be available in any worksheet.

---

## Custom Excel Functions

All functions are grouped under the **DeepSigma** category in Excel's Insert Function dialog (`Shift+F3`).

### AddTwo

Adds two numbers together and returns the sum.

```
=AddTwo(x, y)
```

| Parameter | Type | Description |
|---|---|---|
| `x` | `double` | The first number |
| `y` | `double` | The second number |

**Example:**
```
=AddTwo(10, 25)  →  35
```

---

### SayHello

Returns a greeting string for a given name.

```
=SayHello(name)
```

| Parameter | Type | Description |
|---|---|---|
| `name` | `string` | The name of the person to greet |

**Example:**
```
=SayHello("Alice")  →  "Hello Alice"
```

---

### ReturnArray

Returns a 1-row × 3-column array of numbers. Must be entered as an **array formula** or used with Excel's dynamic array spill behavior.

```
=ReturnArray()
```

**Output:**

| Col 1 | Col 2 | Col 3 |
|---|---|---|
| 1 | 2 | 3 |

**Example (dynamic array):**
Select a cell and enter `=ReturnArray()`. The values spill into adjacent cells automatically in modern Excel.

---

### MixedOutput

Returns a 3-row × 3-column array with mixed data types, including strings, numbers, `#N/A` errors, and empty cells. This function demonstrates how Excel-DNA handles `ExcelError` and `ExcelEmpty` return values.

```
=MixedOutput()
```

**Output:**

| Ticker | Price | Status |
|---|---|---|
| MSFT | 425.10 | OK |
| BAD | #N/A | *(empty)* |

---

## IntelliSense Support

This add-in uses the **ExcelDna.IntelliSense** package to provide real-time function descriptions and argument hints directly inside Excel's formula bar — the same experience as built-in Excel functions.

The `AddIn.cs` class manages the IntelliSense server lifecycle:

```csharp
public void AutoOpen()  => IntelliSenseServer.Install();
public void AutoClose() => IntelliSenseServer.Uninstall();
```

Function and argument descriptions are defined using the `[ExcelFunction]` and `[ExcelArgument]` attributes on each method.

---

## Testing

The `DeepSigma.Excel.FunctionAddin.Test` project contains unit tests using **xUnit v3**. The function logic is tested in pure .NET — no Excel instance is required.

### Run Tests

**From Visual Studio:**

Open **Test Explorer** (`View → Test Explorer`) and click **Run All**.

**From the CLI:**

```bash
dotnet test
```

### Test Cases

| Test | Description |
|---|---|
| `TestAddTwo` | Verifies `AddTwo(2, 3)` returns `5` |
| `TestSayHello` | Verifies `SayHello("Alice")` returns `"Hello Alice"` |
| `TestReturnArray` | Verifies the returned array contains `{1, 2, 3}` |

---

## Key NuGet Packages

| Package | Version | Purpose |
|---|---|---|
| [ExcelDna.AddIn](https://www.nuget.org/packages/ExcelDna.AddIn) | 1.9.0 | Core Excel-DNA framework; packages the project as an `.xll` add-in |
| [ExcelDna.IntelliSense](https://www.nuget.org/packages/ExcelDna.IntelliSense) | 1.9.0 | Provides IntelliSense support for custom functions in Excel |
| [xunit.v3](https://www.nuget.org/packages/xunit.v3) | 3.2.2 | xUnit v3 test framework |
| [xunit.runner.visualstudio](https://www.nuget.org/packages/xunit.runner.visualstudio) | 3.1.5 | Visual Studio Test Explorer adapter for xUnit |
| [Microsoft.NET.Test.Sdk](https://www.nuget.org/packages/Microsoft.NET.Test.Sdk) | 18.6.0 | Test host infrastructure required by the VS test runner |

---

## How It Works

Excel-DNA works by hosting the .NET runtime inside a native `.xll` Excel add-in. Here is the high-level flow:

1. **Build time:** The `ExcelDna.AddIn` MSBuild tasks bundle your compiled `.dll`, its dependencies, and the Excel-DNA loader into a self-contained `.xll` file.
2. **Load time:** When Excel loads the `.xll`, it starts the .NET runtime and calls `AddIn.AutoOpen()`, which installs the IntelliSense server.
3. **Registration:** Excel-DNA scans all loaded assemblies for `public static` methods decorated with `[ExcelFunction]` and registers them as native Excel worksheet functions.
4. **Invocation:** When a user calls `=AddTwo(1,2)` in a cell, Excel-DNA marshals the arguments from Excel types to .NET types, invokes the static method, and marshals the return value back to Excel.
5. **Unload:** When Excel exits or the add-in is disabled, `AddIn.AutoClose()` is called to clean up the IntelliSense server.

---

## License

This project is licensed under the terms of the [LICENSE](LICENSE) file in the root of this repository.

