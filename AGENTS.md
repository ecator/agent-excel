# Agent-Excel: Developer and AI Programming Guide

## 1. Project Positioning
`AgentExcel` is a local HTTP API service (accompanied by a lightweight CLI client) designed to control Microsoft Excel via RESTful APIs, specifically tailored for AI Agents.
Its core mission is: **to enable AI to read and write Excel sheets currently open and operated by the user at high frequencies and extreme speeds via simple HTTP APIs.**

## 2. Core Architecture Design: Client-Daemon Mode
This project abandons the traditional "launch a process for each call" CLI mode and adopts a **local C/S (Client-Daemon) architecture** to support millisecond-level dense API calls.

* **Daemon**: A background service that runs persistently, holds the Excel COM object handles long-term, and exposes local HTTP endpoints (Minimal API).
* **Client**: An extremely lightweight CLI trigger used to wake up the Daemon, probe its status, or send termination commands. The actual AI business logic (e.g., Python scripts) interacts with the Daemon as the true client via HTTP.
* **Stateless Service Discovery**: No local temporary `.txt` files are used to store ports. The CLI client retrieves the background PID using `Process.GetProcessesByName` and silently executes `netstat -ano` with regex matching to extract the dynamically allocated port.

## 3. Directory Structure and Project Standards

### 3.1 Directory Structure
* **`src/`**: Core source code directory.
  * **`src/AgentExcel/`**: Main application directory (contains CLI startup logic and the Daemon Web service).
* **`tests/`**: Unit and integration test directory.
  * **`tests/AgentExcel.Tests/`**: Main test project.
* **`lib/`**: Directory for local assembly dependencies (e.g., Excel/Word/PowerPoint COM Interop DLLs).
* **`scripts/`**: Directory for scripts.
* **`AgentExcel.slnx`**: Visual Studio solution definition file.

### 3.2 Tech Stack Standards
* **Language/Framework**: C# 13 / .NET 10 (Windows platform only).
* **Core Dependency**: `Microsoft.Office.Interop.Excel` (COM Interop, depending on local `lib/Microsoft.Office.Interop.Excel.dll`).
* **Web Framework**: **Controllers are strictly prohibited.** Minimal APIs must be used in `Program.cs` top-level statements or the traditional `Main` function to write routes (`app.MapGet`, `app.MapPost`).
* **Compilation & Publishing Standards**: Must be compiled as a single-file, self-contained executable to run independently of the target environment.
    ```xml
    <Project Sdk="Microsoft.NET.Sdk">
      <PropertyGroup>
        <TargetFramework>net10.0-windows7.0</TargetFramework>
        <Nullable>enable</Nullable>
        <ImplicitUsings>enable</ImplicitUsings>
        <OutputType>Exe</OutputType>
        <PublishSingleFile>true</PublishSingleFile>
        <SelfContained>true</SelfContained>
        <RuntimeIdentifier>win-x64</RuntimeIdentifier>
      </PropertyGroup>
      <ItemGroup>
        <FrameworkReference Include="Microsoft.AspNetCore.App" />
      </ItemGroup>
    </Project>
    ```

### 3.3 Code Style Standards
To ensure consistent styling and leverage modern C# language features, C# code in this project must follow these rules:
*   **Formatting**: Use 4 spaces for indentation, UTF-8 encoding, and `CRLF` for line endings (tailored for Windows).
*   **Modern C# features**:
    *   Use **file-scoped namespaces** to reduce nesting.
    *   Prefer `var` when the type is apparent or built-in, but otherwise use explicit types if readability is improved.
    *   Leverage C# 12+ features such as **collection expressions** (e.g., `[1, 2, 3]`), **primary constructors**, and **implicit object creation** (`new()`).
    *   Prefer pattern matching (`is null`, `is not null`) for null checks.
*   **Naming Conventions**:
    *   **PascalCase**: Used for classes, structs, enums, interfaces (prefixed with `I`), methods, properties, public fields, and constants (`const` and `static readonly`). *Note: constants do not use SCREAMING_SNAKE_CASE in .NET in order to align with public API surface consistency, avoid macro legacy, and keep API modifications seamless.*
    *   **camelCase**: Used for parameters and local variables.
    *   **Private/Internal Instance Fields**: Prefixed with an underscore (`_camelCase`).
    *   **Private/Internal Static Fields**: Prefixed with `s_` (`s_camelCase`).

## 4. CLI Interface Standards
AI or users interact with the tool's lifecycle via the command line:

1.  `AgentExcel.exe start`
    * Behavior: Probes if an instance is already running. If not, retrieves an idle local system port (Port 0), silently starts `AgentExcel.exe --run-server <port>` in the background, and immediately returns console control.
2.  `AgentExcel.exe status`
    * Behavior: Probes the background PID and port, and sends a GET request to `http://127.0.0.1:<port>/status`.
    * Output: Prints the YAML report containing PID, port, Excel connection status, and the active workbook name.
3.  `AgentExcel.exe stop`
    * Behavior: Probes the port, sends a POST request to `http://127.0.0.1:<port>/exit`, triggering the server to safely release COM objects and self-terminate.

## 5. HTTP API Route and Response Format Standards (Daemon side)
To ensure maximum readability and parser simplicity for both human developers and AI agents, all API endpoints strictly return plain text formatted in **YAML** (`text/yaml; charset=utf-8`) or **Plain Text** (`text/plain; charset=utf-8`) by default, rather than JSON.

Once the background Web service starts, it exposes the following local endpoints:

* **`GET /status`**: Returns the health status of the daemon and the Excel attachment status (YAML format).
* **`POST /exit`**: Upon receiving the request, calls `Marshal.ReleaseComObject` to release the Application handle and executes `IHostApplicationLifetime.StopApplication()` to gracefully exit (returns shutdown status in YAML format).

## 6. ⚠️ Absolute Red Lines: COM Operations and Memory Leak Prevention
When writing C# code that operates on Excel COM, the AI must strictly follow these rules to prevent leftover `excel.exe` zombie processes or program crashes:

1.  **Attach Instead of Create**: Prefer using `Marshal.GetActiveObject("Excel.Application")` to obtain the Excel instance currently opened by the user.
2.  **Never Call Quit**: **Never** call `excelApp.Quit()` in the code. Doing so will force-close the user's active window and lead to critical user accidents. Simply disconnect.
3.  **Double-Dot Principle (Never use two dots)**: Do not chain multiple properties or methods of COM objects using two or more consecutive dots.
    * ❌ Incorrect: `var name = excelApp.ActiveWorkbook.ActiveSheet.Name;` (creates implicit intermediate COM objects that cannot be released).
    * ✅ Correct: Declare variables separately, e.g., `Excel.Workbook wb = excelApp.ActiveWorkbook;`.
4.  **Explicit Manual Release**: All explicitly declared COM objects (Application, Workbook, Worksheet, Range) must be freed using `Marshal.ReleaseComObject(obj)` before completing usage or returning from HTTP endpoints.

## 7. Exception Handling Guide
* **RPC_E_CALL_REJECTED (0x80010001)**: If this exception is caught, it typically means the user is editing a cell (Cell Edit Mode), causing Excel to suspend the COM channel. The API should return a clear error message prompting the AI to request the user to press Enter to exit the cell editing mode.

## 8. Communication and Language Policy
* **User Language Matching**: Identify the user's language and communicate with them using the same language.
* **Code & Docs in English**: All written code and project documentation (including code comments, API schemas, etc.) must be written in English.
* **Explanations in User Language**: Explain code logic, troubleshooting, or implementation details to the user using the user's language, even though the code and comments themselves are in English.

## 9. Build and Verification Policy
* **Build Script Requirement**: Do NOT run `dotnet build` directly to compile the project, as the executable might be occupied by running processes or during testing, causing compilation failures. You MUST use the `scripts/build.ps1` script, which stops any potentially running daemon processes before compiling.
* **Test Script Requirement**: Do NOT run `dotnet test` directly to execute tests. You MUST use the `scripts/test.ps1` script instead. This script transparently forwards all arguments to `dotnet test`, meaning options like `--filter` are fully supported.
* **Test Pre-requisite**: Before running `scripts/test.ps1` (or any testing commands), you MUST execute `scripts/build.ps1` once to ensure the latest code is successfully compiled and any blocking processes are terminated.
* **Verification Scope**: After code modifications, you only need to compile with the build script and run tests via `scripts/test.ps1` to verify the changes. You do not need to run the compiled `.exe` executable to verify the execution results. Generally, verifying results is done by the user, unless the user explicitly specifies that you need to verify the results through the entire flow yourself.

## 10. Test Case Writing Standards
To ensure software quality, correctness, and maintainability, all tests must follow these industry best practices:

### 10.1 Framework and Tooling
* **Test Framework**: Use **NUnit 4** (as configured in the test project).
* **Assertion Style**: Use the NUnit **Constraint Model** (e.g., `Assert.That(actual, Is.EqualTo(expected))`) instead of the classic assertion style (e.g., `Assert.AreEqual(expected, actual)`).

### 10.2 Naming Conventions
Test classes and methods must follow consistent naming rules to be self-documenting:
* **Test Class**: Match the class under test, suffixed with `Tests`. For example, `CliCommandHandlerTests` tests `CliCommandHandler`.
* **Test Method**: Use the `UnitOfWork_StateUnderTest_ExpectedBehavior` pattern.
  * *Example*: `Start_WhenInstanceAlreadyRunning_ReturnsError`
  * *Example*: `Exit_WhenCalled_ReleasesComObjectsAndTerminates`

### 10.3 Directory Structure and Base Class
* **Folder Mirroring**: The directory structure of the test project must match the project under test.
  * *Example*: A file in `src/AgentExcel/Commands/` must have its tests in `tests/AgentExcel.Tests/Commands/`.
* **Test Base Class**: Test classes should inherit from the `BaseTests` class to leverage shared helper methods, configuration, and teardown logic.
* **Test Data Directories**: `BaseTests` provides properties to dynamically resolve absolute paths of test data folders:
  * `TestDataPath`: Points to the absolute path of the `tests/TestData` directory containing static test files (e.g. template worksheets, dummy files).
  * `TestDataTempPath`: Points to the absolute path of the `tests/TestData/temp` directory. This is used for creating, saving, and cleaning up temporary test output files (it is automatically created if it doesn't exist).
  * **Rule**: Always use `TestDataPath` or `TestDataTempPath` in test cases to load or save test worksheets instead of hardcoding relative paths or outputting directly to `AppDomain.CurrentDomain.BaseDirectory`.

### 10.4 Test Structure (AAA Pattern)
Every test should clearly structure its execution into three distinct phases separated by blank lines:
1. **Arrange**: Set up the environment, mock objects, inputs, and preconditions.
2. **Act**: Invoke the action or method under test.
3. **Assert**: Verify that the outcomes match expectations.

```csharp
[Test]
public void Calculate_WithValidInputs_ReturnsExpectedResult()
{
    // Arrange
    var calculator = new Calculator();
    int a = 5;
    int b = 10;

    // Act
    int result = calculator.Add(a, b);

    // Assert
    Assert.That(result, Is.EqualTo(15));
}
```

### 10.5 Test Isolation and Independence
* **No Shared State Mutation**: Tests must run independently. No test should rely on the state left by another test or the order in which tests run.
* **Cleanup**: Use `[TearDown]` or `[OneTimeTearDown]` to clean up temporary files, registry entries, or process instances.
* **Mocking**: Use mocking (e.g., Moq or NSubstitute) to isolate the unit under test from external dependencies like file systems, network calls, or CLI processes.

### 10.6 COM Release in Integration Tests
If an integration test interacts with actual Office/Excel COM objects:
* Always wrap the COM operations in `try-finally` blocks.
* Clean up all instantiated COM objects explicitly using `Marshal.ReleaseComObject` in the `finally` block or `[TearDown]`.
* Ensure that no Excel background processes remain running after the test run.

### 10.7 Categorization
* Use the `[Category]` attribute to classify tests:
  * `[Category("Unit")]` for fast, lightweight in-memory unit tests that have no external dependencies.
  * `[Category("Integration")]` for tests that interact with external services, files, or processes.
  * `[Category("COM")]` for tests requiring Excel or other Office applications to be installed and running.