# Agent-Excel: Developer and AI Programming Guide

## 1. Project Positioning
`AgentExcel` is a local Windows CLI tool specifically designed for AI Agents.
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

## 4. CLI Interface Standards
AI or users interact with the tool's lifecycle via the command line:

1.  `AgentExcel.exe start`
    * Behavior: Probes if an instance is already running. If not, retrieves an idle local system port (Port 0), silently starts `AgentExcel.exe --run-server <port>` in the background, and immediately returns console control.
2.  `AgentExcel.exe status`
    * Behavior: Probes the background PID and port, and sends a GET request to `http://127.0.0.1:<port>/status`.
    * Output: Prints a pretty JSON report containing PID, port, Excel connection status, and the active workbook name.
3.  `AgentExcel.exe stop`
    * Behavior: Probes the port, sends a POST request to `http://127.0.0.1:<port>/exit`, triggering the server to safely release COM objects and self-terminate.

## 5. HTTP API Route Standards (Daemon side)
Once the background Web service starts, it exposes the following local endpoints for AI to call:

* **`GET /status`**: Returns the health status of the daemon and the Excel attachment status (JSON).
* **`POST /exit`**: Upon receiving the request, calls `Marshal.ReleaseComObject` to release the Application handle and executes `IHostApplicationLifetime.StopApplication()` to gracefully exit.

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
* **Verification Scope**: After code modifications, you only need to run `dotnet build` and tests to verify the changes. You do not need to run the compiled `.exe` executable to verify the execution results. Generally, verifying results is done by the user, unless the user explicitly specifies that you need to verify the results through the entire flow yourself.