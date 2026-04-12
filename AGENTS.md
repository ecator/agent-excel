# Agent-Excel: 开发者与 AI 编程指南

## 1. 项目定位
`AgentExcel` 是一个专为 AI 代理 (AI Agent) 设计的 Windows 本地命令行工具。
它的核心使命是：**让 AI 能够通过极简的 HTTP API，高频、极速地读写用户当前正在打开和操作的 Excel 表格。**

## 2. 核心架构设计：Client-Daemon 模式
本项目摒弃了传统的“每次调用启动一次进程”的 CLI 模式，采用 **本地 C/S (Client-Daemon) 架构** 以支持毫秒级的密集 API 调用。

* **Daemon (守护进程)**：常驻后台，长期持有 Excel 的 COM 对象句柄，并暴露本地 HTTP 接口 (Minimal API)。
* **Client (客户端)**：极轻量的 CLI 触发器，用于唤醒 Daemon、探活、或发送终结指令。AI 业务逻辑（如 Python 脚本）作为真正的 Client 通过 HTTP 与 Daemon 交互。
* **无状态发现 (Service Discovery)**：不使用任何本地 `.txt` 临时文件存储端口。CLI 客户端通过 `Process.GetProcessesByName` 获取后台 PID，并静默执行 `netstat -ano` 正则匹配出动态分配的可用端口。

## 3. 技术栈与工程规范
* **语言框架**：C# 12 / .NET 8 (仅限 Windows 平台)。
* **核心依赖**：`Microsoft.Office.Interop.Excel` (COM 互操作)。
* **Web 框架**：**绝对禁止使用 Controllers**。必须使用 `Program.cs` 顶级语句或传统的 `Main` 函数结合 **Minimal API** 编写路由 (`app.MapGet`, `app.MapPost`)。
* **编译发布规范**：必须编译为单文件、自包含的可执行文件，以便脱离环境独立运行。
    ```xml
    <Project Sdk="Microsoft.NET.Sdk.Web">
      <PropertyGroup>
        <TargetFramework>net8.0-windows</TargetFramework>
        <OutputType>Exe</OutputType>
        <PublishSingleFile>true</PublishSingleFile>
        <SelfContained>true</SelfContained>
        <RuntimeIdentifier>win-x64</RuntimeIdentifier>
        </PropertyGroup>
    </Project>
    ```

## 4. CLI 接口规范
AI 或用户通过命令行与工具进行基础生命周期交互：

1.  `AgentExcel.exe start`
    * 行为：探测是否已有实例。若无，获取本地系统空闲端口 (Port 0)，在后台静默启动 `AgentExcel.exe --run-server <port>`，并立刻返回控制台权限。
2.  `AgentExcel.exe status`
    * 行为：探测后台 PID 与端口，向 `http://127.0.0.1:<port>/status` 发送 GET 请求。
    * 输出：打印 JSON 解析后的美化报告（包含 PID、端口、Excel 接管状态、活动工作簿名称）。
3.  `AgentExcel.exe stop`
    * 行为：探测端口，向 `http://127.0.0.1:<port>/exit` 发送 POST 请求，触发服务端安全释放 COM 对象并自我结束进程。

## 5. HTTP API 路由规范 (Daemon 端)
后台 Web 服务启动后，暴露以下本地接口供 AI 高频调用：

* **`GET /status`**: 返回当前守护进程的健康状态与 Excel 挂载情况 (JSON)。
* **`POST /exit`**: 收到请求后，调用 `Marshal.ReleaseComObject` 释放 Application 句柄，并执行 `IHostApplicationLifetime.StopApplication()` 优雅退出。

## 6. ⚠️ 绝对红线：COM 操作与内存防漏指南
在编写任何操作 Excel COM 的 C# 代码时，AI 必须严格遵守以下原则，否则将导致 `excel.exe` 幽灵进程残留或程序崩溃：

1.  **接管而非创建**：优先使用 `Marshal.GetActiveObject("Excel.Application")` 获取用户已打开的实例。
2.  **严禁调用 Quit**：**永远不要**在代码中调用 `excelApp.Quit()`，否则会强制关闭用户正在观看的界面，导致严重事故。只需要断开连接即可。
3.  **双点原则 (Never use two dots)**：严禁在 COM 对象链式调用中使用两个以上的点号。
    * ❌ 错误：`var name = excelApp.ActiveWorkbook.ActiveSheet.Name;` (会产生无法释放的隐式中间对象)。
    * ✅ 正确：分别声明变量，如 `Excel.Workbook wb = excelApp.ActiveWorkbook;`。
4.  **绝对手动释放**：所有显式声明的 COM 对象（Application, Workbook, Worksheet, Range），在使用完毕或 HTTP 接口返回前，必须调用 `Marshal.ReleaseComObject(obj)` 释放。

## 7. 异常处理指导
* **RPC_E_CALL_REJECTED (0x80010001)**：如果捕获到此异常，通常是因为用户正在双击编辑某个单元格（处于 Cell Edit Mode），导致 COM 通道被 Excel 挂起。API 应返回明确的错误信息提示，告知 AI 要求用户按回车退出编辑状态。