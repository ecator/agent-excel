using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

using AgentExcel.Models;
using AgentExcel.Providers;
using AgentExcel.Utils;

using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel.Services;

public class MacroService : ExcelServiceBase
{
    public MacroService(IExcelConnectionProvider connectionProvider) : base(connectionProvider)
    {
    }

    public object? RunMacro(string workbookName, string macroName, object[]? args)
    {
        return ExecuteWithRetry(() =>
        {
            var app = GetApp(createNew: false);
            if (app == null)
            {
                throw new Exception("Excel is not running. Please open Excel first.");
            }

            string targetMacro = macroName;
            if (!macroName.Contains('!'))
            {
                var wb = GetWorkbook(workbookName);
                try
                {
                    string wbName = wb.Name;
                    targetMacro = $"'{wbName}'!{macroName}";
                }
                finally
                {
                    SafeReleaseComObject(wb);
                }
            }

            if (args == null || args.Length == 0)
            {
                return app.Run(targetMacro);
            }

            object[] fullArgs = new object[args.Length + 1];
            fullArgs[0] = targetMacro;
            for (int i = 0; i < args.Length; i++)
            {
                fullArgs[i + 1] = JsonHelper.ConvertJsonValue(args[i])!;
            }

            return typeof(Excel.Application).InvokeMember("Run",
                System.Reflection.BindingFlags.InvokeMethod, null, app, fullArgs);
        });
    }

    public List<MacroInfo> ListMacros(string workbookName)
    {
        return ExecuteWithRetry(() =>
        {
            var app = GetApp(createNew: false);
            if (app == null)
            {
                throw new Exception("Excel is not running. Please open Excel first.");
            }

            var wb = GetWorkbook(workbookName);
            var resultList = new List<MacroInfo>();
            try
            {
                if (!wb.HasVBProject)
                {
                    return resultList;
                }

                object? vbProject = null;
                object? components = null;
                try
                {
                    vbProject = GetVbProject(wb);
                    components = ((dynamic)vbProject).VBComponents;
                    if (components != null)
                    {
                        int componentCount = ((dynamic)components).Count;
                        for (int i = 1; i <= componentCount; i++)
                        {
                            object? component = null;
                            object? codeModule = null;
                            try
                            {
                                component = ((dynamic)components).Item(i);
                                if (component == null) continue;

                                string componentName = ((dynamic)component).Name;
                                int componentTypeVal = ((dynamic)component).Type;
                                string moduleType = componentTypeVal switch
                                {
                                    1 => "Module",
                                    2 => "Class",
                                    3 => "Form",
                                    100 => "Document",
                                    _ => "Unknown"
                                };

                                codeModule = ((dynamic)component).CodeModule;
                                if (codeModule != null)
                                {
                                    int countOfLines = ((dynamic)codeModule).CountOfLines;
                                    if (countOfLines > 0)
                                    {
                                        object[] linesArgs = [1, countOfLines];
                                        string codeText = (codeModule.GetType().InvokeMember(
                                            "Lines",
                                            System.Reflection.BindingFlags.GetProperty,
                                            null,
                                            codeModule,
                                            linesArgs) as string) ?? string.Empty;
                                        var macros = ParseVbaProcedures(codeText, componentName, moduleType);
                                        resultList.AddRange(macros);
                                    }
                                }
                            }
                            finally
                            {
                                SafeReleaseComObject(codeModule);
                                SafeReleaseComObject(component);
                            }
                        }
                    }
                }
                finally
                {
                    SafeReleaseComObject(components);
                    SafeReleaseComObject(vbProject);
                }
            }
            finally
            {
                SafeReleaseComObject(wb);
            }

            return resultList;
        });
    }

    public void ExportModule(string workbookName, string moduleName, string outputFile)
    {
        if (string.IsNullOrWhiteSpace(moduleName))
        {
            throw new ArgumentException("Module name cannot be null or empty.", nameof(moduleName));
        }

        if (string.IsNullOrWhiteSpace(outputFile))
        {
            throw new ArgumentException("Output file path cannot be null or empty.", nameof(outputFile));
        }

        string fullPath = Path.GetFullPath(outputFile);
        string? directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            throw new DirectoryNotFoundException($"The directory '{directory}' does not exist.");
        }

        ExecuteWithRetry(() =>
        {
            var app = GetApp(createNew: false);
            if (app == null)
            {
                throw new Exception("Excel is not running. Please open Excel first.");
            }

            var wb = GetWorkbook(workbookName);
            try
            {
                if (!wb.HasVBProject)
                {
                    throw new Exception($"Workbook '{workbookName}' does not have a VBA project.");
                }

                object? vbProject = null;
                object? components = null;
                object? targetComponent = null;
                try
                {
                    vbProject = GetVbProject(wb);
                    components = ((dynamic)vbProject).VBComponents;
                    if (components == null)
                    {
                        throw new Exception($"Failed to access VBComponents in workbook '{workbookName}'.");
                    }

                    targetComponent = FindComponent(components, moduleName);
                    if (targetComponent == null)
                    {
                        throw new Exception($"Module '{moduleName}' not found in workbook '{workbookName}'.");
                    }

                    ((dynamic)targetComponent).Export(fullPath);
                }
                finally
                {
                    SafeReleaseComObject(targetComponent);
                    SafeReleaseComObject(components);
                    SafeReleaseComObject(vbProject);
                }
            }
            finally
            {
                SafeReleaseComObject(wb);
            }
        });
    }

    public string ImportModule(string workbookName, string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ArgumentException("File path cannot be null or empty.", nameof(filePath));
        }

        string fullPath = Path.GetFullPath(filePath);
        if (!File.Exists(fullPath))
        {
            throw new FileNotFoundException($"The file '{filePath}' does not exist.", filePath);
        }

        string targetModuleName = GetModuleNameFromFile(fullPath);

        return ExecuteWithRetry(() =>
        {
            var app = GetApp(createNew: false);
            if (app == null)
            {
                throw new Exception("Excel is not running. Please open Excel first.");
            }

            var wb = GetWorkbook(workbookName);
            try
            {
                if (!wb.HasVBProject)
                {
                    throw new Exception($"Workbook '{workbookName}' does not have a VBA project.");
                }

                object? vbProject = null;
                object? components = null;
                object? importedComponent = null;
                try
                {
                    vbProject = GetVbProject(wb);
                    components = ((dynamic)vbProject).VBComponents;
                    if (components == null)
                    {
                        throw new Exception($"Failed to access VBComponents in workbook '{workbookName}'.");
                    }

                    object? existingComponent = FindComponent(components, targetModuleName);
                    if (existingComponent != null)
                    {
                        SafeReleaseComObject(existingComponent);
                        throw new InvalidOperationException($"Module '{targetModuleName}' already exists in workbook '{workbookName}'. Please delete the existing module before importing.");
                    }

                    importedComponent = ((dynamic)components).Import(fullPath);
                    string importedName = ((dynamic)importedComponent).Name;
                    return importedName;
                }
                finally
                {
                    SafeReleaseComObject(importedComponent);
                    SafeReleaseComObject(components);
                    SafeReleaseComObject(vbProject);
                }
            }
            finally
            {
                SafeReleaseComObject(wb);
            }
        });
    }

    public void DeleteModule(string workbookName, string moduleName)
    {
        if (string.IsNullOrWhiteSpace(moduleName))
        {
            throw new ArgumentException("Module name cannot be null or empty.", nameof(moduleName));
        }

        ExecuteWithRetry(() =>
        {
            var app = GetApp(createNew: false);
            if (app == null)
            {
                throw new Exception("Excel is not running. Please open Excel first.");
            }

            var wb = GetWorkbook(workbookName);
            try
            {
                if (!wb.HasVBProject)
                {
                    throw new Exception($"Workbook '{workbookName}' does not have a VBA project.");
                }

                object? vbProject = null;
                object? components = null;
                object? targetComponent = null;
                try
                {
                    vbProject = GetVbProject(wb);
                    components = ((dynamic)vbProject).VBComponents;
                    if (components == null)
                    {
                        throw new Exception($"Failed to access VBComponents in workbook '{workbookName}'.");
                    }

                    targetComponent = FindComponent(components, moduleName);
                    if (targetComponent == null)
                    {
                        throw new Exception($"Module '{moduleName}' not found in workbook '{workbookName}'.");
                    }

                    int componentTypeVal = ((dynamic)targetComponent).Type;
                    if (componentTypeVal == 100)
                    {
                        throw new InvalidOperationException($"Cannot delete document component '{moduleName}' (e.g., Sheet or ThisWorkbook).");
                    }

                    ((dynamic)components).Remove(targetComponent);
                }
                finally
                {
                    SafeReleaseComObject(targetComponent);
                    SafeReleaseComObject(components);
                    SafeReleaseComObject(vbProject);
                }
            }
            finally
            {
                SafeReleaseComObject(wb);
            }
        });
    }

    private object? FindComponent(object components, string moduleName)
    {
        int componentCount = ((dynamic)components).Count;
        for (int i = 1; i <= componentCount; i++)
        {
            object? comp = null;
            try
            {
                comp = ((dynamic)components).Item(i);
                if (comp != null)
                {
                    string compName = ((dynamic)comp).Name;
                    if (compName.Equals(moduleName, StringComparison.OrdinalIgnoreCase))
                    {
                        object found = comp;
                        comp = null;
                        return found;
                    }
                }
            }
            finally
            {
                SafeReleaseComObject(comp);
            }
        }
        return null;
    }

    internal static string GetModuleNameFromFile(string filePath)
    {
        try
        {
            var ansiEncoding = Encoding.GetEncoding(0);
            string content = File.ReadAllText(filePath, ansiEncoding);
            var match = Regex.Match(content, @"^\s*Attribute\s+VB_Name\s*=\s*""([^""]+)""", RegexOptions.IgnoreCase | RegexOptions.Multiline);
            if (match.Success)
            {
                return match.Groups[1].Value.Trim();
            }
        }
        catch
        {
            // If file reading fails or format differs, fallback to filename without extension
        }

        return Path.GetFileNameWithoutExtension(filePath);
    }

    private static object GetVbProject(Excel.Workbook wb)
    {
        try
        {
            // Access VBProject dynamically to avoid compile-time dependencies on Microsoft.Vbe.Interop
            object? vbProject = ((dynamic)wb).VBProject;
            if (vbProject == null)
            {
                throw new Exception($"Failed to access VBA project in workbook '{wb.Name}'.");
            }
            return vbProject;
        }
        catch (COMException ex) when ((uint)ex.ErrorCode == 0x800A03EC || ex.Message.Contains("programmatic access") || ex.Message.Contains("not trusted"))
        {
            throw new Exception("Access to the VBA project is not trusted. Please enable 'Trust access to the VBA project object model' in Excel's Trust Center settings.");
        }
    }

    internal static List<MacroInfo> ParseVbaProcedures(string codeText, string moduleName, string moduleType)
    {
        var list = new List<MacroInfo>();
        if (string.IsNullOrEmpty(codeText))
        {
            return list;
        }

        // Normalize VBA line continuations: space + underscore followed by newline
        string normalizedCode = Regex.Replace(codeText, @"_\s*[\r\n]+", " ", RegexOptions.IgnoreCase);

        // Pattern to match Sub or Function declarations
        // Format: [Public|Private|Friend] [Static] Sub/Function Name ( [Params] )
        string pattern = @"^\s*(?:(?:Public|Private|Friend|Static)\s+)*(Sub|Function)\s+(\w+)\s*\((.*?)\)";
        var matches = Regex.Matches(normalizedCode, pattern, RegexOptions.IgnoreCase | RegexOptions.Multiline);

        foreach (Match match in matches)
        {
            string type = match.Groups[1].Value; // "Sub" or "Function"
            string name = match.Groups[2].Value; // procedure name
            string paramsText = match.Groups[3].Value.Trim(); // parameters list

            var parameters = new List<string>();
            if (!string.IsNullOrEmpty(paramsText))
            {
                var parts = paramsText.Split(',');
                foreach (var p in parts)
                {
                    parameters.Add(p.Trim());
                }
            }

            list.Add(new MacroInfo(moduleName, moduleType, name, type, parameters));
        }

        return list;
    }
}
