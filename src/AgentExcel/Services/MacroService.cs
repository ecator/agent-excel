
using System.Runtime.InteropServices;
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
                    // Access VBProject dynamically to avoid compile-time dependencies on Microsoft.Vbe.Interop
                    vbProject = ((dynamic)wb).VBProject;
                }
                catch (COMException ex) when ((uint)ex.ErrorCode == 0x800A03EC || ex.Message.Contains("programmatic access") || ex.Message.Contains("not trusted"))
                {
                    throw new Exception("Access to the VBA project is not trusted. Please enable 'Trust access to the VBA project object model' in Excel's Trust Center settings.");
                }

                if (vbProject != null)
                {
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

                SafeReleaseComObject(components);
                SafeReleaseComObject(vbProject);
            }
            finally
            {
                SafeReleaseComObject(wb);
            }

            return resultList;
        });
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
