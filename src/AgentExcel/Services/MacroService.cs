
using AgentExcel.Providers;

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
            Array.Copy(args, 0, fullArgs, 1, args.Length);

            return typeof(Excel.Application).InvokeMember("Run",
                System.Reflection.BindingFlags.InvokeMethod, null, app, fullArgs);
        });
    }
}
