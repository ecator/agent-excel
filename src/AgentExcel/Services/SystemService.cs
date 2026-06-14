using AgentExcel.Models;
using AgentExcel.Providers;

using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel.Services;

public class SystemService : ExcelServiceBase
{
    public SystemService(IExcelConnectionProvider connectionProvider) : base(connectionProvider)
    {
    }

    public void Shutdown()
    {
        ConnectionProvider.Dispose();
    }

    public void Calculate()
    {
        ExecuteWithRetry(() =>
        {
            var app = GetApp(createNew: true);
            if (app != null)
            {
                app.Calculate();
            }
        });
    }

    public List<QueryInfo> ListQueries(string workbookName)
    {
        return ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            dynamic? queries = null;
            var result = new List<QueryInfo>();
            try
            {
                wb = GetWorkbook(workbookName);
                dynamic dynWb = wb;
                queries = dynWb.Queries;
                if (queries != null)
                {
                    foreach (dynamic q in queries)
                    {
                        result.Add(new QueryInfo(q.Name, q.Formula, q.Description));
                        SafeReleaseComObject(q);
                    }
                }
                return result;
            }
            finally
            {
                SafeReleaseComObject(queries);
                SafeReleaseComObject(wb);
            }
        });
    }

    public void AddOrUpdateQuery(string workbookName, string name, string formula, string? description)
    {
        ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            dynamic? queries = null;
            dynamic? query = null;
            try
            {
                wb = GetWorkbook(workbookName, createNew: true);
                dynamic dynWb = wb;
                queries = dynWb.Queries;
                if (queries == null) return;

                bool exists = false;
                foreach (dynamic q in queries)
                {
                    string qName = q.Name;
                    if (qName.Equals(name, StringComparison.OrdinalIgnoreCase))
                    {
                        query = q;
                        exists = true;
                        break;
                    }
                    SafeReleaseComObject(q);
                }

                if (exists && query != null)
                {
                    query!.Formula = formula;
                    if (description != null) query!.Description = description;
                }
                else
                {
                    queries.Add(name, formula, description);
                }
            }
            finally
            {
                SafeReleaseComObject(query);
                SafeReleaseComObject(queries);
                SafeReleaseComObject(wb);
            }
        });
    }

    public void DeleteQuery(string workbookName, string name)
    {
        ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            dynamic? queries = null;
            try
            {
                wb = GetWorkbook(workbookName, createNew: true);
                dynamic dynWb = wb;
                queries = dynWb.Queries;
                if (queries == null) return;

                foreach (dynamic q in queries)
                {
                    string qName = q.Name;
                    if (qName.Equals(name, StringComparison.OrdinalIgnoreCase))
                    {
                        q.Delete();
                        SafeReleaseComObject(q);
                        break;
                    }
                    SafeReleaseComObject(q);
                }
            }
            finally
            {
                SafeReleaseComObject(queries);
                SafeReleaseComObject(wb);
            }
        });
    }

    public void RefreshAllDataConnections(string workbookName)
    {
        ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            try
            {
                wb = GetWorkbook(workbookName);
                wb.RefreshAll();
            }
            finally
            {
                SafeReleaseComObject(wb);
            }
        });
    }

    public void RefreshModel(string workbookName)
    {
        ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Model? model = null;
            try
            {
                wb = GetWorkbook(workbookName);
                model = wb.Model;
                model.Refresh();
            }
            finally
            {
                SafeReleaseComObject(model);
                SafeReleaseComObject(wb);
            }
        });
    }
}
