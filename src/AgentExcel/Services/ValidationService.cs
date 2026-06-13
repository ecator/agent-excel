using System.Runtime.InteropServices;

using AgentExcel.Providers;

using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel.Services;

public class ValidationService : ExcelServiceBase
{
    public ValidationService(IExcelConnectionProvider connectionProvider) : base(connectionProvider)
    {
    }

    public void SetListValidation(string workbookName, string sheetName, string rangeAddress, string formula)
    {
        ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Range? range = null;
            Excel.Validation? validation = null;
            try
            {
                wb = GetWorkbook(workbookName, createNew: true);
                ws = GetWorksheet(wb, sheetName);
                range = ws.Range[rangeAddress];
                validation = range.Validation;
                validation.Delete();
                validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop,
                    Excel.XlFormatConditionOperator.xlBetween, formula);
                validation.IgnoreBlank = true;
                validation.InCellDropdown = true;
            }
            finally
            {
                if (validation != null) Marshal.ReleaseComObject(validation);
                if (range != null) Marshal.ReleaseComObject(range);
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null) Marshal.ReleaseComObject(wb);
            }
        });
    }
}
