namespace AgentExcel.Models;

/// <summary>
/// Information about an Excel workbook.
/// </summary>
/// <param name="Name">The name of the workbook file.</param>
/// <param name="Path">The full file path of the workbook.</param>
/// <param name="Sheets">The list of sheet names within the workbook, or null if sheets are not loaded.</param>
public record WorkbookInfo(string Name, string Path, List<string>? Sheets);
