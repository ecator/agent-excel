namespace AgentExcel.Models;

/// <summary>
/// Information about an Excel worksheet.
/// </summary>
/// <param name="Name">The name of the worksheet.</param>
/// <param name="Color">The hex color code of the worksheet tab (e.g., '#FF0000'), or 'None' if not set.</param>
/// <param name="Visibility">The visibility of the worksheet: 'Visible', 'Hidden', or 'VeryHidden'.</param>
public record WorksheetInfo(string Name, string Color, string Visibility);
