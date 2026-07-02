namespace AgentExcel.Models.Requests;

/// <summary>
/// Request for exporting a workbook as a PDF file.
/// </summary>
public record ExportPdfRequest : WorkbookRequestBase
{
    /// <summary>
    /// The absolute path where the PDF file will be saved.
    /// </summary>
    public required string OutputFile { get; init; }
}
