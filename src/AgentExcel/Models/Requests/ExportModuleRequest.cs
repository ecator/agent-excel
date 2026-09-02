namespace AgentExcel.Models.Requests;

/// <summary>
/// Request for exporting a VBA component (Module, Class, Form, or Document) from a workbook to a file.
/// </summary>
public record ExportModuleRequest : WorkbookRequestBase
{
    /// <summary>
    /// The name of the VBA component/module to export (e.g., 'Module1', 'Class1', 'UserForm1').
    /// </summary>
    public required string Module { get; init; }

    /// <summary>
    /// The absolute file path where the component file will be saved (encoded in system default ANSI).
    /// </summary>
    public required string OutputFile { get; init; }
}
