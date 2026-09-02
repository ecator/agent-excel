namespace AgentExcel.Models.Requests;

/// <summary>
/// Request for importing a VBA component file (Module, Class, Form) into a workbook.
/// </summary>
public record ImportModuleRequest : WorkbookRequestBase
{
    /// <summary>
    /// The absolute file path of the VBA component file to import (e.g., 'C:\path\Module1.bas').
    /// </summary>
    public required string Path { get; init; }
}
