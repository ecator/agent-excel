namespace AgentExcel.Models.Requests;

/// <summary>
/// Request for deleting a VBA component (Module, Class, Form) from a workbook.
/// </summary>
public record DeleteModuleRequest : WorkbookRequestBase
{
    /// <summary>
    /// The name of the VBA component/module to delete (e.g., 'Module1', 'Class1', 'UserForm1').
    /// </summary>
    public required string Module { get; init; }
}
