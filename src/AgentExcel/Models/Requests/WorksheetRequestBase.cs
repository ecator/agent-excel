namespace AgentExcel.Models.Requests;

/// <summary>
/// Base request model for worksheet operations.
/// </summary>
public abstract record WorksheetRequestBase : WorkbookRequestBase
{
    /// <summary>
    /// The target worksheet name (e.g., 'Sheet1').
    /// </summary>
    public required string Sheet { get; init; }
}
