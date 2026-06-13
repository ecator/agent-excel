namespace AgentExcel.Models;

/// <summary>
/// Request for running a VBA macro.
/// </summary>
public record RunMacroRequest : WorkbookRequestBase
{
    /// <summary>
    /// The name of the macro to execute.
    /// </summary>
    public required string MacroName { get; init; }

    /// <summary>
    /// Optional array of arguments to pass to the macro.
    /// </summary>
    public object[]? Args { get; init; }
}
