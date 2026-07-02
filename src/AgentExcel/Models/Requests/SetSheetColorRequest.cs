namespace AgentExcel.Models.Requests;

/// <summary>
/// Request model for setting a worksheet's tab color.
/// </summary>
public record SetSheetColorRequest : WorksheetRequestBase
{
    /// <summary>
    /// The target color to set (Hex e.g. '#FF0000', or 'None' to clear/reset).
    /// </summary>
    public required string Color { get; init; }
}
