namespace AgentExcel.Models.Requests;

/// <summary>
/// Request model for setting a worksheet's visibility state.
/// </summary>
public record SetSheetVisibilityRequest : WorksheetRequestBase
{
    /// <summary>
    /// The target visibility state to set ('Visible', 'Hidden', or 'VeryHidden').
    /// </summary>
    public required string Visibility { get; init; }
}
