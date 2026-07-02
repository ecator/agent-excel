namespace AgentExcel.Models.Requests;

/// <summary>
/// Request for setting list-based data validation on a range of cells.
/// </summary>
public record SetListValidationRequest : WorksheetRequestBase
{
    /// <summary>
    /// The cell range where the validation will be applied.
    /// </summary>
    public required string Range { get; init; }

    /// <summary>
    /// The comma-separated values or sheet range reference for the list options (e.g., '"Yes,No"' or '=Sheet2!$A$1:$A$5').
    /// </summary>
    public required string Formula { get; init; }
}
