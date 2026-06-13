namespace AgentExcel.Models;

/// <summary>
/// Request for grep searching a directory of Excel files.
/// </summary>
public record GrepRequest
{
    /// <summary>
    /// The directory path containing Excel files to search.
    /// </summary>
    public required string FolderPath { get; init; }

    /// <summary>
    /// The text pattern or regular expression to search for.
    /// </summary>
    public required string Pattern { get; init; }
}
