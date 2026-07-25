namespace AgentExcel.Models;

/// <summary>
/// Information about a comment in a worksheet.
/// </summary>
/// <param name="Address">The cell address (e.g., 'A1').</param>
/// <param name="Visible">Whether the comment is visible.</param>
/// <param name="Author">The author of the comment.</param>
/// <param name="Text">The text content of the comment.</param>
public record CommentInfo(
    string Address,
    bool Visible,
    string Author,
    string Text
);
