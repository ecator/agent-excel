namespace AgentExcel.Models;

/// <summary>
/// Information about a shape in a worksheet.
/// </summary>
/// <param name="Name">The unique name of the shape.</param>
/// <param name="Type">The type or geometry of the shape.</param>
/// <param name="Left">Horizontal position in points.</param>
/// <param name="Top">Vertical position in points.</param>
/// <param name="Width">Width of the shape in points.</param>
/// <param name="Height">Height of the shape in points.</param>
/// <param name="Text">The text inside the shape, if any.</param>
/// <param name="Connection">Flowchart connection details if the shape is a connector.</param>
public record ShapeInfo(
    string Name,
    string Type,
    float Left,
    float Top,
    float Width,
    float Height,
    string? Text,
    ShapeConnectionInfo? Connection = null
);
