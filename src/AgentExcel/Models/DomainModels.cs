namespace AgentExcel.Models;

/// <summary>
/// Styling options for formatting cells.
/// </summary>
/// <param name="FontName">The font family name (e.g., 'Arial').</param>
/// <param name="FontSize">The font size in points (e.g., 11).</param>
/// <param name="Bold">True to apply bold formatting.</param>
/// <param name="Italic">True to apply italic formatting.</param>
/// <param name="Color">Hex color representation for text (e.g., '#FF0000').</param>
/// <param name="BackgroundColor">Hex color representation for background fill (e.g., '#FFFF00').</param>
/// <param name="HorizontalAlignment">Horizontal alignment: 'Left', 'Center', or 'Right'.</param>
/// <param name="VerticalAlignment">Vertical alignment: 'Top', 'Center', or 'Bottom'.</param>
public record CellStyle(
    string? FontName = null,
    double? FontSize = null,
    bool? Bold = null,
    bool? Italic = null,
    string? Color = null,           // Hex like "#FF0000"
    string? BackgroundColor = null, // Hex like "#FFFF00"
    string? HorizontalAlignment = null, // "Left", "Center", "Right"
    string? VerticalAlignment = null   // "Top", "Center", "Bottom"
);

/// <summary>
/// Represents a search result from a grep operation on Excel files.
/// </summary>
/// <param name="File">The path of the file containing the match.</param>
/// <param name="Sheet">The name of the sheet.</param>
/// <param name="Type">The type of the cell (e.g. formula or value).</param>
/// <param name="Location">The cell address (e.g., 'A1').</param>
/// <param name="Content">The matching text content.</param>
public record GrepResult(string File, string Sheet, string Type, string Location, string Content);

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
public record ShapeInfo(string Name, string Type, float Left, float Top, float Width, float Height, string? Text);

/// <summary>
/// Information about a Power Query formula.
/// </summary>
/// <param name="Name">The query name.</param>
/// <param name="Formula">The M-formula of the query.</param>
/// <param name="Description">Optional description of the query.</param>
public record QueryInfo(string Name, string Formula, string Description);

/// <summary>
/// Represents a match from a find operation in a worksheet.
/// </summary>
/// <param name="Sheet">The sheet name where the match was found.</param>
/// <param name="Address">The cell address (e.g., 'A1').</param>
/// <param name="Value">The content of the matching cell.</param>
public record FindResult(string Sheet, string Address, string Value);


/// <summary>
/// Information about an Excel workbook.
/// </summary>
/// <param name="Name">The name of the workbook file.</param>
/// <param name="Path">The full file path of the workbook.</param>
/// <param name="Sheets">The list of sheet names within the workbook, or null if sheets are not loaded.</param>
public record WorkbookInfo(string Name, string Path, List<string>? Sheets);

