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
