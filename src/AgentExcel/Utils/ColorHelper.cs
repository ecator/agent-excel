using System;

namespace AgentExcel.Utils;

/// <summary>
/// Helper utilities for color conversions.
/// </summary>
public static class ColorHelper
{
    /// <summary>
    /// Converts a hex color string (e.g. '#FF0000' or 'FF0000') to OLE Color.
    /// </summary>
    public static int HexToOleColor(string hex)
    {
        try
        {
            hex = hex.TrimStart('#');
            if (hex.Length == 6)
            {
                int r = int.Parse(hex.Substring(0, 2), System.Globalization.NumberStyles.HexNumber);
                int g = int.Parse(hex.Substring(2, 2), System.Globalization.NumberStyles.HexNumber);
                int b = int.Parse(hex.Substring(4, 2), System.Globalization.NumberStyles.HexNumber);
                return (b << 16) | (g << 8) | r;
            }
        }
        catch { }
        return 0;
    }

    /// <summary>
    /// Converts an OLE Color value to a Hex color string (e.g. '#FF0000').
    /// </summary>
    public static string OleColorToHex(long oleColor)
    {
        int r = (int)(oleColor & 0xFF);
        int g = (int)((oleColor >> 8) & 0xFF);
        int b = (int)((oleColor >> 16) & 0xFF);
        return $"#{r:X2}{g:X2}{b:X2}";
    }
}
