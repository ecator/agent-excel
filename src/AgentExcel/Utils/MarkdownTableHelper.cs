namespace AgentExcel.Utils;

/// <summary>
/// Helper for converting data matrices into Markdown tables.
/// </summary>
public static class MarkdownTableHelper
{
    /// <summary>
    /// Formats a headers array and a 2D body array into a Markdown table string.
    /// </summary>
    /// <param name="headers">The headers of the table.</param>
    /// <param name="body">The 2D body matrix of the table.</param>
    /// <returns>A formatted Markdown table string.</returns>
    /// <exception cref="ArgumentNullException">Thrown when headers or body is null.</exception>
    /// <exception cref="ArgumentException">Thrown when the number of columns in headers and body do not match.</exception>
    public static string ToMarkdownTable(string[] headers, string[,] body)
    {
        if (headers is null)
        {
            throw new ArgumentNullException(nameof(headers));
        }
        if (body is null)
        {
            throw new ArgumentNullException(nameof(body));
        }

        int headerCols = headers.Length;
        int bodyCols = body.GetLength(1);

        if (headerCols != bodyCols)
        {
            throw new ArgumentException($"Header column count ({headerCols}) does not match body column count ({bodyCols}).");
        }

        int rowStart = body.GetLowerBound(0);
        int rowEnd = body.GetUpperBound(0);
        int colStart = body.GetLowerBound(1);
        int colEnd = body.GetUpperBound(1);

        var sb = new System.Text.StringBuilder();

        // Row 1: Headers
        sb.Append("|");
        for (int c = 0; c < headerCols; c++)
        {
            string val = EscapeMarkdownCell(headers[c]);
            sb.Append(" ").Append(val).Append(" |");
        }
        sb.AppendLine();

        // Row 2: Separator
        sb.Append("|");
        for (int c = 0; c < headerCols; c++)
        {
            sb.Append(" --- |");
        }
        sb.AppendLine();

        // Row 3+: Data Rows
        for (int r = rowStart; r <= rowEnd; r++)
        {
            sb.Append("|");
            for (int c = colStart; c <= colEnd; c++)
            {
                string val = EscapeMarkdownCell(body[r, c]);
                sb.Append(" ").Append(val).Append(" |");
            }
            sb.AppendLine();
        }

        return sb.ToString().TrimEnd('\r', '\n');
    }

    private static string EscapeMarkdownCell(string? val)
    {
        if (val is null)
        {
            return "";
        }
        string str = val;
        str = str.Replace("\\", "\\\\");
        str = str.Replace("|", "\\|");
        str = str.Replace("\r\n", "<br>").Replace("\n", "<br>").Replace("\r", "<br>");
        return str;
    }
}
