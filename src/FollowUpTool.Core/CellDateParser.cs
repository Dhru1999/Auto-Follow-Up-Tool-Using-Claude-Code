using System.Globalization;

namespace FollowUpTool.Core;

/// <summary>
/// Turns the text of an Excel cell into a date. Trackers kept by hand mix
/// real Excel dates (stored as serial numbers), ISO dates and German-style
/// "dd.MM.yyyy" text, so all three are handled explicitly.
/// </summary>
public static class CellDateParser
{
    private static readonly string[] ExplicitFormats =
    {
        "yyyy-MM-dd", "yyyy-MM-ddTHH:mm:ss", "yyyy-MM-dd HH:mm:ss",
        "dd.MM.yyyy", "d.M.yyyy", "dd.MM.yy", "d.M.yy",
    };

    // Excel serial 1 = 1900-01-01; 2958465 = 9999-12-31.
    private const double MinSerial = 1;
    private const double MaxSerial = 2958465;

    /// <param name="text">Raw cell text.</param>
    /// <param name="fallbackCulture">
    /// Culture used for free-form text that matches none of the explicit
    /// formats (e.g. "15 May 2026"). Defaults to the current UI culture, which
    /// mirrors how the user typed the value in Excel.
    /// </param>
    public static DateTime? Parse(string? text, CultureInfo? fallbackCulture = null)
    {
        if (string.IsNullOrWhiteSpace(text)) return null;
        text = text.Trim();

        // 1. Excel serial number - the most common case for real date cells.
        if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var serial))
        {
            if (serial >= MinSerial && serial <= MaxSerial)
            {
                try { return DateTime.FromOADate(serial).Date; }
                catch (ArgumentException) { return null; }
            }
            return null;
        }

        // 2. Unambiguous formats, culture-independent.
        if (DateTime.TryParseExact(text, ExplicitFormats, CultureInfo.InvariantCulture,
                DateTimeStyles.AllowWhiteSpaces, out var exact))
            return exact.Date;

        // 3. Anything else, interpreted the way the user's machine would.
        if (DateTime.TryParse(text, fallbackCulture ?? CultureInfo.CurrentCulture,
                DateTimeStyles.AllowWhiteSpaces, out var loose))
            return loose.Date;

        return null;
    }
}
