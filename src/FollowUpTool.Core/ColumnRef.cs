namespace FollowUpTool.Core;

/// <summary>Conversions between Excel cell references ("AB12") and column numbers.</summary>
public static class ColumnRef
{
    /// <summary>"AB12" -> "AB".</summary>
    public static string LettersOf(string? cellReference) =>
        new((cellReference ?? "").TakeWhile(char.IsLetter).ToArray());

    /// <summary>1 -> "A", 27 -> "AA". Returns "" for values below 1.</summary>
    public static string ToLetters(int columnNumber)
    {
        var result = "";
        while (columnNumber > 0)
        {
            columnNumber--;
            result = (char)('A' + columnNumber % 26) + result;
            columnNumber /= 26;
        }
        return result;
    }

    /// <summary>"A" -> 1, "AA" -> 27. Returns 0 for an empty or non-letter string.</summary>
    public static int ToNumber(string? letters)
    {
        if (string.IsNullOrEmpty(letters) || !letters.All(char.IsAsciiLetter)) return 0;
        return letters.ToUpperInvariant().Aggregate(0, (acc, c) => acc * 26 + (c - 'A' + 1));
    }
}
