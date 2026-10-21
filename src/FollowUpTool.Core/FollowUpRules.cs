namespace FollowUpTool.Core;

/// <summary>
/// Decides which tracker rows need a follow-up email. Kept separate from the
/// UI so the rules are unit-tested.
/// </summary>
public static class FollowUpRules
{
    /// <summary>Days after the application date when the first follow-up is due, if no date is set.</summary>
    public const int DefaultFollowUpDelayDays = 7;

    private static readonly HashSet<string> ClosedStatuses = new(StringComparer.OrdinalIgnoreCase)
    {
        "Rejected", "Withdrawn", "Offer Received",
    };

    // Values people type while still waiting to hear back ("No yet" is a real typo from the source tracker).
    private static readonly HashSet<string> AwaitingResponseStatuses = new(StringComparer.OrdinalIgnoreCase)
    {
        "", "Not yet", "No yet", "Applied", "No Response",
    };

    /// <summary>The application is finished; never follow up.</summary>
    public static bool IsClosed(string? status) => ClosedStatuses.Contains((status ?? "").Trim());

    /// <summary>No reply has been recorded yet.</summary>
    public static bool IsAwaitingResponse(string? status) => AwaitingResponseStatuses.Contains((status ?? "").Trim());

    /// <summary>"Yes" in the "Followup done" column means the email was already sent.</summary>
    public static bool IsMarkedDone(string? followupDone) =>
        string.Equals((followupDone ?? "").Trim(), "Yes", StringComparison.OrdinalIgnoreCase);

    /// <summary>Explicit follow-up date if present, otherwise applied date + 7 days.</summary>
    public static DateTime? DueDate(DateTime? appliedDate, DateTime? followUpDate) =>
        followUpDate?.Date ?? appliedDate?.Date.AddDays(DefaultFollowUpDelayDays);

    /// <summary>
    /// Due today or overdue. (The original check was "== today", which meant a
    /// follow-up missed on its day was never shown again.)
    /// </summary>
    public static bool IsDue(DateTime? dueDate, DateTime today) =>
        dueDate.HasValue && dueDate.Value.Date <= today.Date;

    /// <summary>Full eligibility check for one tracker row.</summary>
    public static bool NeedsFollowUp(string? status, string? followupDone, DateTime? appliedDate,
                                     DateTime? followUpDate, DateTime today) =>
        !IsMarkedDone(followupDone)
        && !IsClosed(status)
        && IsAwaitingResponse(status)
        && IsDue(DueDate(appliedDate, followUpDate), today);

    /// <summary>"first", "second", "third", "4th", ...</summary>
    public static string Ordinal(int followUpNumber) => followUpNumber switch
    {
        1 => "first",
        2 => "second",
        3 => "third",
        _ => $"{followUpNumber}th",
    };
}
