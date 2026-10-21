namespace FollowUpTool.Core;

/// <summary>
/// Maps the column headers people actually use in their trackers
/// ("Date of applied", "company_email", "Repsonse received", ...) to the
/// canonical field names the app works with.
/// </summary>
public static class HeaderMapper
{
    private static readonly Dictionary<string, string> Aliases = new(StringComparer.OrdinalIgnoreCase)
    {
        ["company"]             = Fields.Company,
        ["company name"]        = Fields.Company,
        ["companyname"]         = Fields.Company,
        ["company_name"]        = Fields.Company,
        ["role"]                = Fields.Role,
        ["job title"]           = Fields.Role,
        ["date of applied"]     = Fields.AppliedDate,
        ["date applied"]        = Fields.AppliedDate,
        ["applied date"]        = Fields.AppliedDate,
        ["applied_date"]        = Fields.AppliedDate,
        ["follow up date"]      = Fields.FollowUpDate,
        ["follow_up_date"]      = Fields.FollowUpDate,
        ["followup_date"]       = Fields.FollowUpDate,
        ["followup date"]       = Fields.FollowUpDate,
        ["company email"]       = Fields.CompanyEmail,
        ["company_email"]       = Fields.CompanyEmail,
        ["company emails"]      = Fields.CompanyEmail,
        ["company_emails"]      = Fields.CompanyEmail,
        ["recruiter email"]     = Fields.RecruiterEmail,
        ["recruiter_email"]     = Fields.RecruiterEmail,
        ["hr email"]            = Fields.HrEmail,
        ["hr_email"]            = Fields.HrEmail,
        ["recruiter name"]      = Fields.RecruiterName,
        ["recruiter_name"]      = Fields.RecruiterName,
        ["status"]              = Fields.Status,
        ["repsonse received"]   = Fields.Status, // common typo kept on purpose
        ["response received"]   = Fields.Status,
        ["response"]            = Fields.Status,
        ["feedback"]            = Fields.Status,
        ["notes"]               = Fields.Notes,
        ["additional comments"] = Fields.Notes,
        ["follow up count"]     = Fields.FollowUpCount,
        ["follow_up_count"]     = Fields.FollowUpCount,
        ["followup done"]       = Fields.FollowupDone,
        ["followup_done"]       = Fields.FollowupDone,
        ["follow up done"]      = Fields.FollowupDone,
        ["follow_up_done"]      = Fields.FollowupDone,
    };

    /// <summary>Returns the canonical field name, or the trimmed header if it is unknown.</summary>
    public static string Map(string? rawHeader)
    {
        if (rawHeader is null) return "";
        var key = Normalize(rawHeader);
        return Aliases.TryGetValue(key, out var mapped) ? mapped : rawHeader.Trim();
    }

    /// <summary>
    /// Collapses whitespace variants Excel users produce (non-breaking spaces,
    /// line breaks inside a header cell, double spaces) and trims.
    /// </summary>
    internal static string Normalize(string header)
    {
        var cleaned = header.Replace(' ', ' ').Replace('\r', ' ').Replace('\n', ' ');
        return string.Join(' ', cleaned.Split(' ', StringSplitOptions.RemoveEmptyEntries)).ToLowerInvariant();
    }
}

/// <summary>Canonical field names produced by <see cref="HeaderMapper"/>.</summary>
public static class Fields
{
    public const string Company = "Company";
    public const string Role = "Role";
    public const string AppliedDate = "AppliedDate";
    public const string FollowUpDate = "FollowUpDate";
    public const string CompanyEmail = "CompanyEmail";
    public const string RecruiterEmail = "RecruiterEmail";
    public const string HrEmail = "HrEmail";
    public const string RecruiterName = "RecruiterName";
    public const string Status = "Status";
    public const string Notes = "Notes";
    public const string FollowUpCount = "FollowUpCount";
    public const string FollowupDone = "FollowupDone";
}
