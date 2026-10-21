using System.Text.Json;
using System.Text.Json.Nodes;

namespace FollowUpTool.Core;

/// <summary>Inputs for one follow-up draft.</summary>
public sealed record DraftRequest(
    string Company,
    string Role,
    string RecruiterName,
    DateTime? AppliedDate,
    int PreviousFollowUps,
    string Notes,
    string SenderName);

/// <summary>
/// Builds the Claude Messages API request for a follow-up draft and reads the
/// reply. No network code lives here, so both halves are unit-tested.
/// </summary>
public static class DraftPrompt
{
    /// <summary>Model used when the CLAUDE_MODEL environment variable is not set.</summary>
    public const string DefaultModel = "claude-sonnet-4-6";
    public const string ApiVersion = "2023-06-01";
    public const int MaxTokens = 400;

    // Notes come from a free-text spreadsheet cell; cap them so one huge cell
    // cannot blow up the request size or dominate the prompt.
    public const int MaxNotesLength = 500;

    public static string BuildPrompt(DraftRequest r)
    {
        var ordinal = FollowUpRules.Ordinal(r.PreviousFollowUps + 1);
        var applied = r.AppliedDate?.ToString("d MMMM yyyy", System.Globalization.CultureInfo.InvariantCulture) ?? "recently";
        var recruiter = string.IsNullOrWhiteSpace(r.RecruiterName) ? "Hiring Team" : r.RecruiterName.Trim();
        var notes = (r.Notes ?? "").Trim();
        if (notes.Length > MaxNotesLength) notes = notes[..MaxNotesLength];
        var sender = string.IsNullOrWhiteSpace(r.SenderName) ? "the applicant" : r.SenderName.Trim();
        var signOff = string.IsNullOrWhiteSpace(r.SenderName) ? "Regards," : $"Regards,\n{sender}";

        return $"""
            Write a {ordinal} follow-up email for a job application. Follow the tone profile exactly.

            Application Details:
            - Company: {r.Company}
            - Role: {r.Role}
            - Recruiter/Contact: {recruiter}
            - Applied Date: {applied}
            - Sender Name: {sender}
            - Notes: {(notes.Length == 0 ? "None" : notes)}

            Tone Profile (follow strictly):
            - Greeting: "Hi [First Name]," for known contacts, "Hi," for unknown. Never "Dear".
            - Opening: "Thank you for reaching out!" only when they initiated. For follow-ups, skip opener and get to the point directly.
            - Body: 2-3 short sentences. One idea per sentence. Direct and warm. No filler phrases.
            - Sign-off: "{signOff.Replace("\n", "\\n")}"
            - No bullet points, no emojis, no placeholders in square brackets.

            Output only the email body (no subject line). Keep it under 100 words.
            """;
    }

    /// <summary>JSON body for POST /v1/messages.</summary>
    public static string BuildRequestJson(DraftRequest r, string model)
    {
        var payload = new
        {
            model = string.IsNullOrWhiteSpace(model) ? DefaultModel : model,
            max_tokens = MaxTokens,
            messages = new[] { new { role = "user", content = BuildPrompt(r) } },
        };
        return JsonSerializer.Serialize(payload);
    }

    /// <summary>
    /// Extracts the text of the first "text" content block from a successful
    /// Messages API response.
    /// </summary>
    /// <exception cref="FormatException">The response has no text block.</exception>
    public static string ParseResponse(string responseJson)
    {
        JsonNode? doc;
        try { doc = JsonNode.Parse(responseJson); }
        catch (JsonException ex) { throw new FormatException("Claude API returned invalid JSON.", ex); }

        if (doc?["content"] is JsonArray blocks)
        {
            foreach (var block in blocks)
            {
                if (block?["type"]?.GetValue<string>() == "text")
                {
                    var text = block["text"]?.GetValue<string>()?.Trim();
                    if (!string.IsNullOrEmpty(text)) return text;
                }
            }
        }
        throw new FormatException("Claude API response contained no text.");
    }

    /// <summary>
    /// Short, user-facing description of an API error response. Uses the
    /// API's own error message when present instead of dumping the raw body.
    /// </summary>
    public static string DescribeError(int statusCode, string responseBody)
    {
        string? message = null;
        try { message = JsonNode.Parse(responseBody)?["error"]?["message"]?.GetValue<string>(); }
        catch (JsonException) { /* not JSON - fall through */ }

        var hint = statusCode switch
        {
            401 => " Check that the Anthropic API key is correct.",
            429 => " Rate limit reached - wait a moment and try again.",
            >= 500 => " The API is having problems - try again shortly.",
            _ => "",
        };
        return $"Claude API error {statusCode}: {message ?? "unexpected response"}.{hint}";
    }
}
