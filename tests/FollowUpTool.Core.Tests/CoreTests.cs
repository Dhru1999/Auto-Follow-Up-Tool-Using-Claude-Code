using System.Globalization;
using System.Text.Json.Nodes;
using FollowUpTool.Core;
using Xunit;

namespace FollowUpTool.Core.Tests;

public class HeaderMapperTests
{
    [Theory]
    [InlineData("Company Name", Fields.Company)]
    [InlineData("  company_name ", Fields.Company)]
    [InlineData("Date of applied", Fields.AppliedDate)]
    [InlineData("Repsonse received", Fields.Status)]
    [InlineData("Followup done", Fields.FollowupDone)]
    [InlineData("Company Email", Fields.CompanyEmail)]
    [InlineData("Follow up\ndate", Fields.FollowUpDate)]
    [InlineData("Role ", Fields.Role)]
    public void Maps_known_aliases(string header, string expected) =>
        Assert.Equal(expected, HeaderMapper.Map(header));

    [Fact]
    public void Unknown_header_is_returned_trimmed() =>
        Assert.Equal("Salary Range", HeaderMapper.Map("  Salary Range "));

    [Fact]
    public void Null_header_maps_to_empty() =>
        Assert.Equal("", HeaderMapper.Map(null));
}

public class CellDateParserTests
{
    [Fact]
    public void Parses_excel_serial_number() =>
        Assert.Equal(new DateTime(2026, 5, 15), CellDateParser.Parse("46157"));

    [Fact]
    public void Parses_iso_date() =>
        Assert.Equal(new DateTime(2026, 5, 15), CellDateParser.Parse("2026-05-15"));

    [Theory]
    [InlineData("15.05.2026")]
    [InlineData("15.5.2026")]
    [InlineData("15.05.26")]
    public void Parses_german_dotted_dates_regardless_of_culture(string text)
    {
        // en-US would read "05.06" style text month-first; the explicit formats must win.
        var result = CellDateParser.Parse(text, CultureInfo.GetCultureInfo("en-US"));
        Assert.Equal(new DateTime(2026, 5, 15), result);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("not a date")]
    [InlineData("0")]
    [InlineData("-5")]
    public void Returns_null_for_unparseable_input(string? text) =>
        Assert.Null(CellDateParser.Parse(text, CultureInfo.InvariantCulture));
}

public class FollowUpRulesTests
{
    private static readonly DateTime Today = new(2026, 10, 21);

    [Fact]
    public void Due_date_defaults_to_applied_plus_seven_days() =>
        Assert.Equal(new DateTime(2026, 10, 21), FollowUpRules.DueDate(new DateTime(2026, 10, 14), null));

    [Fact]
    public void Explicit_follow_up_date_wins_over_default() =>
        Assert.Equal(new DateTime(2026, 11, 1),
            FollowUpRules.DueDate(new DateTime(2026, 10, 14), new DateTime(2026, 11, 1)));

    [Fact]
    public void Overdue_follow_up_is_still_due() =>
        Assert.True(FollowUpRules.IsDue(Today.AddDays(-3), Today));

    [Fact]
    public void Future_follow_up_is_not_due() =>
        Assert.False(FollowUpRules.IsDue(Today.AddDays(1), Today));

    [Fact]
    public void Missing_dates_are_never_due() =>
        Assert.False(FollowUpRules.NeedsFollowUp("Applied", "", null, null, Today));

    [Theory]
    [InlineData("", "", true)]
    [InlineData("Applied", "", true)]
    [InlineData("No yet", "", true)]
    [InlineData("Applied", "Yes", false)]       // already sent
    [InlineData("Applied", " yes ", false)]
    [InlineData("Rejected", "", false)]
    [InlineData("Offer Received", "", false)]
    [InlineData("Interview Scheduled", "", false)] // got a reply - not "awaiting"
    public void Eligibility_respects_status_and_done_flag(string status, string done, bool expected) =>
        Assert.Equal(expected,
            FollowUpRules.NeedsFollowUp(status, done, Today.AddDays(-10), null, Today));

    [Theory]
    [InlineData(1, "first")]
    [InlineData(3, "third")]
    [InlineData(4, "4th")]
    public void Ordinal_words(int n, string expected) =>
        Assert.Equal(expected, FollowUpRules.Ordinal(n));
}

public class ColumnRefTests
{
    [Theory]
    [InlineData("A1", "A")]
    [InlineData("AB12", "AB")]
    [InlineData("", "")]
    [InlineData(null, "")]
    public void Extracts_letters(string? cellRef, string expected) =>
        Assert.Equal(expected, ColumnRef.LettersOf(cellRef));

    [Theory]
    [InlineData(1, "A")]
    [InlineData(26, "Z")]
    [InlineData(27, "AA")]
    [InlineData(702, "ZZ")]
    [InlineData(703, "AAA")]
    public void Number_and_letters_round_trip(int number, string letters)
    {
        Assert.Equal(letters, ColumnRef.ToLetters(number));
        Assert.Equal(number, ColumnRef.ToNumber(letters));
        Assert.Equal(number, ColumnRef.ToNumber(letters.ToLowerInvariant()));
    }

    [Fact]
    public void Invalid_letters_map_to_zero() => Assert.Equal(0, ColumnRef.ToNumber("A1"));
}

public class DraftPromptTests
{
    private static DraftRequest Sample(string notes = "Met at career fair", string sender = "Alex Doe") =>
        new("Example GmbH", "Backend Developer", "Jane Smith", new DateTime(2026, 10, 1), 0, notes, sender);

    [Fact]
    public void Prompt_contains_application_details()
    {
        var prompt = DraftPrompt.BuildPrompt(Sample());
        Assert.Contains("Example GmbH", prompt);
        Assert.Contains("Backend Developer", prompt);
        Assert.Contains("Jane Smith", prompt);
        Assert.Contains("1 October 2026", prompt);
        Assert.Contains("first follow-up", prompt);
        Assert.Contains("Met at career fair", prompt);
    }

    [Fact]
    public void Very_long_notes_are_truncated()
    {
        var prompt = DraftPrompt.BuildPrompt(Sample(notes: new string('x', 5000)));
        Assert.DoesNotContain(new string('x', DraftPrompt.MaxNotesLength + 1), prompt);
        Assert.Contains(new string('x', DraftPrompt.MaxNotesLength), prompt);
    }

    [Fact]
    public void Missing_sender_name_does_not_inject_someone_elses_name()
    {
        var prompt = DraftPrompt.BuildPrompt(Sample(sender: ""));
        Assert.Contains("Sender Name: the applicant", prompt);
    }

    [Fact]
    public void Request_json_uses_given_model_and_one_user_message()
    {
        var json = JsonNode.Parse(DraftPrompt.BuildRequestJson(Sample(), "claude-test-model"))!;
        Assert.Equal("claude-test-model", json["model"]!.GetValue<string>());
        Assert.Equal(DraftPrompt.MaxTokens, json["max_tokens"]!.GetValue<int>());
        Assert.Single(json["messages"]!.AsArray());
        Assert.Equal("user", json["messages"]![0]!["role"]!.GetValue<string>());
    }

    [Fact]
    public void Request_json_falls_back_to_default_model() =>
        Assert.Equal(DraftPrompt.DefaultModel,
            JsonNode.Parse(DraftPrompt.BuildRequestJson(Sample(), ""))!["model"]!.GetValue<string>());

    [Fact]
    public void Parses_first_text_block()
    {
        const string response = """
            {"content":[{"type":"thinking","thinking":"..."},{"type":"text","text":"  Hi Jane,\n\nJust following up.  "}]}
            """;
        Assert.Equal("Hi Jane,\n\nJust following up.", DraftPrompt.ParseResponse(response));
    }

    [Theory]
    [InlineData("{\"content\":[]}")]
    [InlineData("{\"error\":{\"message\":\"boom\"}}")]
    [InlineData("not json")]
    public void Unusable_response_throws_format_exception(string response) =>
        Assert.Throws<FormatException>(() => DraftPrompt.ParseResponse(response));

    [Fact]
    public void Error_description_uses_api_message_and_hint()
    {
        var text = DraftPrompt.DescribeError(401,
            """{"type":"error","error":{"type":"authentication_error","message":"invalid x-api-key"}}""");
        Assert.Contains("401", text);
        Assert.Contains("invalid x-api-key", text);
        Assert.Contains("API key", text);
    }

    [Fact]
    public void Error_description_survives_non_json_body() =>
        Assert.Contains("unexpected response", DraftPrompt.DescribeError(502, "<html>Bad gateway</html>"));
}

public class SettingsSecretsTests
{
    /// <summary>Reversible stand-in for DPAPI (XOR) so the round trip runs on any OS.</summary>
    private sealed class FakeProtector : ISecretProtector
    {
        public byte[] Protect(byte[] plain) => plain.Select(b => (byte)(b ^ 0x5A)).ToArray();
        public byte[] Unprotect(byte[] cipher) => Protect(cipher);
    }

    private sealed class FailingProtector : ISecretProtector
    {
        public byte[] Protect(byte[] plain) => plain;
        public byte[] Unprotect(byte[] cipher) =>
            throw new System.Security.Cryptography.CryptographicException("wrong user");
    }

    [Fact]
    public void Round_trips_a_secret()
    {
        var p = new FakeProtector();
        var stored = SettingsSecrets.Encrypt(p, "s3cret value");
        Assert.NotEqual("s3cret value", stored);
        Assert.Equal("s3cret value", SettingsSecrets.Decrypt(p, stored));
    }

    [Fact]
    public void Empty_secret_is_stored_as_empty() =>
        Assert.Equal("", SettingsSecrets.Encrypt(new FakeProtector(), ""));

    [Theory]
    [InlineData("%%% not base64 %%%")]
    [InlineData("")]
    [InlineData(null)]
    public void Unreadable_values_decrypt_to_empty(string? stored) =>
        Assert.Equal("", SettingsSecrets.Decrypt(new FakeProtector(), stored));

    [Fact]
    public void Decryption_failure_returns_empty_instead_of_throwing() =>
        Assert.Equal("", SettingsSecrets.Decrypt(new FailingProtector(), Convert.ToBase64String(new byte[] { 1, 2, 3 })));
}
