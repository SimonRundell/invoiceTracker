using RenewalBilling.Services;

namespace RenewalBilling.Tests;

public class RpiServiceTests
{
    [Fact]
    public void ParseMonthsJson_RealOnsShape_ReturnsLatestNumericEntry()
    {
        // Matches the real shape confirmed against the live ONS endpoint: a top-level "months"
        // array, each entry a string "value", oldest first.
        const string json = """
            {
              "months": [
                { "date": "1948 JUN", "value": "9.7", "year": "1948", "month": "June" },
                { "date": "2026 JUL", "value": "3.2", "year": "2026", "month": "July" },
                { "date": "2026 AUG", "value": "3.4", "year": "2026", "month": "August" }
              ]
            }
            """;

        var (percent, periodLabel) = RpiService.ParseMonthsJson(json);

        Assert.Equal(3.4m, percent);
        Assert.Equal("2026 AUG", periodLabel);
    }

    [Fact]
    public void ParseMonthsJson_TrailingPlaceholderWithoutNumericValue_SkipsBackToLatestNumeric()
    {
        const string json = """
            {
              "months": [
                { "date": "2026 JUL", "value": "3.2" },
                { "date": "2026 AUG", "value": "3.4" },
                { "date": "2026 SEP", "value": "" }
              ]
            }
            """;

        var (percent, periodLabel) = RpiService.ParseMonthsJson(json);

        Assert.Equal(3.4m, percent);
        Assert.Equal("2026 AUG", periodLabel);
    }

    [Fact]
    public void ParseMonthsJson_NegativeRpi_ParsesCorrectly()
    {
        const string json = """{"months":[{"date":"2026 AUG","value":"-0.3"}]}""";

        var (percent, periodLabel) = RpiService.ParseMonthsJson(json);

        Assert.Equal(-0.3m, percent);
        Assert.Equal("2026 AUG", periodLabel);
    }

    [Fact]
    public void ParseMonthsJson_NoMonthsArray_ThrowsRpiParseException()
    {
        const string json = """{"years":[{"date":"2026","value":"3.4"}]}""";

        var ex = Assert.Throws<RpiParseException>(() => RpiService.ParseMonthsJson(json));
        Assert.Contains("months", ex.Message);
    }

    [Fact]
    public void ParseMonthsJson_EmptyMonthsArray_ThrowsRpiParseException()
    {
        const string json = """{"months":[]}""";

        Assert.Throws<RpiParseException>(() => RpiService.ParseMonthsJson(json));
    }

    [Fact]
    public void ParseMonthsJson_EveryEntryNonNumeric_ThrowsRpiParseException()
    {
        const string json = """{"months":[{"date":"2026 JUL","value":""},{"date":"2026 AUG","value":"n/a"}]}""";

        Assert.Throws<RpiParseException>(() => RpiService.ParseMonthsJson(json));
    }

    [Fact]
    public void ParseMonthsJson_MalformedJson_ThrowsRpiParseException()
    {
        const string json = "{ this is not json";

        Assert.Throws<RpiParseException>(() => RpiService.ParseMonthsJson(json));
    }
}
