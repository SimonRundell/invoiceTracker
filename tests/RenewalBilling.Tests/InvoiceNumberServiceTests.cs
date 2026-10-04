using RenewalBilling.Services;

namespace RenewalBilling.Tests;

public class InvoiceNumberServiceTests
{
    private const string Format = "INV-{yyyy}-{seq:0000}";

    [Fact]
    public void Render_PadsSequenceToConfiguredWidth()
    {
        var result = InvoiceNumberService.Render(Format, 2026, 7);

        Assert.Equal("INV-2026-0007", result);
    }

    [Fact]
    public void Render_SequenceWiderThanPadding_IsNotTruncated()
    {
        var result = InvoiceNumberService.Render(Format, 2026, 12345);

        Assert.Equal("INV-2026-12345", result);
    }

    [Fact]
    public void FindHighestSequence_NoExistingNumbers_ReturnsZero()
    {
        var highest = InvoiceNumberService.FindHighestSequence(Format, 2026, []);

        Assert.Equal(0, highest);
    }

    [Fact]
    public void FindHighestSequence_MixedPaddingAndCase_FindsTrueHighest()
    {
        // The Invoices sheet can accumulate numbers with inconsistent zero-padding and case over
        // time (manual edits, old formats). Parsing must tolerate that, since the format's
        // digit width is only a display preference.
        var existing = new[]
        {
            "INV-2026-0001",
            "inv-2026-012",      // lower case, under-padded, but still the highest this year
            "INV-2026-0005",
            "INV-2025-0099",     // a different year, must be ignored
            "garbage",
            "",
        };

        var highest = InvoiceNumberService.FindHighestSequence(Format, 2026, existing);

        Assert.Equal(12, highest);
    }

    [Fact]
    public void FindHighestSequence_OnlyOtherYears_ReturnsZero()
    {
        var existing = new[] { "INV-2025-0099", "INV-2024-0050" };

        var highest = InvoiceNumberService.FindHighestSequence(Format, 2026, existing);

        Assert.Equal(0, highest);
    }

    [Fact]
    public void BuildNext_IncrementsPastHighestExisting()
    {
        var existing = new[] { "INV-2026-0001", "INV-2026-0007", "INV-2025-0099" };

        var next = InvoiceNumberService.BuildNext(Format, 2026, existing);

        Assert.Equal("INV-2026-0008", next);
    }

    [Fact]
    public void BuildNext_NoExistingNumbersForYear_StartsAtOne()
    {
        var next = InvoiceNumberService.BuildNext(Format, 2027, ["INV-2026-0099"]);

        Assert.Equal("INV-2027-0001", next);
    }

    [Fact]
    public void FindHighestSequence_DifferentFormatWithExtraLiteralText_StillParses()
    {
        const string format = "RB/{yyyy}/{seq:000}";
        var existing = new[] { "RB/2026/001", "RB/2026/045", "RB/2025/999" };

        var highest = InvoiceNumberService.FindHighestSequence(format, 2026, existing);

        Assert.Equal(45, highest);
    }
}
