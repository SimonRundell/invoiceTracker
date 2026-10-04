using RenewalBilling.Models;
using RenewalBilling.Services;

namespace RenewalBilling.Tests;

public class TagMapTests
{
    private static InvoiceCandidate MakeCandidate()
    {
        var client = new Client
        {
            ClientId = "C1",
            ClientName = "Acme Ltd",
            ContactName = "Jo Bloggs",
            Address1 = "1 High Street",
            Address2 = "Suite 2",
            Town = "Anytown",
            Postcode = "AB1 2CD",
            Email = "jo@acme.test",
            DueDate = new DateTime(2026, 10, 20),
        };

        var items = new List<ProposedLineItem>
        {
            new(new BillingItem { ClientId = "C1", Description = "Annual licence", Quantity = 1, UnitCost = 100m }, 103m, 3m),
            new(new BillingItem { ClientId = "C1", Description = "Support plan", Quantity = 2, UnitCost = 50m }, 51.5m, 3m),
        };

        return new InvoiceCandidate
        {
            Client = client,
            Status = CandidateStatus.Due,
            Items = items,
            RpiPercentUsed = 3.456m,
            InvoiceDate = new DateTime(2026, 10, 4),
            PeriodStart = new DateTime(2026, 10, 20),
            PeriodEnd = new DateTime(2027, 10, 19),
            PayByDate = new DateTime(2026, 10, 20),
            Vat = 0m,
            InvoiceNumber = "INV-2026-0001",
        };
    }

    [Fact]
    public void BuildScalarTags_FillsAddressAndContactTags()
    {
        var tags = TagMapBuilder.BuildScalarTags(MakeCandidate());

        Assert.Equal("Acme Ltd", tags["##CLIENTNAME##"]);
        Assert.Equal("Jo Bloggs", tags["##CONTACTNAME##"]);
        Assert.Equal("1 High Street", tags["##ADDRESS1##"]);
        Assert.Equal("Suite 2", tags["##ADDRESS2##"]);
        Assert.Equal("Anytown", tags["##TOWN##"]);
        Assert.Equal("AB1 2CD", tags["##POSTCODE##"]);
        Assert.Equal("jo@acme.test", tags["##EMAIL##"]);
        Assert.Equal("INV-2026-0001", tags["##INVOICENO##"]);
    }

    [Fact]
    public void BuildScalarTags_FormatsDatesAsDdMmmYyyy()
    {
        var tags = TagMapBuilder.BuildScalarTags(MakeCandidate());

        Assert.Equal("04 Oct 2026", tags["##INVOICEDATE##"]);
        Assert.Equal("20 Oct 2026", tags["##PAYBYDATE##"]);
        Assert.Equal("20 Oct 2026", tags["##PERIODSTART##"]);
        Assert.Equal("19 Oct 2027", tags["##PERIODEND##"]);
    }

    [Fact]
    public void BuildScalarTags_FormatsRpiToOneDecimalPlace()
    {
        var tags = TagMapBuilder.BuildScalarTags(MakeCandidate());

        Assert.Equal("3.5", tags["##RPI##"]);
    }

    [Fact]
    public void BuildScalarTags_FormatsMoneyAsPoundsWithTwoDecimals()
    {
        var candidate = MakeCandidate();

        var tags = TagMapBuilder.BuildScalarTags(candidate);

        // Net = (1 * 103) + (2 * 51.5) = 206.00
        Assert.Equal("£206.00", tags["##SUBTOTAL##"]);
        Assert.Equal("£0.00", tags["##VAT##"]);
        Assert.Equal("£206.00", tags["##TOTAL##"]);
    }

    [Fact]
    public void BuildLineItemTags_ReturnsOneEntryPerLineInOrder()
    {
        var lines = TagMapBuilder.BuildLineItemTags(MakeCandidate());

        Assert.Equal(2, lines.Count);
        Assert.Equal("Annual licence", lines[0].Item);
        Assert.Equal("1", lines[0].Qty);
        Assert.Equal("£103.00", lines[0].UnitCost);
        Assert.Equal("£103.00", lines[0].LineTotal);

        Assert.Equal("Support plan", lines[1].Item);
        Assert.Equal("2", lines[1].Qty);
        Assert.Equal("£51.50", lines[1].UnitCost);
        Assert.Equal("£103.00", lines[1].LineTotal); // quantity 2 * 51.50
    }

    [Fact]
    public void FindUnknownTags_NoTags_ReturnsEmpty()
    {
        var result = TagMapBuilder.FindUnknownTags("Thank you for your business.");

        Assert.Empty(result);
    }

    [Fact]
    public void FindUnknownTags_FindsLeftoverTagsOnce()
    {
        var text = "Dear ##CLIENTNAME##, your total is ##TOTAL##. Reference ##CLIENTNAME## again.";

        var result = TagMapBuilder.FindUnknownTags(text);

        Assert.Equal(new[] { "##CLIENTNAME##", "##TOTAL##" }, result);
    }

    [Fact]
    public void FindUnknownTags_IgnoresTextThatIsNotTagShaped()
    {
        var result = TagMapBuilder.FindUnknownTags("50% off, ## not a tag ##, #SINGLE#");

        Assert.Empty(result);
    }
}
