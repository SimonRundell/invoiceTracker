using System.Globalization;
using System.Text.RegularExpressions;
using RenewalBilling.Models;

namespace RenewalBilling.Services;

/// <summary>
/// One line item's tag values, ready to fill a row in the invoice template's item table.
/// </summary>
/// <param name="Item">The <c>##ITEM##</c> value: the line's description.</param>
/// <param name="Qty">The <c>##QTY##</c> value.</param>
/// <param name="UnitCost">The <c>##UNITCOST##</c> value, formatted as money.</param>
/// <param name="LineTotal">The <c>##LINETOTAL##</c> value, formatted as money.</param>
public sealed record LineItemTags(string Item, string Qty, string UnitCost, string LineTotal);

/// <summary>
/// Builds the tag values described in section 10 of the specification from an
/// <see cref="InvoiceCandidate"/>. Pure logic: holds no state and touches no Word document, so
/// it can be unit tested directly. The actual find-and-replace against the Word template is
/// COM-dependent and lives in the Word invoice builder.
/// </summary>
public static class TagMapBuilder
{
    private static readonly CultureInfo EnGb = CultureInfo.GetCultureInfo("en-GB");
    private static readonly Regex UnknownTagRegex = new(@"##[A-Z0-9]+##", RegexOptions.Compiled);

    /// <summary>
    /// Builds the scalar (non-repeating) tag values for one invoice candidate.
    /// </summary>
    /// <param name="candidate">The candidate to build tags for.</param>
    /// <returns>A map of tag name to the text that should replace it.</returns>
    public static IReadOnlyDictionary<string, string> BuildScalarTags(InvoiceCandidate candidate)
    {
        var client = candidate.Client;

        return new Dictionary<string, string>
        {
            ["##CLIENTNAME##"] = client.ClientName,
            ["##CONTACTNAME##"] = client.ContactName,
            ["##ADDRESS1##"] = client.Address1,
            ["##ADDRESS2##"] = client.Address2,
            ["##TOWN##"] = client.Town,
            ["##POSTCODE##"] = client.Postcode,
            ["##EMAIL##"] = client.Email,
            ["##INVOICENO##"] = candidate.InvoiceNumber ?? string.Empty,
            ["##INVOICEDATE##"] = FormatDate(candidate.InvoiceDate),
            ["##PAYBYDATE##"] = FormatDate(candidate.PayByDate),
            ["##PERIODSTART##"] = FormatDate(candidate.PeriodStart),
            ["##PERIODEND##"] = FormatDate(candidate.PeriodEnd),
            ["##RPI##"] = candidate.RpiPercentUsed.ToString("0.0", EnGb),
            ["##SUBTOTAL##"] = FormatMoney(candidate.Net),
            ["##VAT##"] = FormatMoney(candidate.Vat),
            ["##TOTAL##"] = FormatMoney(candidate.Total),
        };
    }

    /// <summary>
    /// Builds the per-line tag values for the invoice candidate's item table, one entry per
    /// proposed line item, in order.
    /// </summary>
    /// <param name="candidate">The candidate to build line item tags for.</param>
    /// <returns>One <see cref="LineItemTags"/> per proposed line item.</returns>
    public static IReadOnlyList<LineItemTags> BuildLineItemTags(InvoiceCandidate candidate)
    {
        return candidate.Items
            .Select(line => new LineItemTags(
                line.Item.Description,
                line.Item.Quantity.ToString("0.##", EnGb),
                FormatMoney(line.ProposedUnitCost),
                FormatMoney(line.LineTotal)))
            .ToList();
    }

    /// <summary>
    /// Scans text for any <c>##TAG##</c> pattern that remains after replacement, so a leftover
    /// or misspelled tag can be logged as a warning instead of silently appearing in the
    /// finished document.
    /// </summary>
    /// <param name="text">The text to scan, typically one story range of the finished document.</param>
    /// <returns>The distinct leftover tags found, in the order they first appear.</returns>
    public static IReadOnlyList<string> FindUnknownTags(string text)
    {
        return UnknownTagRegex.Matches(text)
            .Select(match => match.Value)
            .Distinct()
            .ToList();
    }

    private static string FormatDate(DateTime value) => value.ToString("dd MMM yyyy", EnGb);

    private static string FormatMoney(decimal value) => value.ToString("C2", EnGb);
}
