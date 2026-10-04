using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace RenewalBilling.Services;

/// <summary>
/// Builds and parses invoice numbers from the configured <c>invoiceNumberFormat</c>, for example
/// <c>INV-{yyyy}-{seq:0000}</c>. Pure logic: holds no state and touches no file, network or COM
/// resource, so it can be unit tested directly.
/// </summary>
public static class InvoiceNumberService
{
    private static readonly Regex PlaceholderRegex = new(@"\{(yyyy|seq:(0+))\}", RegexOptions.Compiled);

    /// <summary>
    /// Renders an invoice number for a given year and sequence number.
    /// </summary>
    /// <param name="format">The configured format, for example <c>INV-{yyyy}-{seq:0000}</c>.</param>
    /// <param name="year">The invoice year.</param>
    /// <param name="sequence">The sequence number within that year.</param>
    /// <returns>The rendered invoice number.</returns>
    public static string Render(string format, int year, int sequence)
    {
        return PlaceholderRegex.Replace(format, match =>
        {
            if (match.Groups[1].Value == "yyyy")
            {
                return year.ToString("0000", CultureInfo.InvariantCulture);
            }

            var paddingWidth = match.Groups[2].Value.Length;
            return sequence.ToString(new string('0', paddingWidth), CultureInfo.InvariantCulture);
        });
    }

    /// <summary>
    /// Scans a set of existing invoice numbers for the given year and returns the highest
    /// sequence number found, or zero if none match. Matching tolerates mixed zero-padding and
    /// case in the existing data, since a format's digit width is a display preference, not a
    /// parsing constraint, and worksheet data can accumulate inconsistencies over time.
    /// </summary>
    /// <param name="format">The configured format, for example <c>INV-{yyyy}-{seq:0000}</c>.</param>
    /// <param name="year">The year to find the highest sequence for.</param>
    /// <param name="existingInvoiceNumbers">Invoice numbers already present in the Invoices worksheet.</param>
    /// <returns>The highest sequence number found for that year, or zero.</returns>
    public static int FindHighestSequence(string format, int year, IEnumerable<string> existingInvoiceNumbers)
    {
        var regex = BuildParseRegex(format, year);
        var highest = 0;

        foreach (var number in existingInvoiceNumbers)
        {
            if (string.IsNullOrWhiteSpace(number))
            {
                continue;
            }

            var match = regex.Match(number.Trim());
            if (match.Success && int.TryParse(match.Groups[1].Value, out var sequence) && sequence > highest)
            {
                highest = sequence;
            }
        }

        return highest;
    }

    /// <summary>
    /// Builds the next invoice number for a year, one past the highest existing sequence found.
    /// </summary>
    /// <param name="format">The configured format, for example <c>INV-{yyyy}-{seq:0000}</c>.</param>
    /// <param name="year">The invoice year.</param>
    /// <param name="existingInvoiceNumbers">Invoice numbers already present in the Invoices worksheet.</param>
    /// <returns>The next invoice number to use.</returns>
    public static string BuildNext(string format, int year, IEnumerable<string> existingInvoiceNumbers)
    {
        var highest = FindHighestSequence(format, year, existingInvoiceNumbers);
        return Render(format, year, highest + 1);
    }

    private static Regex BuildParseRegex(string format, int year)
    {
        var pattern = new StringBuilder("^");
        var lastIndex = 0;

        foreach (Match placeholder in PlaceholderRegex.Matches(format))
        {
            pattern.Append(Regex.Escape(format[lastIndex..placeholder.Index]));

            pattern.Append(placeholder.Groups[1].Value == "yyyy"
                ? Regex.Escape(year.ToString("0000", CultureInfo.InvariantCulture))
                : @"(\d+)");

            lastIndex = placeholder.Index + placeholder.Length;
        }

        pattern.Append(Regex.Escape(format[lastIndex..]));
        pattern.Append('$');

        return new Regex(pattern.ToString(), RegexOptions.IgnoreCase);
    }
}
