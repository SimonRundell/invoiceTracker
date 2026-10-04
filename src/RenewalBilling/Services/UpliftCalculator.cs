using RenewalBilling.Models;

namespace RenewalBilling.Services;

/// <summary>
/// Pure calculations for proposed renewal prices, VAT and uplift clamping. Holds no state and
/// touches no file, network or COM resource, so it can be unit tested directly.
/// </summary>
public static class UpliftCalculator
{
    /// <summary>
    /// Clamps a raw RPI percentage between the configured minimum and (optional) maximum.
    /// </summary>
    /// <param name="rpiPercent">The raw RPI percentage.</param>
    /// <param name="minimumPercent">The lowest uplift percentage allowed.</param>
    /// <param name="maximumPercent">The highest uplift percentage allowed, or null for no cap.</param>
    /// <returns>The clamped percentage to apply.</returns>
    public static decimal ClampPercent(decimal rpiPercent, decimal minimumPercent, decimal? maximumPercent)
    {
        var result = Math.Max(rpiPercent, minimumPercent);
        if (maximumPercent.HasValue)
        {
            result = Math.Min(result, maximumPercent.Value);
        }

        return result;
    }

    /// <summary>
    /// Calculates the proposed unit cost for one item: the current cost uplifted by the clamped
    /// RPI percentage, then rounded to the nearest <paramref name="roundTo"/>.
    /// </summary>
    /// <param name="currentUnitCost">The item's current unit cost.</param>
    /// <param name="rpiPercent">The raw RPI percentage to apply.</param>
    /// <param name="roundTo">The amount to round the result to, for example 1.00.</param>
    /// <param name="minimumPercent">The lowest uplift percentage allowed.</param>
    /// <param name="maximumPercent">The highest uplift percentage allowed, or null for no cap.</param>
    /// <returns>The proposed unit cost.</returns>
    public static decimal CalculateProposedUnitCost(
        decimal currentUnitCost,
        decimal rpiPercent,
        decimal roundTo,
        decimal minimumPercent,
        decimal? maximumPercent)
    {
        var clampedPercent = ClampPercent(rpiPercent, minimumPercent, maximumPercent);
        var raw = currentUnitCost * (1 + clampedPercent / 100m);
        var steps = Math.Round(raw / roundTo, 0, MidpointRounding.AwayFromZero);
        return steps * roundTo;
    }

    /// <summary>
    /// Builds the proposed line items for a client's billing items, applying the same clamped
    /// RPI percentage to each one.
    /// </summary>
    /// <param name="items">The client's current billing items.</param>
    /// <param name="rpiPercent">The raw RPI percentage to apply.</param>
    /// <param name="roundTo">The amount to round proposed prices to.</param>
    /// <param name="minimumPercent">The lowest uplift percentage allowed.</param>
    /// <param name="maximumPercent">The highest uplift percentage allowed, or null for no cap.</param>
    /// <returns>One <see cref="ProposedLineItem"/> per input item.</returns>
    public static IReadOnlyList<ProposedLineItem> BuildProposedItems(
        IEnumerable<BillingItem> items,
        decimal rpiPercent,
        decimal roundTo,
        decimal minimumPercent,
        decimal? maximumPercent)
    {
        var clampedPercent = ClampPercent(rpiPercent, minimumPercent, maximumPercent);

        return items
            .Select(item => new ProposedLineItem(
                item,
                CalculateProposedUnitCost(item.UnitCost, rpiPercent, roundTo, minimumPercent, maximumPercent),
                clampedPercent))
            .ToList();
    }

    /// <summary>
    /// Calculates VAT on a net amount, rounded to the nearest penny.
    /// </summary>
    /// <param name="net">The net amount.</param>
    /// <param name="vatRatePercent">The VAT rate, as a percentage.</param>
    /// <returns>The VAT amount, rounded to two decimal places.</returns>
    public static decimal CalculateVat(decimal net, decimal vatRatePercent)
    {
        return Math.Round(net * vatRatePercent / 100m, 2, MidpointRounding.AwayFromZero);
    }
}
