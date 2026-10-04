using RenewalBilling.Models;
using RenewalBilling.Services;

namespace RenewalBilling.Tests;

public class UpliftCalculatorTests
{
    [Fact]
    public void ClampPercent_NegativeRpiWithMinimumZero_ClampsToZero()
    {
        var result = UpliftCalculator.ClampPercent(rpiPercent: -1.2m, minimumPercent: 0m, maximumPercent: null);

        Assert.Equal(0m, result);
    }

    [Fact]
    public void ClampPercent_AboveMaximum_ClampsToMaximum()
    {
        var result = UpliftCalculator.ClampPercent(rpiPercent: 12m, minimumPercent: 0m, maximumPercent: 5m);

        Assert.Equal(5m, result);
    }

    [Fact]
    public void ClampPercent_WithinRange_ReturnsUnchanged()
    {
        var result = UpliftCalculator.ClampPercent(rpiPercent: 3.5m, minimumPercent: 0m, maximumPercent: 10m);

        Assert.Equal(3.5m, result);
    }

    [Fact]
    public void CalculateProposedUnitCost_NegativeRpiWithMinimumZero_LeavesPriceUnchanged()
    {
        var proposed = UpliftCalculator.CalculateProposedUnitCost(
            currentUnitCost: 100m, rpiPercent: -2m, roundTo: 1.00m, minimumPercent: 0m, maximumPercent: null);

        Assert.Equal(100m, proposed);
    }

    [Theory]
    [InlineData(100, 2.5, 1.00, 103)]   // 102.5 rounds away from zero at the midpoint
    [InlineData(100, 1.5, 1.00, 102)]   // 101.5 rounds away from zero at the midpoint
    [InlineData(50, 2.5, 1.00, 51)]     // 51.25 rounds down, not a midpoint
    public void CalculateProposedUnitCost_RoundsMidpointsAwayFromZero(
        decimal currentUnitCost, decimal rpiPercent, decimal roundTo, decimal expected)
    {
        var proposed = UpliftCalculator.CalculateProposedUnitCost(
            currentUnitCost, rpiPercent, roundTo, minimumPercent: 0m, maximumPercent: null);

        Assert.Equal(expected, proposed);
    }

    [Fact]
    public void CalculateProposedUnitCost_RoundsToConfiguredStep()
    {
        var proposed = UpliftCalculator.CalculateProposedUnitCost(
            currentUnitCost: 99m, rpiPercent: 3m, roundTo: 5.00m, minimumPercent: 0m, maximumPercent: null);

        // 99 * 1.03 = 101.97 -> 101.97 / 5 = 20.394 -> rounds to 20 -> 20 * 5 = 100
        Assert.Equal(100m, proposed);
    }

    [Fact]
    public void BuildProposedItems_AppliesSameClampedPercentToEveryItem()
    {
        var items = new[]
        {
            new BillingItem { ClientId = "C1", Description = "Licence", Quantity = 1, UnitCost = 100m },
            new BillingItem { ClientId = "C1", Description = "Support", Quantity = 2, UnitCost = 50m },
        };

        var proposed = UpliftCalculator.BuildProposedItems(
            items, rpiPercent: -5m, roundTo: 1.00m, minimumPercent: 0m, maximumPercent: null);

        Assert.All(proposed, line => Assert.Equal(0m, line.UpliftPercentApplied));
        Assert.Equal(100m, proposed[0].ProposedUnitCost);
        Assert.Equal(50m, proposed[1].ProposedUnitCost);
        Assert.Equal(100m, proposed[1].LineTotal); // quantity 2 * unit cost 50
    }

    [Theory]
    [InlineData(100, 20, 20.00)]
    [InlineData(33.33, 20, 6.67)] // rounds to the nearest penny, away from zero
    [InlineData(100, 0, 0.00)]
    public void CalculateVat_RoundsToNearestPenny(decimal net, decimal vatRatePercent, decimal expected)
    {
        var vat = UpliftCalculator.CalculateVat(net, vatRatePercent);

        Assert.Equal(expected, vat);
    }
}
