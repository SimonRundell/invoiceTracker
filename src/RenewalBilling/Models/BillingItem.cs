namespace RenewalBilling.Models;

/// <summary>
/// One row from the Items worksheet: a single billable line that belongs to a client.
/// </summary>
public sealed class BillingItem
{
    /// <summary>The id of the client this item belongs to.</summary>
    public required string ClientId { get; init; }

    /// <summary>What the line is for.</summary>
    public string Description { get; init; } = string.Empty;

    /// <summary>How many units are billed.</summary>
    public decimal Quantity { get; init; }

    /// <summary>The current annual price per unit.</summary>
    public decimal UnitCost { get; init; }

    /// <summary>The unit cost before the last uplift, if known.</summary>
    public decimal? PreviousUnitCost { get; init; }

    /// <summary>The uplift percentage applied last time, if known.</summary>
    public decimal? LastUpliftPercent { get; init; }
}
