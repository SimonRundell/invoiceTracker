namespace RenewalBilling.Models;

/// <summary>
/// Whether a candidate's renewal is due within the lead time, or has already passed.
/// </summary>
public enum CandidateStatus
{
    /// <summary>The renewal falls due within the configured lead time, but has not passed yet.</summary>
    Due,

    /// <summary>The renewal due date has already passed with no invoice logged against it.</summary>
    Overdue,
}

/// <summary>
/// One billing item with its current price alongside the price proposed for the new renewal.
/// </summary>
/// <param name="Item">The underlying billing item, including its current unit cost.</param>
/// <param name="ProposedUnitCost">The new unit cost proposed for the renewal.</param>
/// <param name="UpliftPercentApplied">The uplift percentage actually applied, after clamping.</param>
public sealed record ProposedLineItem(BillingItem Item, decimal ProposedUnitCost, decimal UpliftPercentApplied)
{
    /// <summary>The line total: quantity multiplied by the proposed unit cost.</summary>
    public decimal LineTotal => Item.Quantity * ProposedUnitCost;
}

/// <summary>
/// A client who is eligible for renewal this run, together with the proposed invoice for them.
/// </summary>
public sealed class InvoiceCandidate
{
    /// <summary>The client being renewed.</summary>
    public required Client Client { get; init; }

    /// <summary>Whether the renewal is due or already overdue.</summary>
    public required CandidateStatus Status { get; init; }

    /// <summary>The proposed line items for the new invoice.</summary>
    public required IReadOnlyList<ProposedLineItem> Items { get; init; }

    /// <summary>The RPI percentage used to calculate the proposed prices.</summary>
    public required decimal RpiPercentUsed { get; init; }

    /// <summary>The date the invoice is created.</summary>
    public required DateTime InvoiceDate { get; init; }

    /// <summary>The first day of the new renewal period. Equal to the client's current due date.</summary>
    public required DateTime PeriodStart { get; init; }

    /// <summary>The last day of the new renewal period.</summary>
    public required DateTime PeriodEnd { get; init; }

    /// <summary>The date payment is due.</summary>
    public required DateTime PayByDate { get; init; }

    /// <summary>The VAT amount for this invoice.</summary>
    public required decimal Vat { get; init; }

    /// <summary>The invoice number, assigned when the user approves this candidate.</summary>
    public string? InvoiceNumber { get; set; }

    /// <summary>The net amount: the sum of every line's total.</summary>
    public decimal Net => Items.Sum(i => i.LineTotal);

    /// <summary>The total amount: net plus VAT.</summary>
    public decimal Total => Net + Vat;
}
