using RenewalBilling.Models;

namespace RenewalBilling.Services;

/// <summary>
/// The outcome of checking whether one client is eligible for renewal this run.
/// </summary>
/// <param name="Client">The client that was evaluated.</param>
/// <param name="IsEligible">Whether the client should be offered for renewal.</param>
/// <param name="Status">The candidate's status when eligible, otherwise null.</param>
/// <param name="SkipReason">A human readable reason the client was skipped, otherwise null.</param>
public sealed record EligibilityResult(Client Client, bool IsEligible, CandidateStatus? Status, string? SkipReason)
{
    /// <summary>Creates a result for a client that was skipped, with the given reason.</summary>
    public static EligibilityResult Skipped(Client client, string reason) => new(client, false, null, reason);

    /// <summary>Creates a result for a client that is not due yet. Not an error, just not a candidate.</summary>
    public static EligibilityResult NotDue(Client client) => new(client, false, null, null);

    /// <summary>Creates a result for a client that is eligible, with the given status.</summary>
    public static EligibilityResult Eligible(Client client, CandidateStatus status) => new(client, true, status, null);
}

/// <summary>
/// Pure logic that decides whether a client is due for renewal, and the dates that go with a
/// renewal. Holds no state and touches no file, network or COM resource, so it can be unit
/// tested directly.
/// </summary>
public static class EligibilityService
{
    /// <summary>
    /// Evaluates one client against the eligibility rules: must be active, must have a due date,
    /// must have at least one billing item, the due date must fall within the lead time, and
    /// there must be no existing invoice already logged for the same renewal.
    /// </summary>
    /// <param name="client">The client to evaluate.</param>
    /// <param name="hasItems">Whether the client has at least one row in the Items worksheet.</param>
    /// <param name="runDate">The date the run is treating as "today".</param>
    /// <param name="leadDays">How many days before the due date a renewal becomes eligible.</param>
    /// <param name="existingInvoices">
    /// Invoice records already logged. A client already invoiced for the same renewal (matched
    /// on <see cref="InvoiceRecord.PeriodStart"/>, which holds the due date at the time that
    /// invoice was raised) is excluded.
    /// </param>
    /// <returns>The eligibility outcome for this client.</returns>
    public static EligibilityResult Evaluate(
        Client client,
        bool hasItems,
        DateTime runDate,
        int leadDays,
        IEnumerable<InvoiceRecord> existingInvoices)
    {
        if (!client.Active)
        {
            return EligibilityResult.Skipped(client, "Client is not active.");
        }

        if (client.DueDate is null)
        {
            return EligibilityResult.Skipped(client, "DueDate is blank.");
        }

        if (!hasItems)
        {
            return EligibilityResult.Skipped(client, "Client has no Items rows.");
        }

        var dueDate = client.DueDate.Value.Date;
        var runDateOnly = runDate.Date;

        var alreadyInvoiced = existingInvoices.Any(record =>
            string.Equals(record.ClientId, client.ClientId, StringComparison.OrdinalIgnoreCase)
            && record.PeriodStart.Date == dueDate);

        if (alreadyInvoiced)
        {
            return EligibilityResult.Skipped(client, "Already invoiced for this renewal.");
        }

        if (dueDate > runDateOnly.AddDays(leadDays))
        {
            return EligibilityResult.NotDue(client);
        }

        var status = dueDate < runDateOnly ? CandidateStatus.Overdue : CandidateStatus.Due;
        return EligibilityResult.Eligible(client, status);
    }

    /// <summary>
    /// Calculates the last day of a renewal period: one day before the due date advanced by the
    /// renewal length. Uses <see cref="DateTime.AddMonths(int)"/>, which clamps to the last valid
    /// day of the target month, so a due date of 29 February correctly yields 28 February in a
    /// non-leap target year.
    /// </summary>
    /// <param name="dueDate">The due date that starts the renewal period.</param>
    /// <param name="renewalMonths">How many months the renewal period covers.</param>
    /// <returns>The last day of the renewal period.</returns>
    public static DateTime ComputePeriodEnd(DateTime dueDate, int renewalMonths)
    {
        return dueDate.AddMonths(renewalMonths).AddDays(-1);
    }

    /// <summary>
    /// Calculates the next due date after a renewal is approved.
    /// </summary>
    /// <param name="dueDate">The due date of the renewal just approved.</param>
    /// <param name="renewalMonths">How many months the renewal period covers.</param>
    /// <returns>The due date for the following renewal.</returns>
    public static DateTime ComputeNextDueDate(DateTime dueDate, int renewalMonths)
    {
        return dueDate.AddMonths(renewalMonths);
    }

    /// <summary>
    /// Calculates the pay-by date for an invoice.
    /// </summary>
    /// <param name="invoiceDate">The date the invoice is created.</param>
    /// <param name="dueDate">The client's renewal due date.</param>
    /// <param name="payByDays">Days after the invoice date that payment is due, or null to use the due date.</param>
    /// <returns>The pay-by date.</returns>
    public static DateTime ComputePayByDate(DateTime invoiceDate, DateTime dueDate, int? payByDays)
    {
        return payByDays.HasValue ? invoiceDate.AddDays(payByDays.Value) : dueDate;
    }
}
