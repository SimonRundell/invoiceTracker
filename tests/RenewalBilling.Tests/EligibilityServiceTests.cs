using RenewalBilling.Models;
using RenewalBilling.Services;

namespace RenewalBilling.Tests;

public class EligibilityServiceTests
{
    private static Client MakeClient(DateTime? dueDate, bool active = true, string id = "C1") => new()
    {
        ClientId = id,
        ClientName = "Test Client",
        DueDate = dueDate,
        Active = active,
    };

    [Fact]
    public void Evaluate_InactiveClient_IsSkipped()
    {
        var client = MakeClient(dueDate: new DateTime(2026, 10, 10), active: false);

        var result = EligibilityService.Evaluate(client, hasItems: true, runDate: new DateTime(2026, 10, 4), leadDays: 21, existingInvoices: []);

        Assert.False(result.IsEligible);
        Assert.Contains("not active", result.SkipReason, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Evaluate_BlankDueDate_IsSkipped()
    {
        var client = MakeClient(dueDate: null);

        var result = EligibilityService.Evaluate(client, hasItems: true, runDate: new DateTime(2026, 10, 4), leadDays: 21, existingInvoices: []);

        Assert.False(result.IsEligible);
        Assert.Contains("blank", result.SkipReason, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Evaluate_NoItems_IsSkippedWithWarning()
    {
        var client = MakeClient(dueDate: new DateTime(2026, 10, 10));

        var result = EligibilityService.Evaluate(client, hasItems: false, runDate: new DateTime(2026, 10, 4), leadDays: 21, existingInvoices: []);

        Assert.False(result.IsEligible);
        Assert.Contains("Items", result.SkipReason);
    }

    [Fact]
    public void Evaluate_DueDateBeyondLeadTime_IsNotDueAndNotAnError()
    {
        var client = MakeClient(dueDate: new DateTime(2026, 12, 1));

        var result = EligibilityService.Evaluate(client, hasItems: true, runDate: new DateTime(2026, 10, 4), leadDays: 21, existingInvoices: []);

        Assert.False(result.IsEligible);
        Assert.Null(result.SkipReason);
    }

    [Fact]
    public void Evaluate_DueDateWithinLeadTime_IsEligibleAndDue()
    {
        var client = MakeClient(dueDate: new DateTime(2026, 10, 20));

        var result = EligibilityService.Evaluate(client, hasItems: true, runDate: new DateTime(2026, 10, 4), leadDays: 21, existingInvoices: []);

        Assert.True(result.IsEligible);
        Assert.Equal(CandidateStatus.Due, result.Status);
    }

    [Fact]
    public void Evaluate_DueDateInThePast_IsEligibleAndOverdue()
    {
        var client = MakeClient(dueDate: new DateTime(2026, 9, 1));

        var result = EligibilityService.Evaluate(client, hasItems: true, runDate: new DateTime(2026, 10, 4), leadDays: 21, existingInvoices: []);

        Assert.True(result.IsEligible);
        Assert.Equal(CandidateStatus.Overdue, result.Status);
    }

    [Fact]
    public void Evaluate_OverdueWithNoInvoiceLogged_IsStillEligible()
    {
        // An overdue client with nothing in the Invoices sheet for this due date must still be
        // offered, flagged as overdue, rather than silently dropped.
        var client = MakeClient(dueDate: new DateTime(2026, 8, 15));
        var invoices = new[]
        {
            new InvoiceRecord
            {
                InvoiceNumber = "INV-2025-0001",
                ClientId = "C1",
                PeriodStart = new DateTime(2025, 8, 15), // a different renewal, not this one
            },
        };

        var result = EligibilityService.Evaluate(client, hasItems: true, runDate: new DateTime(2026, 10, 4), leadDays: 21, existingInvoices: invoices);

        Assert.True(result.IsEligible);
        Assert.Equal(CandidateStatus.Overdue, result.Status);
    }

    [Fact]
    public void Evaluate_AlreadyInvoicedForSameDueDate_IsSkipped()
    {
        var dueDate = new DateTime(2026, 10, 10);
        var client = MakeClient(dueDate: dueDate);
        var invoices = new[]
        {
            new InvoiceRecord
            {
                InvoiceNumber = "INV-2026-0005",
                ClientId = "C1",
                PeriodStart = dueDate,
            },
        };

        var result = EligibilityService.Evaluate(client, hasItems: true, runDate: new DateTime(2026, 10, 4), leadDays: 21, existingInvoices: invoices);

        Assert.False(result.IsEligible);
        Assert.Contains("invoiced", result.SkipReason, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Evaluate_InvoiceLoggedForDifferentClient_DoesNotSuppressThisClient()
    {
        var dueDate = new DateTime(2026, 10, 10);
        var client = MakeClient(dueDate: dueDate, id: "C1");
        var invoices = new[]
        {
            new InvoiceRecord { InvoiceNumber = "INV-2026-0005", ClientId = "C2", PeriodStart = dueDate },
        };

        var result = EligibilityService.Evaluate(client, hasItems: true, runDate: new DateTime(2026, 10, 4), leadDays: 21, existingInvoices: invoices);

        Assert.True(result.IsEligible);
    }

    [Fact]
    public void ComputePeriodEnd_LeapDayDueDate_ClampsThenSubtractsOneDay()
    {
        // AddMonths clamps 29 Feb 2024 + 12 months to 28 Feb 2025, then the period end rule
        // subtracts one more day, landing on 27 Feb 2025.
        var periodEnd = EligibilityService.ComputePeriodEnd(new DateTime(2024, 2, 29), renewalMonths: 12);

        Assert.Equal(new DateTime(2025, 2, 27), periodEnd);
    }

    [Fact]
    public void ComputeNextDueDate_LeapDayDueDate_ClampsToTwentyEighthInNonLeapYear()
    {
        var nextDueDate = EligibilityService.ComputeNextDueDate(new DateTime(2024, 2, 29), renewalMonths: 12);

        Assert.Equal(new DateTime(2025, 2, 28), nextDueDate);
    }

    [Fact]
    public void ComputePeriodEnd_LeapYearToLeapYear_KeepsLeapDayOneDayEarlier()
    {
        // 29 Feb 2024 plus 48 months lands back on a leap year (2028), so the period end is the
        // day before: 28 Feb 2028.
        var periodEnd = EligibilityService.ComputePeriodEnd(new DateTime(2024, 2, 29), renewalMonths: 48);

        Assert.Equal(new DateTime(2028, 2, 28), periodEnd);
    }

    [Fact]
    public void ComputePayByDate_WithPayByDays_AddsDaysToInvoiceDate()
    {
        var payBy = EligibilityService.ComputePayByDate(
            invoiceDate: new DateTime(2026, 10, 4), dueDate: new DateTime(2026, 10, 20), payByDays: 14);

        Assert.Equal(new DateTime(2026, 10, 18), payBy);
    }

    [Fact]
    public void ComputePayByDate_WithoutPayByDays_UsesDueDate()
    {
        var dueDate = new DateTime(2026, 10, 20);
        var payBy = EligibilityService.ComputePayByDate(invoiceDate: new DateTime(2026, 10, 4), dueDate: dueDate, payByDays: null);

        Assert.Equal(dueDate, payBy);
    }
}
