namespace RenewalBilling.Models;

/// <summary>
/// One row from the append-only Invoices worksheet, logging an invoice that was created.
/// </summary>
public sealed class InvoiceRecord
{
    /// <summary>The invoice number, built from the configured format.</summary>
    public required string InvoiceNumber { get; init; }

    /// <summary>The id of the client that was invoiced.</summary>
    public required string ClientId { get; init; }

    /// <summary>The date the invoice was created.</summary>
    public DateTime InvoiceDate { get; init; }

    /// <summary>
    /// The date payment is due. This is the worksheet's "DueDate" column, which holds the
    /// invoice's pay-by date and is distinct from a client's renewal due date.
    /// </summary>
    public DateTime PayByDate { get; init; }

    /// <summary>The first day of the renewal period this invoice covers. Equal to the client's renewal due date at the time of invoicing.</summary>
    public DateTime PeriodStart { get; init; }

    /// <summary>The last day of the renewal period this invoice covers.</summary>
    public DateTime PeriodEnd { get; init; }

    /// <summary>The net amount invoiced.</summary>
    public decimal Net { get; init; }

    /// <summary>The VAT amount invoiced.</summary>
    public decimal Vat { get; init; }

    /// <summary>The total amount invoiced, net plus VAT.</summary>
    public decimal Total { get; init; }

    /// <summary>The RPI percentage used to calculate this invoice's uplift.</summary>
    public decimal RpiPercent { get; init; }

    /// <summary>Path to the saved .docx copy of the invoice.</summary>
    public string DocxPath { get; init; } = string.Empty;

    /// <summary>Path to the saved PDF copy of the invoice.</summary>
    public string PdfPath { get; init; } = string.Empty;
}
