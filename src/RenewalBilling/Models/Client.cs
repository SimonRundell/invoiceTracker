namespace RenewalBilling.Models;

/// <summary>
/// One row from the Clients worksheet.
/// </summary>
public sealed class Client
{
    /// <summary>The client's unique identifier.</summary>
    public required string ClientId { get; init; }

    /// <summary>The client's display name.</summary>
    public string ClientName { get; init; } = string.Empty;

    /// <summary>The name of the person to address correspondence to.</summary>
    public string ContactName { get; init; } = string.Empty;

    /// <summary>First address line.</summary>
    public string Address1 { get; init; } = string.Empty;

    /// <summary>Second address line.</summary>
    public string Address2 { get; init; } = string.Empty;

    /// <summary>Town or city.</summary>
    public string Town { get; init; } = string.Empty;

    /// <summary>Postcode.</summary>
    public string Postcode { get; init; } = string.Empty;

    /// <summary>Contact email address.</summary>
    public string Email { get; init; } = string.Empty;

    /// <summary>The date the current renewal falls due. Null means the workbook left it blank.</summary>
    public DateTime? DueDate { get; init; }

    /// <summary>The date the client last paid. Null if never recorded.</summary>
    public DateTime? LastPaidDate { get; init; }

    /// <summary>Whether the client is active. Inactive clients are never eligible for renewal.</summary>
    public bool Active { get; init; } = true;
}
