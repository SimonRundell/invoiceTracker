using System.Text.Json;
using System.Text.Json.Serialization;

namespace RenewalBilling;

/// <summary>
/// Settings for the ONS RPI lookup.
/// </summary>
public sealed class RpiConfig
{
    /// <summary>The ONS dataset URL to fetch the latest RPI figure from.</summary>
    public string Url { get; set; } = string.Empty;

    /// <summary>How long to wait for the ONS request before giving up.</summary>
    public int TimeoutSeconds { get; set; } = 15;

    /// <summary>The percentage to use when there is no cached value and the online fetch fails.</summary>
    public decimal FallbackPercent { get; set; }

    /// <summary>File name (in the exe folder) used to cache the last successful RPI result.</summary>
    public string CacheFile { get; set; } = "rpi-cache.json";
}

/// <summary>
/// Settings that control how the proposed uplift is calculated.
/// </summary>
public sealed class UpliftConfig
{
    /// <summary>The nearest amount to round proposed prices to, for example 1.00.</summary>
    public decimal RoundTo { get; set; } = 1.00m;

    /// <summary>The lowest uplift percentage allowed, applied before rounding.</summary>
    public decimal MinimumPercent { get; set; }

    /// <summary>The highest uplift percentage allowed, or null for no cap.</summary>
    public decimal? MaximumPercent { get; set; }
}

/// <summary>
/// Names of the three worksheets in the client workbook.
/// </summary>
public sealed class SheetsConfig
{
    /// <summary>Name of the clients worksheet.</summary>
    public string Clients { get; set; } = "Clients";

    /// <summary>Name of the billable items worksheet.</summary>
    public string Items { get; set; } = "Items";

    /// <summary>Name of the invoice log worksheet.</summary>
    public string Invoices { get; set; } = "Invoices";
}

/// <summary>
/// Header names for the Clients worksheet.
/// </summary>
public sealed class ClientColumnsConfig
{
    /// <summary>Header for the client id column.</summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>Header for the client name column.</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>Header for the contact name column.</summary>
    public string Contact { get; set; } = string.Empty;

    /// <summary>Header for the first address line column.</summary>
    public string Address1 { get; set; } = string.Empty;

    /// <summary>Header for the second address line column.</summary>
    public string Address2 { get; set; } = string.Empty;

    /// <summary>Header for the town column.</summary>
    public string Town { get; set; } = string.Empty;

    /// <summary>Header for the postcode column.</summary>
    public string Postcode { get; set; } = string.Empty;

    /// <summary>Header for the email column.</summary>
    public string Email { get; set; } = string.Empty;

    /// <summary>Header for the renewal due date column.</summary>
    public string DueDate { get; set; } = string.Empty;

    /// <summary>Header for the last paid date column.</summary>
    public string LastPaidDate { get; set; } = string.Empty;

    /// <summary>Header for the active flag column.</summary>
    public string Active { get; set; } = string.Empty;
}

/// <summary>
/// Header names for the Items worksheet.
/// </summary>
public sealed class ItemColumnsConfig
{
    /// <summary>Header for the owning client id column.</summary>
    public string ClientId { get; set; } = string.Empty;

    /// <summary>Header for the line description column.</summary>
    public string Description { get; set; } = string.Empty;

    /// <summary>Header for the quantity column.</summary>
    public string Quantity { get; set; } = string.Empty;

    /// <summary>Header for the current unit cost column.</summary>
    public string UnitCost { get; set; } = string.Empty;

    /// <summary>Header for the previous unit cost column.</summary>
    public string PreviousUnitCost { get; set; } = string.Empty;

    /// <summary>Header for the last uplift percentage column.</summary>
    public string LastUpliftPercent { get; set; } = string.Empty;
}

/// <summary>
/// Header names for the Invoices worksheet.
/// </summary>
public sealed class InvoiceColumnsConfig
{
    /// <summary>Header for the invoice number column.</summary>
    public string InvoiceNumber { get; set; } = string.Empty;

    /// <summary>Header for the client id column.</summary>
    public string ClientId { get; set; } = string.Empty;

    /// <summary>Header for the invoice date column.</summary>
    public string InvoiceDate { get; set; } = string.Empty;

    /// <summary>Header for the pay-by date column.</summary>
    public string DueDate { get; set; } = string.Empty;

    /// <summary>Header for the renewal period start column.</summary>
    public string PeriodStart { get; set; } = string.Empty;

    /// <summary>Header for the renewal period end column.</summary>
    public string PeriodEnd { get; set; } = string.Empty;

    /// <summary>Header for the net amount column.</summary>
    public string Net { get; set; } = string.Empty;

    /// <summary>Header for the VAT amount column.</summary>
    public string Vat { get; set; } = string.Empty;

    /// <summary>Header for the total amount column.</summary>
    public string Total { get; set; } = string.Empty;

    /// <summary>Header for the RPI percentage used column.</summary>
    public string RpiPercent { get; set; } = string.Empty;

    /// <summary>Header for the saved .docx path column.</summary>
    public string DocxPath { get; set; } = string.Empty;

    /// <summary>Header for the saved PDF path column.</summary>
    public string PdfPath { get; set; } = string.Empty;
}

/// <summary>
/// The worksheet header maps for all three sheets.
/// </summary>
public sealed class ColumnsConfig
{
    /// <summary>Column headers for the Clients worksheet.</summary>
    public ClientColumnsConfig Clients { get; set; } = new();

    /// <summary>Column headers for the Items worksheet.</summary>
    public ItemColumnsConfig Items { get; set; } = new();

    /// <summary>Column headers for the Invoices worksheet.</summary>
    public InvoiceColumnsConfig Invoices { get; set; } = new();
}

/// <summary>
/// The full set of settings read from <c>.config.json</c>.
/// </summary>
public sealed class AppConfig
{
    /// <summary>Path to the client workbook.</summary>
    public string WorkbookPath { get; set; } = string.Empty;

    /// <summary>Path to the Word invoice template.</summary>
    public string TemplatePath { get; set; } = string.Empty;

    /// <summary>Folder that finished invoices are saved into.</summary>
    public string InvoiceFolder { get; set; } = string.Empty;

    /// <summary>Folder that workbook backups are saved into.</summary>
    public string BackupFolder { get; set; } = string.Empty;

    /// <summary>Folder that log files are written to.</summary>
    public string LogFolder { get; set; } = string.Empty;

    /// <summary>How many days before the due date a renewal becomes eligible.</summary>
    public int LeadDays { get; set; } = 21;

    /// <summary>How many months a renewal period covers.</summary>
    public int RenewalMonths { get; set; } = 12;

    /// <summary>Format string used to build invoice numbers, for example <c>INV-{yyyy}-{seq:0000}</c>.</summary>
    public string InvoiceNumberFormat { get; set; } = string.Empty;

    /// <summary>Days after the invoice date that payment is due, or null to use the renewal due date.</summary>
    public int? PayByDays { get; set; }

    /// <summary>VAT rate to apply to the net amount, as a percentage.</summary>
    public decimal VatRatePercent { get; set; }

    /// <summary>RPI lookup settings.</summary>
    public RpiConfig Rpi { get; set; } = new();

    /// <summary>Uplift calculation settings.</summary>
    public UpliftConfig Uplift { get; set; } = new();

    /// <summary>Worksheet name settings.</summary>
    public SheetsConfig Sheets { get; set; } = new();

    /// <summary>Worksheet column header settings.</summary>
    public ColumnsConfig Columns { get; set; } = new();
}

/// <summary>
/// Thrown when <c>.config.json</c> is missing, malformed, or fails validation.
/// </summary>
public sealed class ConfigException : Exception
{
    /// <summary>Creates a new <see cref="ConfigException"/> with a message describing the problem.</summary>
    /// <param name="message">A clear, human readable description of what is wrong.</param>
    public ConfigException(string message) : base(message)
    {
    }
}

/// <summary>
/// Loads and validates <see cref="AppConfig"/> from a JSON file, resolving relative paths
/// against a base folder (normally the exe folder).
/// </summary>
public static class AppConfigLoader
{
    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        ReadCommentHandling = JsonCommentHandling.Skip,
        NumberHandling = JsonNumberHandling.AllowReadingFromString,
    };

    /// <summary>
    /// Reads, parses and validates the config file at <paramref name="path"/>.
    /// </summary>
    /// <param name="path">Full path to the config JSON file.</param>
    /// <param name="baseDirectory">Folder that relative paths in the config are resolved against.</param>
    /// <returns>A fully populated and validated <see cref="AppConfig"/>.</returns>
    /// <exception cref="ConfigException">
    /// Thrown when the file is missing, the JSON cannot be parsed, or a required value is missing.
    /// </exception>
    public static AppConfig Load(string path, string baseDirectory)
    {
        if (!File.Exists(path))
        {
            throw new ConfigException($"Config file not found: {path}");
        }

        string json;
        try
        {
            json = File.ReadAllText(path);
        }
        catch (IOException ex)
        {
            throw new ConfigException($"Could not read config file '{path}': {ex.Message}");
        }

        AppConfig? config;
        try
        {
            config = JsonSerializer.Deserialize<AppConfig>(json, SerializerOptions);
        }
        catch (JsonException ex)
        {
            throw new ConfigException($"Config file '{path}' is not valid JSON: {ex.Message}");
        }

        if (config is null)
        {
            throw new ConfigException($"Config file '{path}' is empty or not a JSON object.");
        }

        config.WorkbookPath = ResolvePath(config.WorkbookPath, baseDirectory);
        config.TemplatePath = ResolvePath(config.TemplatePath, baseDirectory);
        config.InvoiceFolder = ResolvePath(config.InvoiceFolder, baseDirectory);
        config.BackupFolder = ResolvePath(config.BackupFolder, baseDirectory);
        config.LogFolder = ResolvePath(config.LogFolder, baseDirectory);
        config.Rpi.CacheFile = ResolvePath(config.Rpi.CacheFile, baseDirectory);

        Validate(config, path);

        return config;
    }

    private static string ResolvePath(string value, string baseDirectory)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return value;
        }

        return Path.IsPathRooted(value) ? value : Path.GetFullPath(Path.Combine(baseDirectory, value));
    }

    private static void Validate(AppConfig config, string path)
    {
        var problems = new List<string>();

        RequireValue(problems, "workbookPath", config.WorkbookPath);
        RequireValue(problems, "templatePath", config.TemplatePath);
        RequireValue(problems, "invoiceFolder", config.InvoiceFolder);
        RequireValue(problems, "backupFolder", config.BackupFolder);
        RequireValue(problems, "logFolder", config.LogFolder);
        RequireValue(problems, "invoiceNumberFormat", config.InvoiceNumberFormat);

        if (config.LeadDays < 0)
        {
            problems.Add("leadDays must be zero or greater.");
        }

        if (config.RenewalMonths <= 0)
        {
            problems.Add("renewalMonths must be greater than zero.");
        }

        if (config.VatRatePercent < 0)
        {
            problems.Add("vatRatePercent must be zero or greater.");
        }

        RequireValue(problems, "rpi.url", config.Rpi.Url);
        RequireValue(problems, "rpi.cacheFile", config.Rpi.CacheFile);

        if (config.Rpi.TimeoutSeconds <= 0)
        {
            problems.Add("rpi.timeoutSeconds must be greater than zero.");
        }

        if (config.Uplift.RoundTo <= 0)
        {
            problems.Add("uplift.roundTo must be greater than zero.");
        }

        if (config.Uplift.MaximumPercent.HasValue && config.Uplift.MaximumPercent.Value < config.Uplift.MinimumPercent)
        {
            problems.Add("uplift.maximumPercent cannot be less than uplift.minimumPercent.");
        }

        RequireValue(problems, "sheets.clients", config.Sheets.Clients);
        RequireValue(problems, "sheets.items", config.Sheets.Items);
        RequireValue(problems, "sheets.invoices", config.Sheets.Invoices);

        RequireValue(problems, "columns.clients.id", config.Columns.Clients.Id);
        RequireValue(problems, "columns.clients.name", config.Columns.Clients.Name);
        RequireValue(problems, "columns.clients.dueDate", config.Columns.Clients.DueDate);
        RequireValue(problems, "columns.clients.active", config.Columns.Clients.Active);

        RequireValue(problems, "columns.items.clientId", config.Columns.Items.ClientId);
        RequireValue(problems, "columns.items.description", config.Columns.Items.Description);
        RequireValue(problems, "columns.items.quantity", config.Columns.Items.Quantity);
        RequireValue(problems, "columns.items.unitCost", config.Columns.Items.UnitCost);

        RequireValue(problems, "columns.invoices.invoiceNumber", config.Columns.Invoices.InvoiceNumber);
        RequireValue(problems, "columns.invoices.clientId", config.Columns.Invoices.ClientId);
        RequireValue(problems, "columns.invoices.invoiceDate", config.Columns.Invoices.InvoiceDate);
        RequireValue(problems, "columns.invoices.dueDate", config.Columns.Invoices.DueDate);

        if (problems.Count > 0)
        {
            throw new ConfigException($"Config file '{path}' failed validation: {string.Join(" ", problems)}");
        }
    }

    private static void RequireValue(List<string> problems, string fieldName, string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            problems.Add($"{fieldName} is required.");
        }
    }
}
