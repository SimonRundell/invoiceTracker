using System.Runtime.InteropServices;
using RenewalBilling.Models;
using RenewalBilling.Services;

namespace RenewalBilling.Forms;

/// <summary>
/// One eligible client for this run, with a working (editable) copy of its proposed line items.
/// Mutated directly by the review form as the user edits prices or the RPI percentage changes.
/// </summary>
public sealed class CandidateRow
{
    /// <summary>The client being renewed.</summary>
    public required Client Client { get; init; }

    /// <summary>Whether the renewal is due or already overdue.</summary>
    public required CandidateStatus Status { get; init; }

    /// <summary>The client's current billing items, unchanged.</summary>
    public required IReadOnlyList<BillingItem> OriginalItems { get; init; }

    /// <summary>The first day of the new renewal period: the client's current due date.</summary>
    public required DateTime PeriodStart { get; init; }

    /// <summary>The last day of the new renewal period.</summary>
    public required DateTime PeriodEnd { get; init; }

    /// <summary>The date payment is due.</summary>
    public required DateTime PayByDate { get; init; }

    /// <summary>The configured VAT rate, as a percentage.</summary>
    public required decimal VatRatePercent { get; init; }

    /// <summary>The working set of proposed line items, recalculated whenever the RPI percentage changes.</summary>
    public List<ProposedLineItem> ProposedItems { get; set; } = new();

    /// <summary>Whether the user has approved this renewal. Defaults to checked.</summary>
    public bool Approved { get; set; } = true;

    /// <summary>The net amount: the sum of every proposed line's total.</summary>
    public decimal Net => ProposedItems.Sum(i => i.LineTotal);

    /// <summary>The VAT amount on the net total.</summary>
    public decimal Vat => UpliftCalculator.CalculateVat(Net, VatRatePercent);

    /// <summary>The total amount: net plus VAT.</summary>
    public decimal Total => Net + Vat;
}

/// <summary>
/// The result of the discovery phase: opening the workbook, fetching RPI, and finding candidates.
/// </summary>
/// <param name="ExitCode">What the process should exit with if the run stops here.</param>
/// <param name="ErrorMessage">The error to log, if <paramref name="ExitCode"/> is not <see cref="RenewalBilling.ExitCode.Ok"/>.</param>
/// <param name="Rpi">The RPI result for this run, or null if discovery failed before it was fetched.</param>
/// <param name="Candidates">Every eligible client found.</param>
public sealed record DiscoveryOutcome(ExitCode ExitCode, string? ErrorMessage, RpiResult? Rpi, IReadOnlyList<CandidateRow> Candidates);

/// <summary>
/// Progress for one invoice during the build phase.
/// </summary>
public sealed record BuildProgressEventArgs(int Completed, int Total, string ClientName);

/// <summary>
/// One line in the finished build summary.
/// </summary>
public sealed record BuildResultItem(string ClientName, string? InvoiceNumber, string? Reason);

/// <summary>
/// The outcome of the build phase, once every approved candidate has been attempted.
/// </summary>
public sealed record BuildCompleteEventArgs(
    IReadOnlyList<BuildResultItem> Created,
    IReadOnlyList<BuildResultItem> Failed,
    int SkippedCount,
    string InvoiceFolder);

/// <summary>
/// Owns every Excel and Word COM call for one scheduled run, on a single dedicated STA thread,
/// so the review form's UI thread never blocks on Office automation and no COM object is ever
/// touched from a thread other than the one that created it.
///
/// Lifecycle: <see cref="Start"/> opens the workbook, fetches RPI and finds candidates, then
/// signals <see cref="WaitForDiscovery"/>. The caller reads <see cref="Outcome"/> and, if there
/// are candidates, shows the review form. The form later calls either <see cref="RequestBuild"/>
/// (build the approved invoices, write them back, then finish) or <see cref="RequestClose"/>
/// (write nothing). Either way the underlying thread then exits, quitting Excel and Word; the
/// caller must call <see cref="WaitForExit"/> before the process exits, so that cleanup always
/// finishes and no Office process is left behind.
/// </summary>
public sealed class ReviewRunWorker
{
    private readonly AppConfig _config;
    private readonly IAppLogger _logger;
    private readonly DateTime _runDate;
    private readonly Thread _thread;
    private readonly ManualResetEventSlim _discoveryDone = new(false);
    private readonly ManualResetEventSlim _buildRequested = new(false);
    private IReadOnlyList<CandidateRow>? _approvedRows;
    private volatile bool _closeRequested;

    /// <summary>The result of the discovery phase. Valid only after <see cref="WaitForDiscovery"/> returns.</summary>
    public DiscoveryOutcome Outcome { get; private set; } = null!;

    /// <summary>Raised on the worker thread as each approved invoice is attempted. Subscribers must marshal to the UI thread themselves.</summary>
    public event Action<BuildProgressEventArgs>? ProgressChanged;

    /// <summary>Raised on the worker thread once the build phase finishes. Subscribers must marshal to the UI thread themselves.</summary>
    public event Action<BuildCompleteEventArgs>? BuildCompleted;

    /// <summary>
    /// Creates a worker for one run. Nothing happens until <see cref="Start"/> is called.
    /// </summary>
    /// <param name="config">The loaded application config.</param>
    /// <param name="logger">Where to log progress, warnings and errors.</param>
    /// <param name="runDate">The date this run is treating as "today".</param>
    public ReviewRunWorker(AppConfig config, IAppLogger logger, DateTime runDate)
    {
        _config = config;
        _logger = logger;
        _runDate = runDate;
        _thread = new Thread(Run) { IsBackground = false };
        _thread.SetApartmentState(ApartmentState.STA);
    }

    /// <summary>Starts the dedicated worker thread.</summary>
    public void Start() => _thread.Start();

    /// <summary>Blocks until the discovery phase (open, fetch RPI, find candidates) has finished.</summary>
    public void WaitForDiscovery() => _discoveryDone.Wait();

    /// <summary>
    /// Tells the worker thread to build invoices for the given rows, write them back to the
    /// workbook, and then finish. Only one build is supported per run.
    /// </summary>
    /// <param name="approvedRows">The candidate rows the user approved, in the order to process them.</param>
    public void RequestBuild(IReadOnlyList<CandidateRow> approvedRows)
    {
        _approvedRows = approvedRows;
        _buildRequested.Set();
    }

    /// <summary>Tells the worker thread to finish without writing anything.</summary>
    public void RequestClose()
    {
        _closeRequested = true;
        _buildRequested.Set();
    }

    /// <summary>Blocks until the worker thread has finished and every Office process it started has quit.</summary>
    public void WaitForExit() => _thread.Join();

    private void Run()
    {
        try
        {
            RunInternal();
        }
        catch (Exception ex)
        {
            _logger.Error($"Unexpected failure in the review worker: {ex}");
            if (!_discoveryDone.IsSet)
            {
                Outcome = new DiscoveryOutcome(ExitCode.UnexpectedFailure, ex.Message, null, []);
                _discoveryDone.Set();
            }
        }
    }

    private void RunInternal()
    {
        using var repo = new ExcelRepository(_config.WorkbookPath, _config.BackupFolder, _config.Sheets, _config.Columns, _logger);
        WorkbookSnapshot snapshot;

        try
        {
            snapshot = repo.Open();
        }
        catch (WorkbookException ex)
        {
            Outcome = new DiscoveryOutcome(ExitCode.ConfigOrWorkbookError, ex.Message, null, []);
            _discoveryDone.Set();
            return;
        }
        catch (COMException ex)
        {
            Outcome = new DiscoveryOutcome(ExitCode.OfficeNotAvailable, $"Office automation is not available: {ex.Message}", null, []);
            _discoveryDone.Set();
            return;
        }

        var rpiService = new RpiService(_config.Rpi, _logger);
        var rpi = rpiService.GetLatestRpiAsync().GetAwaiter().GetResult();
        _logger.Info($"RPI result: {rpi.Percent}% ({rpi.PeriodLabel}), source {rpi.Source}.");
        if (rpi.Source != RpiSource.Online)
        {
            _logger.Warn($"RPI did not come from a live ONS fetch (source: {rpi.Source}).");
        }

        var itemsByClient = snapshot.Items
            .GroupBy(item => item.ClientId, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => (IReadOnlyList<BillingItem>)g.ToList(), StringComparer.OrdinalIgnoreCase);

        var rows = new List<CandidateRow>();
        foreach (var client in snapshot.Clients)
        {
            var hasItems = itemsByClient.ContainsKey(client.ClientId);
            var result = EligibilityService.Evaluate(client, hasItems, _runDate, _config.LeadDays, snapshot.Invoices);

            if (result.IsEligible)
            {
                var originalItems = itemsByClient.TryGetValue(client.ClientId, out var list) ? list : Array.Empty<BillingItem>();
                var proposed = UpliftCalculator.BuildProposedItems(
                    originalItems, rpi.Percent, _config.Uplift.RoundTo, _config.Uplift.MinimumPercent, _config.Uplift.MaximumPercent);
                var dueDate = client.DueDate!.Value;

                rows.Add(new CandidateRow
                {
                    Client = client,
                    Status = result.Status!.Value,
                    OriginalItems = originalItems,
                    PeriodStart = dueDate,
                    PeriodEnd = EligibilityService.ComputePeriodEnd(dueDate, _config.RenewalMonths),
                    PayByDate = EligibilityService.ComputePayByDate(_runDate, dueDate, _config.PayByDays),
                    ProposedItems = proposed.ToList(),
                    VatRatePercent = _config.VatRatePercent,
                });

                _logger.Info($"Candidate found: {client.ClientName} ({client.ClientId}), {result.Status}, due {dueDate:dd MMM yyyy}.");
            }
            else if (result.SkipReason is not null)
            {
                _logger.Warn($"Skipped {client.ClientName} ({client.ClientId}): {result.SkipReason}");
            }
        }

        Outcome = new DiscoveryOutcome(ExitCode.Ok, null, rpi, rows);
        _discoveryDone.Set();

        if (rows.Count == 0)
        {
            _logger.Info("No candidates found.");
            return;
        }

        _buildRequested.Wait();

        if (_closeRequested || _approvedRows is null)
        {
            _logger.Info("Review form closed without creating any invoices. Nothing was written.");
            return;
        }

        RunBuild(repo, snapshot, _approvedRows, rows.Count);
    }

    private void RunBuild(ExcelRepository repo, WorkbookSnapshot snapshot, IReadOnlyList<CandidateRow> approvedRows, int totalCandidateCount)
    {
        var created = new List<BuildResultItem>();
        var failed = new List<BuildResultItem>();
        var approvals = new List<ApprovedRenewal>();

        var nextSequence = InvoiceNumberService.FindHighestSequence(
            _config.InvoiceNumberFormat, _runDate.Year, snapshot.Invoices.Select(i => i.InvoiceNumber)) + 1;

        using var wordBuilder = new WordInvoiceBuilder(_logger);

        var completed = 0;
        foreach (var row in approvedRows)
        {
            completed++;
            ProgressChanged?.Invoke(new BuildProgressEventArgs(completed, approvedRows.Count, row.Client.ClientName));

            var invoiceNumber = InvoiceNumberService.Render(_config.InvoiceNumberFormat, _runDate.Year, nextSequence);
            nextSequence++;

            var candidate = new InvoiceCandidate
            {
                Client = row.Client,
                Status = row.Status,
                Items = row.ProposedItems,
                RpiPercentUsed = row.ProposedItems.Count > 0 ? row.ProposedItems[0].UpliftPercentApplied : 0m,
                InvoiceDate = _runDate,
                PeriodStart = row.PeriodStart,
                PeriodEnd = row.PeriodEnd,
                PayByDate = row.PayByDate,
                Vat = row.Vat,
                InvoiceNumber = invoiceNumber,
            };

            try
            {
                var files = wordBuilder.CreateInvoice(candidate, _config.TemplatePath, _config.InvoiceFolder);
                approvals.Add(new ApprovedRenewal(candidate, files.DocxPath, files.PdfPath));
                created.Add(new BuildResultItem(row.Client.ClientName, invoiceNumber, null));
                _logger.Info($"Invoice created: {invoiceNumber} for {row.Client.ClientName} ({row.Client.ClientId}).");

                foreach (var tag in files.UnknownTags)
                {
                    _logger.Warn($"Invoice {invoiceNumber}: template tag '{tag}' was never replaced.");
                }
            }
            catch (Exception ex) when (ex is InvoiceBuildException or COMException)
            {
                failed.Add(new BuildResultItem(row.Client.ClientName, null, ex.Message));
                _logger.Error($"Failed to create invoice for {row.Client.ClientName} ({row.Client.ClientId}): {ex.Message}");
            }
        }

        try
        {
            repo.CommitApprovals(approvals, _config.RenewalMonths);
        }
        catch (Exception ex)
        {
            _logger.Error($"Failed to write approved invoices back to the workbook: {ex.Message}");
        }

        var skippedCount = totalCandidateCount - approvedRows.Count;
        BuildCompleted?.Invoke(new BuildCompleteEventArgs(created, failed, skippedCount, _config.InvoiceFolder));
    }
}
