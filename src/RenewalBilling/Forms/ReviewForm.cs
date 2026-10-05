using System.Diagnostics;
using System.Globalization;
using RenewalBilling.Models;
using RenewalBilling.Services;

namespace RenewalBilling.Forms;

/// <summary>
/// Master/detail review form: lists every eligible client, lets the user adjust the RPI
/// percentage or individual proposed prices, approve or skip each one, and build the approved
/// invoices. All Excel and Word automation happens on <see cref="ReviewRunWorker"/>'s dedicated
/// thread; this form only ever touches its own controls, and marshals back to its own UI thread
/// whenever the worker raises a progress or completion event from that other thread.
/// </summary>
public partial class ReviewForm : Form
{
    private static readonly CultureInfo EnGb = CultureInfo.GetCultureInfo("en-GB");

    private readonly ReviewRunWorker _worker;
    private readonly AppConfig _config;
    private readonly IReadOnlyList<CandidateRow> _candidates;
    private readonly Dictionary<DataGridViewRow, CandidateRow> _gridRowToCandidate = new();
    private CandidateRow? _selectedCandidate;
    private bool _buildRequested;

    /// <summary>
    /// Creates the review form.
    /// </summary>
    /// <param name="runDate">The date this run is treating as "today".</param>
    /// <param name="rpi">The RPI result fetched for this run.</param>
    /// <param name="candidates">Every eligible client found.</param>
    /// <param name="worker">The worker that owns Excel and Word for this run.</param>
    /// <param name="config">The loaded application config.</param>
    public ReviewForm(DateTime runDate, RpiResult rpi, IReadOnlyList<CandidateRow> candidates, ReviewRunWorker worker, AppConfig config)
    {
        _worker = worker;
        _config = config;
        _candidates = candidates;

        InitializeComponent();

        _runDateLabel.Text = $"Run date: {runDate.ToString("dd MMM yyyy", EnGb)}";
        _rpiNumericUpDown.Value = Math.Clamp(rpi.Percent, _rpiNumericUpDown.Minimum, _rpiNumericUpDown.Maximum);
        _rpiPeriodLabel.Text = $"({rpi.PeriodLabel})";
        _rpiSourceLabel.Text = $"Source: {rpi.Source}";

        if (rpi.Source != RpiSource.Online)
        {
            _warningBanner.Text = rpi.Source == RpiSource.Cache
                ? "Warning: ONS could not be reached. Using the last cached RPI value."
                : "Warning: ONS could not be reached and no cache was available. Using the configured fallback RPI percentage.";
            _warningBanner.Visible = true;
        }

        PopulateCandidatesGrid();

        _candidatesGrid.SelectionChanged += CandidatesGrid_SelectionChanged;
        _candidatesGrid.CurrentCellDirtyStateChanged += CandidatesGrid_CurrentCellDirtyStateChanged;
        _candidatesGrid.CellValueChanged += CandidatesGrid_CellValueChanged;
        _itemsGrid.CellEndEdit += ItemsGrid_CellEndEdit;
        _rpiNumericUpDown.ValueChanged += RpiNumericUpDown_ValueChanged;
        _createButton.Click += CreateButton_Click;
        _closeButton.Click += (_, _) => Close();
        _openFolderButton.Click += (_, _) => OpenInvoiceFolder();
        FormClosing += ReviewForm_FormClosing;

        _worker.ProgressChanged += OnWorkerProgressChanged;
        _worker.BuildCompleted += OnWorkerBuildCompleted;

        if (_candidatesGrid.Rows.Count > 0)
        {
            _candidatesGrid.Rows[0].Selected = true;
        }
    }

    private void PopulateCandidatesGrid()
    {
        foreach (var candidate in _candidates)
        {
            var rowIndex = _candidatesGrid.Rows.Add(
                candidate.Approved,
                candidate.Client.ClientName,
                candidate.PeriodStart.ToString("dd MMM yyyy", EnGb),
                candidate.Status.ToString(),
                FormatMoney(candidate.Net),
                FormatMoney(candidate.Total));

            var row = _candidatesGrid.Rows[rowIndex];
            if (candidate.Status == CandidateStatus.Overdue)
            {
                row.DefaultCellStyle.ForeColor = Color.DarkRed;
            }

            _gridRowToCandidate[row] = candidate;
        }
    }

    private void CandidatesGrid_SelectionChanged(object? sender, EventArgs e)
    {
        if (_candidatesGrid.SelectedRows.Count == 0)
        {
            _selectedCandidate = null;
            return;
        }

        var row = _candidatesGrid.SelectedRows[0];
        if (!_gridRowToCandidate.TryGetValue(row, out var candidate))
        {
            return;
        }

        _selectedCandidate = candidate;
        LoadDetail(candidate);
    }

    private void LoadDetail(CandidateRow candidate)
    {
        var client = candidate.Client;
        _clientNameLabel.Text = client.ClientName;
        _contactNameLabel.Text = client.ContactName;
        _address1Label.Text = client.Address1;
        _address2Label.Text = client.Address2;
        _townPostcodeLabel.Text = $"{client.Town} {client.Postcode}".Trim();
        _emailLabel.Text = client.Email;

        _itemsGrid.CellEndEdit -= ItemsGrid_CellEndEdit;
        _itemsGrid.Rows.Clear();
        foreach (var item in candidate.ProposedItems)
        {
            _itemsGrid.Rows.Add(
                item.Item.Description,
                item.Item.Quantity.ToString("0.##", EnGb),
                FormatMoney(item.Item.UnitCost),
                FormatMoney(item.ProposedUnitCost),
                item.UpliftPercentApplied.ToString("0.0", EnGb));
        }

        _itemsGrid.CellEndEdit += ItemsGrid_CellEndEdit;

        RefreshTotals(candidate);
    }

    private void RefreshTotals(CandidateRow candidate)
    {
        _subtotalLabel.Text = FormatMoney(candidate.Net);
        _vatLabel.Text = FormatMoney(candidate.Vat);
        _totalLabel.Text = FormatMoney(candidate.Total);
    }

    private void ItemsGrid_CellEndEdit(object? sender, DataGridViewCellEventArgs e)
    {
        if (_selectedCandidate is null || e.ColumnIndex != _proposedPriceColumn.Index)
        {
            return;
        }

        var cellText = _itemsGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? string.Empty;
        if (!decimal.TryParse(cellText, NumberStyles.Currency | NumberStyles.AllowDecimalPoint, EnGb, out var newPrice))
        {
            // Not a usable number; put the old value back rather than silently losing the row.
            var old = _selectedCandidate.ProposedItems[e.RowIndex];
            _itemsGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = FormatMoney(old.ProposedUnitCost);
            return;
        }

        var current = _selectedCandidate.ProposedItems[e.RowIndex];
        var newUpliftPercent = current.Item.UnitCost == 0
            ? 0m
            : Math.Round((newPrice / current.Item.UnitCost - 1) * 100m, 1, MidpointRounding.AwayFromZero);

        _selectedCandidate.ProposedItems[e.RowIndex] = new ProposedLineItem(current.Item, newPrice, newUpliftPercent);

        _itemsGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = FormatMoney(newPrice);
        _itemsGrid.Rows[e.RowIndex].Cells[_upliftColumn.Index].Value = newUpliftPercent.ToString("0.0", EnGb);

        RefreshTotals(_selectedCandidate);
        RefreshCandidateGridRow(_selectedCandidate);
    }

    private void RpiNumericUpDown_ValueChanged(object? sender, EventArgs e)
    {
        var newRpiPercent = _rpiNumericUpDown.Value;

        foreach (var candidate in _candidates)
        {
            var proposed = UpliftCalculator.BuildProposedItems(
                candidate.OriginalItems, newRpiPercent, _config.Uplift.RoundTo, _config.Uplift.MinimumPercent, _config.Uplift.MaximumPercent);
            candidate.ProposedItems = proposed.ToList();
            RefreshCandidateGridRow(candidate);
        }

        if (_selectedCandidate is not null)
        {
            LoadDetail(_selectedCandidate);
        }
    }

    private void RefreshCandidateGridRow(CandidateRow candidate)
    {
        foreach (DataGridViewRow row in _candidatesGrid.Rows)
        {
            if (_gridRowToCandidate.TryGetValue(row, out var rowCandidate) && rowCandidate == candidate)
            {
                row.Cells[_netColumn.Index].Value = FormatMoney(candidate.Net);
                row.Cells[_totalColumn.Index].Value = FormatMoney(candidate.Total);
                return;
            }
        }
    }

    private void CandidatesGrid_CurrentCellDirtyStateChanged(object? sender, EventArgs e)
    {
        // The checkbox column's value is normally only committed when focus leaves the cell.
        // Committing immediately means ticking a box takes effect straight away.
        if (_candidatesGrid.IsCurrentCellDirty && _candidatesGrid.CurrentCell?.ColumnIndex == _approveColumn.Index)
        {
            _candidatesGrid.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }
    }

    private void CandidatesGrid_CellValueChanged(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0 || e.ColumnIndex != _approveColumn.Index)
        {
            return;
        }

        var row = _candidatesGrid.Rows[e.RowIndex];
        if (_gridRowToCandidate.TryGetValue(row, out var candidate))
        {
            candidate.Approved = row.Cells[_approveColumn.Index].Value is true;
        }
    }

    private void CreateButton_Click(object? sender, EventArgs e)
    {
        var approvedRows = _candidates.Where(c => c.Approved).ToList();
        if (approvedRows.Count == 0)
        {
            MessageBox.Show(this, "No renewals are approved. Tick at least one before creating invoices.",
                "Nothing to create", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        _buildRequested = true;
        SetFormEnabled(false);
        _buildProgressBar.Minimum = 0;
        _buildProgressBar.Maximum = approvedRows.Count;
        _buildProgressBar.Value = 0;
        _statusLabel.Text = $"Creating 0 of {approvedRows.Count}...";

        _worker.RequestBuild(approvedRows);
    }

    private void SetFormEnabled(bool enabled)
    {
        _candidatesGrid.Enabled = enabled;
        _itemsGrid.Enabled = enabled;
        _rpiNumericUpDown.Enabled = enabled;
        _createButton.Enabled = enabled;
    }

    private void OnWorkerProgressChanged(BuildProgressEventArgs progress)
    {
        if (InvokeRequired)
        {
            BeginInvoke(() => OnWorkerProgressChanged(progress));
            return;
        }

        _buildProgressBar.Value = Math.Min(progress.Completed, _buildProgressBar.Maximum);
        _statusLabel.Text = $"Creating {progress.Completed} of {progress.Total}: {progress.ClientName}...";
    }

    private void OnWorkerBuildCompleted(BuildCompleteEventArgs result)
    {
        if (InvokeRequired)
        {
            BeginInvoke(() => OnWorkerBuildCompleted(result));
            return;
        }

        _statusLabel.Text = $"Done: {result.Created.Count} created, {result.Failed.Count} failed, {result.SkippedCount} skipped.";

        var lines = new List<string>();
        lines.Add($"Created ({result.Created.Count}):");
        lines.AddRange(result.Created.Select(c => $"  {c.ClientName} - {c.InvoiceNumber}"));
        lines.Add($"Failed ({result.Failed.Count}):");
        lines.AddRange(result.Failed.Select(f => $"  {f.ClientName} - {f.Reason}"));
        lines.Add($"Skipped (not approved): {result.SkippedCount}");

        _summaryTextBox.Text = string.Join(Environment.NewLine, lines);
        _summaryTextBox.Visible = true;
        _buildProgressBar.Visible = false;
        _statusLabel.Visible = false;

        _openFolderButton.Visible = true;
        _closeButton.Enabled = true;
        _createButton.Visible = false;
    }

    private void OpenInvoiceFolder()
    {
        if (!Directory.Exists(_config.InvoiceFolder))
        {
            Directory.CreateDirectory(_config.InvoiceFolder);
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = _config.InvoiceFolder,
            UseShellExecute = true,
        });
    }

    private void ReviewForm_FormClosing(object? sender, FormClosingEventArgs e)
    {
        if (!_buildRequested)
        {
            _worker.RequestClose();
        }
    }

    private static string FormatMoney(decimal value) => value.ToString("C2", EnGb);
}
