using System.Globalization;
using System.Runtime.InteropServices;
using RenewalBilling.Models;

namespace RenewalBilling.Services;

/// <summary>
/// Thrown when the client workbook cannot be read or written: it is locked, missing, or missing
/// one of the configured worksheets or column headers.
/// </summary>
public sealed class WorkbookException : Exception
{
    /// <summary>Creates a new <see cref="WorkbookException"/> with a message describing the problem.</summary>
    /// <param name="message">A clear, human readable description of what is wrong.</param>
    public WorkbookException(string message) : base(message)
    {
    }
}

/// <summary>
/// The three worksheets read from the client workbook at the start of a run.
/// </summary>
public sealed class WorkbookSnapshot
{
    /// <summary>Every row from the Clients worksheet.</summary>
    public required IReadOnlyList<Client> Clients { get; init; }

    /// <summary>Every row from the Items worksheet.</summary>
    public required IReadOnlyList<BillingItem> Items { get; init; }

    /// <summary>Every row from the Invoices worksheet.</summary>
    public required IReadOnlyList<InvoiceRecord> Invoices { get; init; }
}

/// <summary>
/// One client's approved renewal, ready to be written back to the workbook: the candidate with
/// its final invoice number, together with the paths of the invoice files already saved by the
/// Word invoice builder.
/// </summary>
/// <param name="Candidate">The approved candidate, including its assigned invoice number.</param>
/// <param name="DocxPath">Full path to the saved .docx copy of the invoice.</param>
/// <param name="PdfPath">Full path to the saved PDF copy of the invoice.</param>
public sealed record ApprovedRenewal(InvoiceCandidate Candidate, string DocxPath, string PdfPath);

/// <summary>
/// Reads and writes the client workbook.
/// </summary>
public interface IClientRepository : IDisposable
{
    /// <summary>
    /// Checks the workbook is not locked by another process, opens it, and reads all three
    /// worksheets in one pass each.
    /// </summary>
    /// <returns>Every row currently in the workbook.</returns>
    WorkbookSnapshot Open();

    /// <summary>
    /// Backs up the workbook, then writes every approval back: a new Invoices row, updated item
    /// prices, and an advanced due date for each client, before saving once.
    /// </summary>
    /// <param name="approvals">The renewals the user approved this run.</param>
    /// <param name="renewalMonths">How many months to advance each approved client's due date by.</param>
    void CommitApprovals(IReadOnlyList<ApprovedRenewal> approvals, int renewalMonths);
}

/// <summary>
/// Automates a private Excel instance to read and write the client workbook described in section
/// 6 of the specification. The row parsing methods below take plain arrays, not COM objects, so
/// they are pure and can be unit tested without Excel installed.
///
/// Excel itself is automated through late-bound <c>dynamic</c> COM calls against a ProgID
/// (<c>Excel.Application</c>) rather than the typed <c>Microsoft.Office.Interop.Excel</c> PIA.
/// That typed assembly carries a hard dependency on the old "office.dll" core PIA, which a
/// Click-to-Run Office install does not register, and which fails to load at a seemingly
/// unrelated call site. Late binding talks to Excel purely through its registered ProgID and
/// avoids that dependency entirely, matching the fallback the specification calls for.
/// </summary>
public sealed class ExcelRepository : IClientRepository
{
    private readonly string _workbookPath;
    private readonly string _backupFolder;
    private readonly SheetsConfig _sheets;
    private readonly ColumnsConfig _columns;
    private readonly IAppLogger _logger;

    private readonly Dictionary<string, int> _clientRows = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<BillingItem, int> _itemRows = new();
    private Dictionary<string, int> _clientHeaderMap = new();
    private Dictionary<string, int> _itemHeaderMap = new();
    private Dictionary<string, int> _invoiceHeaderMap = new();
    private int _invoicesNextRow;

    private dynamic? _app;
    private int? _appProcessId;
    private dynamic? _workbook;
    private dynamic? _clientsSheet;
    private dynamic? _itemsSheet;
    private dynamic? _invoicesSheet;
    private bool _opened;

    // Excel's XlFileFormat.xlOpenXMLWorkbook enum value (.xlsx without macros). Used as a plain
    // int instead of the typed enum so this class never has to reference the Office Interop PIA.
    private const int XlOpenXmlWorkbook = 51;

    /// <summary>
    /// Creates a repository bound to one workbook.
    /// </summary>
    /// <param name="workbookPath">Full path to the client workbook.</param>
    /// <param name="backupFolder">Full path to the folder backups are copied into.</param>
    /// <param name="sheets">The configured worksheet names.</param>
    /// <param name="columns">The configured column headers for every worksheet.</param>
    /// <param name="logger">Where to log warnings raised while reading or writing.</param>
    public ExcelRepository(string workbookPath, string backupFolder, SheetsConfig sheets, ColumnsConfig columns, IAppLogger logger)
    {
        _workbookPath = workbookPath;
        _backupFolder = backupFolder;
        _sheets = sheets;
        _columns = columns;
        _logger = logger;
    }

    /// <inheritdoc />
    public WorkbookSnapshot Open()
    {
        EnsureNotLocked();

        (_app, _appProcessId) = ComRelease.StartOfficeApplication("Excel.Application", "EXCEL");
        _app.Visible = false;
        _app.DisplayAlerts = false;
        _app.ScreenUpdating = false;

        dynamic workbooks = _app.Workbooks;
        try
        {
            _workbook = workbooks.Open(_workbookPath, 0, false); // Filename, UpdateLinks, ReadOnly
        }
        finally
        {
            Marshal.FinalReleaseComObject(workbooks);
        }

        _clientsSheet = GetSheet(_workbook, _sheets.Clients);
        _itemsSheet = GetSheet(_workbook, _sheets.Items);
        _invoicesSheet = GetSheet(_workbook, _sheets.Invoices);

        var clients = ReadClients(_clientsSheet);
        var items = ReadItems(_itemsSheet);
        var invoices = ReadInvoices(_invoicesSheet);

        _opened = true;
        return new WorkbookSnapshot { Clients = clients, Items = items, Invoices = invoices };
    }

    /// <inheritdoc />
    public void CommitApprovals(IReadOnlyList<ApprovedRenewal> approvals, int renewalMonths)
    {
        if (!_opened)
        {
            throw new InvalidOperationException("Open() must be called before CommitApprovals().");
        }

        if (approvals.Count == 0)
        {
            return;
        }

        BackupWorkbook();
        WriteInvoiceRows(approvals);
        WriteItemUpdates(approvals);
        WriteClientDueDates(approvals, renewalMonths);

        _workbook!.Save();
    }

    /// <summary>
    /// Creates a sample workbook at the configured path, with a handful of clients chosen to
    /// exercise every eligibility case: due soon, overdue, already invoiced, not due yet, a
    /// blank due date, an inactive client, and a client with no billing items.
    /// </summary>
    /// <param name="runDate">The date sample due dates are calculated relative to.</param>
    public void CreateSampleWorkbook(DateTime runDate)
    {
        var folder = Path.GetDirectoryName(_workbookPath);
        if (!string.IsNullOrEmpty(folder))
        {
            Directory.CreateDirectory(folder);
        }

        var (app, appProcessId) = ComRelease.StartOfficeApplication("Excel.Application", "EXCEL");
        app.Visible = false;
        app.DisplayAlerts = false;
        dynamic? workbook = null;
        dynamic workbooks = app.Workbooks;

        try
        {
            workbook = workbooks.Add();
            EnsureThreeNamedSheets(workbook, _sheets);

            WriteToSheetAndRelease((object)workbook, _sheets.Clients, sheet => WriteSampleClients(sheet, runDate));
            WriteToSheetAndRelease((object)workbook, _sheets.Items, sheet => WriteSampleItems(sheet));
            WriteToSheetAndRelease((object)workbook, _sheets.Invoices, sheet => WriteSampleInvoices(sheet, runDate));

            if (File.Exists(_workbookPath))
            {
                File.Delete(_workbookPath);
            }

            workbook.SaveAs(_workbookPath, XlOpenXmlWorkbook);
        }
        finally
        {
            try
            {
                workbook?.Close(false);
            }
            catch (COMException)
            {
                // Best effort; we still need to quit the application and release every COM
                // object below.
            }

            // Release the workbook and force a full GC pass before quitting. See the matching
            // comment in Dispose() for why: untracked COM wrappers from property chains elsewhere
            // in this method must be finalized before Quit(), or Excel stays resident.
            if (workbook is not null)
            {
                Marshal.FinalReleaseComObject(workbook);
            }

            Marshal.FinalReleaseComObject(workbooks);
            ComRelease.CollectTwice();

            ComRelease.QuitAndEnsureProcessExits(app, appProcessId);
        }
    }

    /// <inheritdoc />
    public void Dispose()
    {
        try
        {
            _workbook?.Close(false);
        }
        catch (COMException)
        {
            // Best effort. Even if the workbook will not close cleanly we still need to quit
            // the application and release every COM object below.
        }

        // Release every object we tracked, then force a full GC pass before quitting. Excel
        // property chains (sheet.Cells[...], workbook.Worksheets[...], and so on) each return a
        // fresh, untracked COM wrapper; if any of that garbage has not yet been finalized when
        // Quit() runs, Excel sees an outstanding reference and stays resident as an orphan
        // process instead of actually exiting.
        ComRelease.ReleaseAll(_workbook, _clientsSheet, _itemsSheet, _invoicesSheet);

        if (_app is not null)
        {
            ComRelease.QuitAndEnsureProcessExits(_app, _appProcessId);
        }

        _app = null;
        _appProcessId = null;
        _workbook = null;
        _clientsSheet = null;
        _itemsSheet = null;
        _invoicesSheet = null;
    }

    private void EnsureNotLocked()
    {
        try
        {
            using var stream = new FileStream(_workbookPath, FileMode.Open, FileAccess.Read, FileShare.None);
        }
        catch (IOException)
        {
            throw new WorkbookException("Workbook is open elsewhere, close it and run again.");
        }
    }

    private static dynamic GetSheet(dynamic workbook, string sheetName)
    {
        dynamic worksheets = workbook.Worksheets;
        try
        {
            return worksheets[sheetName];
        }
        catch (COMException)
        {
            throw new WorkbookException($"Workbook is missing the '{sheetName}' worksheet.");
        }
        finally
        {
            Marshal.FinalReleaseComObject(worksheets);
        }
    }

    /// <summary>
    /// Fetches a worksheet by name, runs <paramref name="write"/> against it, and always
    /// releases the worksheet COM reference afterwards. <see cref="GetSheet"/> returns a fresh
    /// RCW on every call, so a sheet fetched only to be used once must still be explicitly
    /// released, or Excel never actually exits when <c>Quit()</c> is called later.
    /// </summary>
    private static void WriteToSheetAndRelease(dynamic workbook, string sheetName, Action<dynamic> write)
    {
        dynamic sheet = GetSheet(workbook, sheetName);
        try
        {
            write(sheet);
        }
        finally
        {
            Marshal.FinalReleaseComObject(sheet);
        }
    }

    private List<Client> ReadClients(dynamic sheet)
    {
        object[,] data = ReadRawData(sheet);
        _clientHeaderMap = BuildHeaderMap(data);
        ValidateHeaders(_clientHeaderMap, AllClientHeaders(_columns.Clients), _sheets.Clients);

        var parsed = ParseClients(data, _clientHeaderMap, _columns.Clients);
        var clients = new List<Client>(parsed.Count);
        foreach (var (client, row) in parsed)
        {
            clients.Add(client);
            _clientRows[client.ClientId] = row;
        }

        return clients;
    }

    private List<BillingItem> ReadItems(dynamic sheet)
    {
        object[,] data = ReadRawData(sheet);
        _itemHeaderMap = BuildHeaderMap(data);
        ValidateHeaders(_itemHeaderMap, AllItemHeaders(_columns.Items), _sheets.Items);

        var parsed = ParseItems(data, _itemHeaderMap, _columns.Items);
        var items = new List<BillingItem>(parsed.Count);
        foreach (var (item, row) in parsed)
        {
            items.Add(item);
            _itemRows[item] = row;
        }

        return items;
    }

    private List<InvoiceRecord> ReadInvoices(dynamic sheet)
    {
        object[,] data = ReadRawData(sheet);
        _invoiceHeaderMap = BuildHeaderMap(data);
        ValidateHeaders(_invoiceHeaderMap, AllInvoiceHeaders(_columns.Invoices), _sheets.Invoices);

        _invoicesNextRow = data.GetLength(0) + 1;
        return ParseInvoices(data, _invoiceHeaderMap, _columns.Invoices).ToList();
    }

    private static object[,] ReadRawData(dynamic sheet)
    {
        dynamic usedRange = sheet.UsedRange;
        try
        {
            return NormalizeToArray(usedRange.Value2);
        }
        finally
        {
            Marshal.FinalReleaseComObject(usedRange);
        }
    }

    private static object[,] NormalizeToArray(object? value)
    {
        if (value is object[,] array)
        {
            return array;
        }

        var single = (object[,])Array.CreateInstance(typeof(object), new[] { 1, 1 }, new[] { 1, 1 });
        single[1, 1] = value ?? string.Empty;
        return single;
    }

    private void BackupWorkbook()
    {
        Directory.CreateDirectory(_backupFolder);
        var backupPath = Path.Combine(_backupFolder, $"Clients_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
        File.Copy(_workbookPath, backupPath);

        var backups = Directory.GetFiles(_backupFolder, "Clients_*.xlsx")
            .OrderByDescending(path => path, StringComparer.OrdinalIgnoreCase)
            .Skip(12);

        foreach (var oldBackup in backups)
        {
            File.Delete(oldBackup);
        }
    }

    private void WriteInvoiceRows(IReadOnlyList<ApprovedRenewal> approvals)
    {
        var columnCount = _invoiceHeaderMap.Values.Max();
        var array = new object[approvals.Count, columnCount];
        var invoices = _columns.Invoices;

        for (var i = 0; i < approvals.Count; i++)
        {
            var candidate = approvals[i].Candidate;

            SetCell(array, i, _invoiceHeaderMap[invoices.InvoiceNumber], candidate.InvoiceNumber ?? string.Empty);
            SetCell(array, i, _invoiceHeaderMap[invoices.ClientId], candidate.Client.ClientId);
            SetCell(array, i, _invoiceHeaderMap[invoices.InvoiceDate], candidate.InvoiceDate.ToOADate());
            SetCell(array, i, _invoiceHeaderMap[invoices.DueDate], candidate.PayByDate.ToOADate());
            SetCell(array, i, _invoiceHeaderMap[invoices.PeriodStart], candidate.PeriodStart.ToOADate());
            SetCell(array, i, _invoiceHeaderMap[invoices.PeriodEnd], candidate.PeriodEnd.ToOADate());
            SetCell(array, i, _invoiceHeaderMap[invoices.Net], (double)candidate.Net);
            SetCell(array, i, _invoiceHeaderMap[invoices.Vat], (double)candidate.Vat);
            SetCell(array, i, _invoiceHeaderMap[invoices.Total], (double)candidate.Total);
            SetCell(array, i, _invoiceHeaderMap[invoices.RpiPercent], (double)candidate.RpiPercentUsed);
            SetCell(array, i, _invoiceHeaderMap[invoices.DocxPath], approvals[i].DocxPath);
            SetCell(array, i, _invoiceHeaderMap[invoices.PdfPath], approvals[i].PdfPath);
        }

        var startRow = _invoicesNextRow;
        var endRow = startRow + approvals.Count - 1;
        var address = $"{ColumnLetter(1)}{startRow}:{ColumnLetter(columnCount)}{endRow}";
        dynamic range = _invoicesSheet!.Range[address];
        try
        {
            range.Value2 = array;
        }
        finally
        {
            Marshal.FinalReleaseComObject(range);
        }

        _invoicesNextRow = endRow + 1;
    }

    private void WriteItemUpdates(IReadOnlyList<ApprovedRenewal> approvals)
    {
        var previousCol = _itemHeaderMap[_columns.Items.PreviousUnitCost];
        var unitCostCol = _itemHeaderMap[_columns.Items.UnitCost];
        var upliftCol = _itemHeaderMap[_columns.Items.LastUpliftPercent];

        foreach (var approval in approvals)
        {
            foreach (var line in approval.Candidate.Items)
            {
                if (!_itemRows.TryGetValue(line.Item, out var row))
                {
                    _logger.Warn($"Could not find the worksheet row for item '{line.Item.Description}' (client {approval.Candidate.Client.ClientId}). Price was not written back.");
                    continue;
                }

                WriteCell(_itemsSheet!, row, previousCol, (double)line.Item.UnitCost);
                WriteCell(_itemsSheet!, row, unitCostCol, (double)line.ProposedUnitCost);
                WriteCell(_itemsSheet!, row, upliftCol, (double)line.UpliftPercentApplied);
            }
        }
    }

    private void WriteClientDueDates(IReadOnlyList<ApprovedRenewal> approvals, int renewalMonths)
    {
        var dueDateCol = _clientHeaderMap[_columns.Clients.DueDate];

        foreach (var approval in approvals)
        {
            var clientId = approval.Candidate.Client.ClientId;
            if (!_clientRows.TryGetValue(clientId, out var row))
            {
                _logger.Warn($"Could not find the worksheet row for client '{clientId}'. Due date was not advanced.");
                continue;
            }

            var nextDueDate = EligibilityService.ComputeNextDueDate(approval.Candidate.PeriodStart, renewalMonths);
            WriteCell(_clientsSheet!, row, dueDateCol, nextDueDate.ToOADate());
        }
    }

    private static void SetCell(object[,] array, int rowIndex, int excelColumn, object value) => array[rowIndex, excelColumn - 1] = value;

    /// <summary>
    /// Converts a 1-based column number to its Excel letter, for example 1 -&gt; "A", 27 -&gt; "AA".
    /// Used to build A1-style range addresses instead of <c>Worksheet.Cells[row, col]</c>, which
    /// requires fetching an extra whole-sheet "Cells" collection object for every call. A1
    /// addresses need only the single target <c>Range</c>, so there is nothing extra to release.
    /// </summary>
    private static string ColumnLetter(int column)
    {
        var letters = string.Empty;
        while (column > 0)
        {
            var remainder = (column - 1) % 26;
            letters = (char)('A' + remainder) + letters;
            column = (column - 1) / 26;
        }

        return letters;
    }

    private static void WriteCell(dynamic sheet, int row, int column, object value)
    {
        dynamic cell = sheet.Range[$"{ColumnLetter(column)}{row}"];
        try
        {
            cell.Value2 = value;
        }
        finally
        {
            Marshal.FinalReleaseComObject(cell);
        }
    }

    private static void EnsureThreeNamedSheets(dynamic workbook, SheetsConfig sheets)
    {
        dynamic worksheets = workbook.Worksheets;
        try
        {
            while ((int)worksheets.Count > 1)
            {
                dynamic extra = worksheets[(int)worksheets.Count];
                extra.Delete();
                Marshal.FinalReleaseComObject(extra);
            }

            dynamic first = worksheets[1];
            first.Name = sheets.Clients;

            dynamic itemsSheet = worksheets.Add(After: first);
            itemsSheet.Name = sheets.Items;

            dynamic invoicesSheet = worksheets.Add(After: itemsSheet);
            invoicesSheet.Name = sheets.Invoices;

            Marshal.FinalReleaseComObject(first);
            Marshal.FinalReleaseComObject(itemsSheet);
            Marshal.FinalReleaseComObject(invoicesSheet);
        }
        finally
        {
            Marshal.FinalReleaseComObject(worksheets);
        }
    }

    private void WriteSampleClients(dynamic sheet, DateTime runDate)
    {
        var headers = AllClientHeaders(_columns.Clients);
        var rows = new List<object[]>
        {
            new object[] { "C1", "Acme Ltd", "Jo Bloggs", "1 High Street", "", "Anytown", "AB1 2CD", "jo@acme.test", runDate.AddDays(10).ToOADate(), "", "Y" },
            new object[] { "C2", "Beta Services", "Pat Smith", "2 Market Road", "", "Someville", "CD3 4EF", "pat@beta.test", runDate.AddDays(-5).ToOADate(), "", "Y" },
            new object[] { "C3", "Gamma Co", "Sam Lee", "3 Church Lane", "Unit 5", "Oldtown", "EF5 6GH", "sam@gamma.test", runDate.AddDays(10).ToOADate(), runDate.AddYears(-1).ToOADate(), "Y" },
            new object[] { "C4", "Delta Partnership", "Alex Grey", "4 Park Avenue", "", "Newtown", "GH7 8IJ", "alex@delta.test", runDate.AddDays(60).ToOADate(), "", "Y" },
            new object[] { "C5", "Epsilon Group", "Robin Day", "5 Mill Lane", "", "Middleton", "IJ9 0KL", "robin@epsilon.test", "", "", "Y" },
            new object[] { "C6", "Zeta Retail", "Casey Fox", "6 Station Road", "", "Eastville", "KL1 2MN", "casey@zeta.test", runDate.AddDays(5).ToOADate(), "", "N" },
            new object[] { "C7", "Eta Consulting", "Drew Hale", "7 Bridge Street", "", "Westfield", "MN3 4OP", "drew@eta.test", runDate.AddDays(5).ToOADate(), "", "Y" },
        };

        WriteHeaderAndRows(sheet, headers, rows);
        FormatColumnAsDate(sheet, Array.IndexOf(headers, _columns.Clients.DueDate) + 1, rows.Count);
        FormatColumnAsDate(sheet, Array.IndexOf(headers, _columns.Clients.LastPaidDate) + 1, rows.Count);
    }

    private void WriteSampleItems(dynamic sheet)
    {
        var headers = AllItemHeaders(_columns.Items);
        var rows = new List<object[]>
        {
            new object[] { "C1", "Annual software licence", 1d, 250.00, 240.00, 4.2 },
            new object[] { "C2", "Support contract", 1d, 500.00, 480.00, 4.2 },
            new object[] { "C3", "Hosting", 12d, 15.00, 14.50, 3.5 },
            new object[] { "C3", "Domain renewal", 1d, 12.00, 11.50, 3.5 },
            new object[] { "C4", "Membership fee", 1d, 99.00, "", "" },
            new object[] { "C5", "Consulting retainer", 1d, 1200.00, "", "" },
            new object[] { "C6", "Retail licence", 3d, 75.00, "", "" },
        };

        WriteHeaderAndRows(sheet, headers, rows);
    }

    private void WriteSampleInvoices(dynamic sheet, DateTime runDate)
    {
        var headers = AllInvoiceHeaders(_columns.Invoices);

        var gammaPeriodStart = runDate.AddDays(10);
        var gammaPeriodEnd = EligibilityService.ComputePeriodEnd(gammaPeriodStart, 12);
        var betaOldPeriodStart = runDate.AddDays(-5).AddYears(-1);
        var betaOldPeriodEnd = EligibilityService.ComputePeriodEnd(betaOldPeriodStart, 12);

        var rows = new List<object[]>
        {
            new object[]
            {
                "INV-2026-0001", "C3", runDate.AddDays(-20).ToOADate(), gammaPeriodStart.ToOADate(),
                gammaPeriodStart.ToOADate(), gammaPeriodEnd.ToOADate(), 324.00, 0.00, 324.00, 3.5,
                @"Invoices\2026\INV-2026-0001 Gamma Co.docx", @"Invoices\2026\INV-2026-0001 Gamma Co.pdf",
            },
            new object[]
            {
                "INV-2025-0099", "C2", betaOldPeriodStart.ToOADate(), betaOldPeriodStart.ToOADate(),
                betaOldPeriodStart.ToOADate(), betaOldPeriodEnd.ToOADate(), 480.00, 0.00, 480.00, 2.1,
                @"Invoices\2025\INV-2025-0099 Beta Services.docx", @"Invoices\2025\INV-2025-0099 Beta Services.pdf",
            },
        };

        WriteHeaderAndRows(sheet, headers, rows);
        FormatColumnAsDate(sheet, Array.IndexOf(headers, _columns.Invoices.InvoiceDate) + 1, rows.Count);
        FormatColumnAsDate(sheet, Array.IndexOf(headers, _columns.Invoices.DueDate) + 1, rows.Count);
        FormatColumnAsDate(sheet, Array.IndexOf(headers, _columns.Invoices.PeriodStart) + 1, rows.Count);
        FormatColumnAsDate(sheet, Array.IndexOf(headers, _columns.Invoices.PeriodEnd) + 1, rows.Count);
    }

    private static void WriteHeaderAndRows(dynamic sheet, string[] headers, List<object[]> rows)
    {
        var totalRows = rows.Count + 1;
        var totalCols = headers.Length;
        var array = new object[totalRows, totalCols];

        for (var col = 0; col < totalCols; col++)
        {
            array[0, col] = headers[col];
        }

        for (var row = 0; row < rows.Count; row++)
        {
            for (var col = 0; col < totalCols; col++)
            {
                array[row + 1, col] = rows[row][col];
            }
        }

        dynamic range = sheet.Range[$"A1:{ColumnLetter(totalCols)}{totalRows}"];
        try
        {
            range.Value2 = array;
        }
        finally
        {
            Marshal.FinalReleaseComObject(range);
        }
    }

    private static void FormatColumnAsDate(dynamic sheet, int column, int dataRowCount)
    {
        if (dataRowCount == 0 || column <= 0)
        {
            return;
        }

        var letter = ColumnLetter(column);
        dynamic range = sheet.Range[$"{letter}2:{letter}{dataRowCount + 1}"];
        try
        {
            range.NumberFormat = "dd mmm yyyy";
        }
        finally
        {
            Marshal.FinalReleaseComObject(range);
        }
    }

    /// <summary>
    /// Builds a map of header name to 1-based column number from a sheet's row 1.
    /// </summary>
    /// <param name="data">The raw sheet data, as returned by a used range's <c>Value2</c>.</param>
    /// <returns>Header name to column number, matched case-insensitively.</returns>
    public static Dictionary<string, int> BuildHeaderMap(object[,] data)
    {
        var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var columnCount = data.GetLength(1);

        for (var col = 1; col <= columnCount; col++)
        {
            var header = GetString(data, 1, col);
            if (!string.IsNullOrWhiteSpace(header) && !map.ContainsKey(header))
            {
                map[header] = col;
            }
        }

        return map;
    }

    /// <summary>
    /// Checks that every required header is present in a header map, raising a clear error
    /// naming every header that is missing.
    /// </summary>
    /// <param name="headerMap">The header map built from the sheet.</param>
    /// <param name="requiredHeaders">The header names the configuration expects to find.</param>
    /// <param name="sheetName">The worksheet name, used in the error message.</param>
    /// <exception cref="WorkbookException">Thrown when one or more headers are missing.</exception>
    public static void ValidateHeaders(IReadOnlyDictionary<string, int> headerMap, IEnumerable<string> requiredHeaders, string sheetName)
    {
        var missing = requiredHeaders.Where(header => !headerMap.ContainsKey(header)).Distinct().ToList();
        if (missing.Count > 0)
        {
            throw new WorkbookException($"Sheet '{sheetName}' is missing required column(s): {string.Join(", ", missing)}.");
        }
    }

    /// <summary>
    /// Parses every data row of the Clients worksheet into a <see cref="Client"/>, alongside the
    /// 1-based worksheet row it came from.
    /// </summary>
    /// <param name="data">The raw sheet data, as returned by a used range's <c>Value2</c>.</param>
    /// <param name="headerMap">The header map for this sheet, from <see cref="BuildHeaderMap"/>.</param>
    /// <param name="columns">The configured column headers for the Clients worksheet.</param>
    /// <returns>One entry per non-blank data row.</returns>
    public static IReadOnlyList<(Client Client, int RowNumber)> ParseClients(
        object[,] data, IReadOnlyDictionary<string, int> headerMap, ClientColumnsConfig columns)
    {
        var rows = data.GetLength(0);
        var cols = data.GetLength(1);
        var results = new List<(Client, int)>();

        var idCol = headerMap[columns.Id];
        var nameCol = headerMap[columns.Name];
        var contactCol = headerMap[columns.Contact];
        var address1Col = headerMap[columns.Address1];
        var address2Col = headerMap[columns.Address2];
        var townCol = headerMap[columns.Town];
        var postcodeCol = headerMap[columns.Postcode];
        var emailCol = headerMap[columns.Email];
        var dueDateCol = headerMap[columns.DueDate];
        var lastPaidCol = headerMap[columns.LastPaidDate];
        var activeCol = headerMap[columns.Active];

        for (var row = 2; row <= rows; row++)
        {
            if (IsRowBlank(data, row, cols))
            {
                continue;
            }

            var id = GetString(data, row, idCol);
            if (string.IsNullOrWhiteSpace(id))
            {
                continue;
            }

            var client = new Client
            {
                ClientId = id,
                ClientName = GetString(data, row, nameCol),
                ContactName = GetString(data, row, contactCol),
                Address1 = GetString(data, row, address1Col),
                Address2 = GetString(data, row, address2Col),
                Town = GetString(data, row, townCol),
                Postcode = GetString(data, row, postcodeCol),
                Email = GetString(data, row, emailCol),
                DueDate = GetDate(data, row, dueDateCol),
                LastPaidDate = GetDate(data, row, lastPaidCol),
                Active = ParseActive(GetString(data, row, activeCol)),
            };

            results.Add((client, row));
        }

        return results;
    }

    /// <summary>
    /// Parses every data row of the Items worksheet into a <see cref="BillingItem"/>, alongside
    /// the 1-based worksheet row it came from.
    /// </summary>
    /// <param name="data">The raw sheet data, as returned by a used range's <c>Value2</c>.</param>
    /// <param name="headerMap">The header map for this sheet, from <see cref="BuildHeaderMap"/>.</param>
    /// <param name="columns">The configured column headers for the Items worksheet.</param>
    /// <returns>One entry per non-blank data row.</returns>
    public static IReadOnlyList<(BillingItem Item, int RowNumber)> ParseItems(
        object[,] data, IReadOnlyDictionary<string, int> headerMap, ItemColumnsConfig columns)
    {
        var rows = data.GetLength(0);
        var cols = data.GetLength(1);
        var results = new List<(BillingItem, int)>();

        var clientIdCol = headerMap[columns.ClientId];
        var descriptionCol = headerMap[columns.Description];
        var quantityCol = headerMap[columns.Quantity];
        var unitCostCol = headerMap[columns.UnitCost];
        var previousUnitCostCol = headerMap[columns.PreviousUnitCost];
        var lastUpliftPercentCol = headerMap[columns.LastUpliftPercent];

        for (var row = 2; row <= rows; row++)
        {
            if (IsRowBlank(data, row, cols))
            {
                continue;
            }

            var clientId = GetString(data, row, clientIdCol);
            if (string.IsNullOrWhiteSpace(clientId))
            {
                continue;
            }

            var item = new BillingItem
            {
                ClientId = clientId,
                Description = GetString(data, row, descriptionCol),
                Quantity = GetDecimal(data, row, quantityCol),
                UnitCost = GetDecimal(data, row, unitCostCol),
                PreviousUnitCost = GetNullableDecimal(data, row, previousUnitCostCol),
                LastUpliftPercent = GetNullableDecimal(data, row, lastUpliftPercentCol),
            };

            results.Add((item, row));
        }

        return results;
    }

    /// <summary>
    /// Parses every data row of the Invoices worksheet into an <see cref="InvoiceRecord"/>.
    /// </summary>
    /// <param name="data">The raw sheet data, as returned by a used range's <c>Value2</c>.</param>
    /// <param name="headerMap">The header map for this sheet, from <see cref="BuildHeaderMap"/>.</param>
    /// <param name="columns">The configured column headers for the Invoices worksheet.</param>
    /// <returns>One entry per non-blank data row.</returns>
    public static IReadOnlyList<InvoiceRecord> ParseInvoices(
        object[,] data, IReadOnlyDictionary<string, int> headerMap, InvoiceColumnsConfig columns)
    {
        var rows = data.GetLength(0);
        var cols = data.GetLength(1);
        var results = new List<InvoiceRecord>();

        var invoiceNumberCol = headerMap[columns.InvoiceNumber];
        var clientIdCol = headerMap[columns.ClientId];
        var invoiceDateCol = headerMap[columns.InvoiceDate];
        var dueDateCol = headerMap[columns.DueDate];
        var periodStartCol = headerMap[columns.PeriodStart];
        var periodEndCol = headerMap[columns.PeriodEnd];
        var netCol = headerMap[columns.Net];
        var vatCol = headerMap[columns.Vat];
        var totalCol = headerMap[columns.Total];
        var rpiPercentCol = headerMap[columns.RpiPercent];
        var docxPathCol = headerMap[columns.DocxPath];
        var pdfPathCol = headerMap[columns.PdfPath];

        for (var row = 2; row <= rows; row++)
        {
            if (IsRowBlank(data, row, cols))
            {
                continue;
            }

            var invoiceNumber = GetString(data, row, invoiceNumberCol);
            if (string.IsNullOrWhiteSpace(invoiceNumber))
            {
                continue;
            }

            results.Add(new InvoiceRecord
            {
                InvoiceNumber = invoiceNumber,
                ClientId = GetString(data, row, clientIdCol),
                InvoiceDate = GetDate(data, row, invoiceDateCol) ?? default,
                PayByDate = GetDate(data, row, dueDateCol) ?? default,
                PeriodStart = GetDate(data, row, periodStartCol) ?? default,
                PeriodEnd = GetDate(data, row, periodEndCol) ?? default,
                Net = GetDecimal(data, row, netCol),
                Vat = GetDecimal(data, row, vatCol),
                Total = GetDecimal(data, row, totalCol),
                RpiPercent = GetDecimal(data, row, rpiPercentCol),
                DocxPath = GetString(data, row, docxPathCol),
                PdfPath = GetString(data, row, pdfPathCol),
            });
        }

        return results;
    }

    private static string[] AllClientHeaders(ClientColumnsConfig c) =>
        [c.Id, c.Name, c.Contact, c.Address1, c.Address2, c.Town, c.Postcode, c.Email, c.DueDate, c.LastPaidDate, c.Active];

    private static string[] AllItemHeaders(ItemColumnsConfig c) =>
        [c.ClientId, c.Description, c.Quantity, c.UnitCost, c.PreviousUnitCost, c.LastUpliftPercent];

    private static string[] AllInvoiceHeaders(InvoiceColumnsConfig c) =>
        [c.InvoiceNumber, c.ClientId, c.InvoiceDate, c.DueDate, c.PeriodStart, c.PeriodEnd, c.Net, c.Vat, c.Total, c.RpiPercent, c.DocxPath, c.PdfPath];

    private static bool IsRowBlank(object[,] data, int row, int columnCount)
    {
        for (var col = 1; col <= columnCount; col++)
        {
            if (!string.IsNullOrWhiteSpace(GetString(data, row, col)))
            {
                return false;
            }
        }

        return true;
    }

    private static object? GetCell(object[,] data, int row, int col) => col <= 0 ? null : data[row, col];

    private static string GetString(object[,] data, int row, int col)
    {
        var value = GetCell(data, row, col);
        return value switch
        {
            null => string.Empty,
            string s => s.Trim(),
            double d => d.ToString(CultureInfo.InvariantCulture),
            _ => value.ToString()?.Trim() ?? string.Empty,
        };
    }

    private static decimal GetDecimal(object[,] data, int row, int col)
    {
        var value = GetCell(data, row, col);
        switch (value)
        {
            case null:
                return 0m;
            case double d:
                return (decimal)d;
            case string s when decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed):
                return parsed;
            case string s when decimal.TryParse(s, NumberStyles.Any, CultureInfo.GetCultureInfo("en-GB"), out var parsed):
                return parsed;
            default:
                return 0m;
        }
    }

    private static decimal? GetNullableDecimal(object[,] data, int row, int col)
    {
        var value = GetCell(data, row, col);
        if (value is null || (value is string s && string.IsNullOrWhiteSpace(s)))
        {
            return null;
        }

        return GetDecimal(data, row, col);
    }

    private static DateTime? GetDate(object[,] data, int row, int col)
    {
        var value = GetCell(data, row, col);
        switch (value)
        {
            case null:
                return null;
            case double oleDate:
                return DateTime.FromOADate(oleDate);
            case DateTime dt:
                return dt;
            case string s:
                var trimmed = s.Trim();
                if (string.IsNullOrEmpty(trimmed))
                {
                    return null;
                }

                if (DateTime.TryParse(trimmed, CultureInfo.GetCultureInfo("en-GB"), DateTimeStyles.None, out var parsedEnGb))
                {
                    return parsedEnGb;
                }

                return DateTime.TryParse(trimmed, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedInvariant)
                    ? parsedInvariant
                    : null;
            default:
                return null;
        }
    }

    private static bool ParseActive(string value) =>
        string.IsNullOrWhiteSpace(value) || !value.Trim().StartsWith("N", StringComparison.OrdinalIgnoreCase);
}
