# RenewalBilling: Design and Build Spec

Audience: Claude Code agent. Read all of this before writing code. Work through the milestones in order and stop for a check at each one.

## 1. Purpose

A Windows Forms app that Windows Task Scheduler launches once a week. It reads a client workbook in Excel, finds clients whose renewal is due within 21 days, proposes a new annual charge uplifted by the latest RPI, and lets the user approve each one in a review form. For every approved client it fills a Word invoice template, saves a .docx and a PDF copy, and writes the new cost and next renewal date back to the workbook.

Decisions already made with the owner:

| Topic | Decision |
|---|---|
| RPI source | Fetched online from ONS at each run |
| Run behaviour | Review form, user approves each invoice |
| Output | .docx plus PDF copy |
| Automation | COM (Excel and Word installed locally) |
| Stack | C# .NET 8 Windows Forms, no web components |

## 2. Tech and conventions

- Target `net8.0-windows`, `UseWindowsForms` true, nullable enabled, AnyCPU. Office runs out of process so bitness does not matter.
- COM: use the NuGet packages `Microsoft.Office.Interop.Excel` and `Microsoft.Office.Interop.Word`. If the packages misbehave under .NET 8, fall back to late binding with `dynamic` and `Type.GetTypeFromProgID`. Hide that choice behind the service interfaces so callers never know.
- All settings come from `.config.json` in the exe folder (System.Text.Json). Set the file to copy to output in the csproj. Nothing is hard coded.
- XML doc comments (`///`) on every public type and member.
- No third party packages except the two Interop packages and xUnit for tests.
- No em-dashes in any text, comments, UI strings or docs. Write plain, human sentences.
- Spelling and formats: British English, dates `dd MMM yyyy`, currency `en-GB` pounds.
- Licence: Creative Commons Attribution-NonCommercial-ShareAlike 4.0 International. Put a `LICENSE.md` with a short precis (see section 13) and link to the full terms.

## 3. Solution layout

```
RenewalBilling/
  RenewalBilling.sln
  README.md
  LICENSE.md
  install-task.ps1
  src/RenewalBilling/
    RenewalBilling.csproj
    .config.json
    Program.cs
    AppConfig.cs
    Models/
      Client.cs
      BillingItem.cs
      InvoiceCandidate.cs
      InvoiceRecord.cs
    Services/
      ExcelRepository.cs       (IClientRepository)
      WordInvoiceBuilder.cs    (IInvoiceBuilder)
      RpiService.cs            (IRpiService)
      UpliftCalculator.cs      (pure, static)
      EligibilityService.cs    (pure)
      InvoiceNumberService.cs  (pure)
      ComRelease.cs            (COM cleanup helpers)
      FileLogger.cs
    Forms/
      ReviewForm.cs (+ .Designer.cs)
  tests/RenewalBilling.Tests/
    UpliftCalculatorTests.cs
    EligibilityServiceTests.cs
    InvoiceNumberServiceTests.cs
    TagMapTests.cs
```

Keep pure logic (uplift, eligibility, numbering, tag map) free of COM so it can be unit tested.

## 4. Command line

| Argument | Effect |
|---|---|
| (none) or `/scheduled` | Normal run: find candidates, show review form only if there are any |
| `/dryrun` | Log what would be invoiced, show nothing, write nothing, create no Word files |
| `/date yyyy-MM-dd` | Override today's date, for testing |
| `/config <path>` | Use a different config file |
| `/makesamples` | Create a sample workbook and sample Word template in the configured paths, then exit |

If there are no candidates: log it and exit with code 0 without showing any window.
Exit codes: 0 ok, 1 config or workbook error, 2 Office not available, 3 unexpected failure.

## 5. Config: `.config.json`

```json
{
  "workbookPath": "C:\\Billing\\Clients.xlsx",
  "templatePath": "C:\\Billing\\InvoiceTemplate.dotx",
  "invoiceFolder": "C:\\Billing\\Invoices",
  "backupFolder": "C:\\Billing\\Backups",
  "logFolder": "C:\\Billing\\Logs",
  "leadDays": 21,
  "renewalMonths": 12,
  "invoiceNumberFormat": "INV-{yyyy}-{seq:0000}",
  "payByDays": null,
  "vatRatePercent": 0,
  "rpi": {
    "url": "https://www.ons.gov.uk/economy/inflationandpriceindices/timeseries/czbh/mm23/data",
    "timeoutSeconds": 15,
    "fallbackPercent": 3.5,
    "cacheFile": "rpi-cache.json"
  },
  "uplift": {
    "roundTo": 1.00,
    "minimumPercent": 0,
    "maximumPercent": null
  },
  "sheets": {
    "clients": "Clients",
    "items": "Items",
    "invoices": "Invoices"
  },
  "columns": {
    "clients": {
      "id": "ClientId", "name": "ClientName", "contact": "ContactName",
      "address1": "Address1", "address2": "Address2", "town": "Town",
      "postcode": "Postcode", "email": "Email",
      "dueDate": "DueDate", "lastPaidDate": "LastPaidDate", "active": "Active"
    },
    "items": {
      "clientId": "ClientId", "description": "Description", "quantity": "Quantity",
      "unitCost": "UnitCost", "previousUnitCost": "PreviousUnitCost", "lastUpliftPercent": "LastUpliftPercent"
    },
    "invoices": {
      "invoiceNumber": "InvoiceNumber", "clientId": "ClientId", "invoiceDate": "InvoiceDate",
      "dueDate": "DueDate", "periodStart": "PeriodStart", "periodEnd": "PeriodEnd",
      "net": "Net", "vat": "Vat", "total": "Total", "rpiPercent": "RpiPercent",
      "docxPath": "DocxPath", "pdfPath": "PdfPath"
    }
  }
}
```

Validate at start-up. Missing file, missing paths, or missing headers must produce a clear log line and exit code 1. Relative paths resolve against the exe folder.

## 6. Workbook model

Three sheets. Column order does not matter, headers are found by name using the `columns` map. Row 1 is the header row.

**Clients** (one row per client): ClientId, ClientName, ContactName, Address1, Address2, Town, Postcode, Email, DueDate, LastPaidDate, Active (Y or N, blank means Y).

**Items** (one row per billable line): ClientId, Description, Quantity, UnitCost (the current annual price), PreviousUnitCost, LastUpliftPercent.

**Invoices** (append only log): InvoiceNumber, ClientId, InvoiceDate, DueDate, PeriodStart, PeriodEnd, Net, Vat, Total, RpiPercent, DocxPath, PdfPath.

Reading rules:
- Read each sheet's used range in one go into an object array (`Value2`). Never loop cell by cell over COM for reads.
- Dates may arrive as a double (OLE date) or a string. Handle both with `DateTime.FromOADate` and a tolerant `en-GB` parse. Blank DueDate means skip the client and log a warning.
- Trim strings. Skip fully blank rows.

## 7. Business rules

**Eligible client**: Active, DueDate not blank, `DueDate <= runDate + leadDays`, and no row in Invoices for that ClientId with the same DueDate. Overdue clients with no invoice logged are included and flagged "Overdue" in the form. Clients with no Items rows are skipped with a warning.

**Uplift** per item: `proposed = round(unitCost * (1 + rpi/100) / roundTo) * roundTo`. Clamp the percentage between `minimumPercent` and `maximumPercent` (if set) before applying. Use `decimal`, never `double`, for money. Use `MidpointRounding.AwayFromZero`.

**Totals**: Net is the sum of `quantity * proposedUnitCost`. VAT is `net * vatRatePercent / 100`, rounded to pence. Total is Net plus VAT.

**Invoice dates**: InvoiceDate is the run date. PeriodStart is the client's DueDate. PeriodEnd is `DueDate + renewalMonths - 1 day`. PayBy is `invoiceDate + payByDays` if set, otherwise the DueDate.

**Invoice number**: built from `invoiceNumberFormat`. `{yyyy}` is the invoice year and `{seq:0000}` is the highest existing sequence for that year in the Invoices sheet plus one. Parse existing numbers with a regex derived from the format. Within one run, increment in memory so approvals in the same session never collide.

**On approval** (all of the following, in this order, only after the Word and PDF files exist):
1. Append a row to Invoices.
2. For each item: set PreviousUnitCost to the old UnitCost, UnitCost to the approved new price, LastUpliftPercent to the RPI used.
3. Advance the client's DueDate by `renewalMonths`.
4. Leave LastPaidDate alone. Only the user changes it when payment arrives.

If any step fails, do not write partial changes for that client. Log the error and continue with the next client.

## 8. RPI service

Series: ONS CZBH, RPI all items, 12 month percentage change, dataset MM23.

1. GET the configured URL with `HttpClient` and the timeout from config. Parse the JSON, read the `months` array and take the most recent entry with a numeric `value`. Expose the value and its `date` label (for example `2026 AUG`).
2. On success, write it to the cache file in the exe folder.
3. On failure (offline, bad JSON, timeout), use the cache file. If there is no cache, use `fallbackPercent`.
4. Return a result object: Percent, PeriodLabel, Source (Online, Cache, Fallback). The review form must show the source clearly, and show a warning banner for Cache and Fallback.
5. Before coding this, fetch the URL once with curl or PowerShell and confirm the real JSON shape. Adapt the parser to what you actually see. If ONS has moved the endpoint, say so in the final report and keep the fallback path working.
6. The user can overwrite the RPI percentage in the review form. Recalculate all proposals when it changes.

## 9. Review form

Master and detail layout, WinForms, resizable, sensible tab order, high DPI aware.

- Top bar: run date, RPI percentage (editable `NumericUpDown`), RPI period label, source label, warning banner when not Online.
- Left: `DataGridView` of candidates. Columns: Approve (checkbox), Client, Due date, Status (Due or Overdue), Net, Total. Default Approve to checked.
- Right: detail panel for the selected client. Shows the address block, then an editable grid of items: Description, Qty, Current price, Proposed price (editable), Uplift %. Subtotal, VAT and Total update live.
- Bottom: `Create approved invoices` and `Close` buttons. A progress bar and a status label while building. Disable the form during the build and keep the UI responsive by running COM work on a single dedicated STA thread (`Thread` with `SetApartmentState(STA)`), reporting progress back to the UI.
- On finish, show a summary list: created, failed, skipped, with a button to open the invoice folder.
- Closing the form without creating anything writes nothing. Unapproved clients simply appear again next week.

Do not put CSS-like styling hacks in code. Keep the look plain and consistent: system font, standard controls.

## 10. Word template and tags

Template is a .dotx or .docx at `templatePath`. Tags use the `##TAG##` form. Replace in the main body, tables, headers, footers and text boxes by looping over every `StoryRange` and its `NextStoryRange`.

Scalar tags:

| Tag | Value |
|---|---|
| `##CLIENTNAME##` | Client name |
| `##CONTACTNAME##` | Contact name |
| `##ADDRESS1##` `##ADDRESS2##` `##TOWN##` `##POSTCODE##` | Address parts |
| `##EMAIL##` | Client email |
| `##INVOICENO##` | Invoice number |
| `##INVOICEDATE##` | Invoice date |
| `##PAYBYDATE##` | Pay-by date |
| `##PERIODSTART##` `##PERIODEND##` | Renewal period |
| `##RPI##` | RPI percentage used, one decimal place |
| `##SUBTOTAL##` `##VAT##` `##TOTAL##` | Money, formatted `£#,##0.00` |

Line item table: the template contains one table with a header row and one body row whose cells hold `##ITEM##`, `##QTY##`, `##UNITCOST##`, `##LINETOTAL##`. Find the row that contains `##ITEM##`, add a new row for each additional item by copying that row, then fill each. Do this before the scalar replacement.

Implementation notes:
- Create the invoice with `Documents.Add(templatePath)` so the template itself is never modified.
- Use `Find.Execute` with `Replace:=wdReplaceAll` and `MatchWildcards:=false`. Find handles tags that Word has split across runs. Replacement text over 255 characters fails in Find, so for long text set the found range's `.Text` directly in a loop.
- After replacing, scan for any leftover `##[A-Z0-9]+##` text. Log a warning naming the tag, and include it in the summary.
- Save as `{invoiceFolder}\{yyyy}\{InvoiceNo} {ClientName}.docx` using `SaveAs2` with `wdFormatXMLDocument`. Sanitise file names (strip `\ / : * ? " < > |`). Create folders as needed. Never overwrite: if the file exists, fail that client with a clear message.
- Export the PDF next to it with `ExportAsFixedFormat` (`wdExportFormatPDF`).
- Close the document with `SaveChanges:=false` after saving. Set `Visible = false` and `DisplayAlerts = wdAlertsNone` on the Word application. Quit Word only if this app started it.

Template font note for `/makesamples`: use Trebuchet MS for text.

## 11. Excel handling

- Start a private `Excel.Application` instance (`Visible=false`, `DisplayAlerts=false`, `ScreenUpdating=false`). Do not attach to a running instance.
- Before opening, test whether the file is locked (try opening with `FileShare.None`). If locked, log "Workbook is open elsewhere, close it and run again" and exit with code 1. Show a message box only if running interactively with candidates.
- Back up first: copy the workbook to `backupFolder\Clients_{yyyyMMdd_HHmmss}.xlsx` before any write. Keep the newest 12 backups, delete older ones.
- Write using a single batch per sheet where practical. Save with `Workbook.Save()` once after all approved clients are processed, then close and quit.
- Preserve existing formatting. Write dates as true Excel dates, money as numbers.

COM cleanup is critical so that no orphan `EXCEL.EXE` or `WINWORD.EXE` is left running:
- Avoid chained "double dot" calls on COM objects. Assign each COM object to a variable.
- Release every object with `Marshal.FinalReleaseComObject` in a `finally` block, in reverse order of creation.
- After release call `GC.Collect(); GC.WaitForPendingFinalizers();` twice.
- Put this in `ComRelease.cs` and use it everywhere.
- Test by running the app 5 times and confirming Task Manager shows no leftover Office processes.

## 12. Task Scheduler

COM automation needs an interactive user session, and the review form needs one too. So the task must run as the logged-in user, only when that user is logged on.

Provide `install-task.ps1` that registers a weekly task (Monday 08:30 by default, parameters for day and time) with `Register-ScheduledTask`:
- Action: the built exe with argument `/scheduled`, working directory set to the exe folder.
- Trigger: weekly.
- Settings: `StartWhenAvailable` (run after a missed start), `AllowStartIfOnBatteries`, `DontStopIfGoingOnBatteries`, execution time limit 1 hour.
- Principal: current user, `LogonType Interactive`, run level Limited.

Add a matching `uninstall-task` switch. Document manual setup in the README as well.

## 13. Logging, README and licence

- `FileLogger`: one file per day in `logFolder`, `RenewalBilling_yyyyMMdd.log`, lines like `2026-10-04 08:30:12 INFO message`. Log: run start and arguments, RPI result and source, each candidate found, each invoice created, each skip or failure with the reason, run end with counts. Never log more personal data than client id and name.
- `README.md`: what it does, requirements (Windows, Office installed, .NET 8 runtime), config reference, workbook layout, template tags table, how to schedule it, how to test with `/dryrun` and `/date`, and troubleshooting (locked workbook, leftover tags, ONS offline).
- `LICENSE.md`: licensed under Creative Commons Attribution-NonCommercial-ShareAlike 4.0 International. Precis: you may copy and share the work in any format, and adapt it; you must give credit and indicate changes; you may not use it commercially; if you adapt it you must share your version under the same licence. Full terms: https://creativecommons.org/licenses/by-nc-sa/4.0/ and legal code at https://creativecommons.org/licenses/by-nc-sa/4.0/legalcode

## 14. Milestones

Stop and report after each one.

1. **Scaffold.** Solution, projects, config loading and validation, logger, command line parsing. Build clean.
2. **Pure logic plus tests.** Uplift, eligibility, invoice numbering, tag map. All xUnit tests green. Include cases: leap day renewals, overdue client, already invoiced, rounding at midpoints, negative RPI with minimum 0, mixed number formats in the Invoices sheet.
3. **Excel repository.** Read all three sheets, back up, write-back, clean COM release. Add `/makesamples` for the workbook so there is data to test with.
4. **RPI service.** Online fetch, cache, fallback. Show the three source paths working (disconnect the network to prove the fallback).
5. **Word builder.** Template fill, line item rows, header and footer tags, docx plus PDF. Extend `/makesamples` to produce the sample template. Open the output and check it visually.
6. **Review form.** Wire it all together with the background STA worker and progress.
7. **Scheduling and docs.** `install-task.ps1`, README, LICENSE. Do a real scheduled trigger test with the task set a couple of minutes ahead.

## 15. Acceptance checks

- Run with `/dryrun` against sample data lists exactly the clients whose DueDate is within 21 days, plus any overdue and uninvoiced.
- Approving one client creates a .docx and .pdf, adds one Invoices row, updates the item prices, advances the DueDate by 12 months, and leaves LastPaidDate unchanged.
- Running the app again straight afterwards does not offer the same client.
- Opening the workbook in Excel, then running the app, ends cleanly with exit code 1 and a clear log line.
- A template with an unknown tag produces a logged warning, not a crash.
- No `EXCEL.EXE` or `WINWORD.EXE` remains after a run, including after a failure.
- With the network off, the form shows the warning banner and uses the cached or fallback RPI.
- Zero candidates means no window appears.

## 16. Out of scope for version 1

Emailing invoices, payment tracking, credit notes, multiple currencies, a database. Note these in the README as possible later work.
