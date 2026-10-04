using System.Runtime.InteropServices;
using System.Text;
using RenewalBilling.Models;

namespace RenewalBilling.Services;

/// <summary>
/// Thrown when an invoice cannot be built: the template is missing a required element, or the
/// target file already exists.
/// </summary>
public sealed class InvoiceBuildException : Exception
{
    /// <summary>Creates a new <see cref="InvoiceBuildException"/> with a message describing the problem.</summary>
    /// <param name="message">A clear, human readable description of what is wrong.</param>
    public InvoiceBuildException(string message) : base(message)
    {
    }
}

/// <summary>
/// The files produced for one invoice.
/// </summary>
/// <param name="DocxPath">Full path to the saved .docx copy.</param>
/// <param name="PdfPath">Full path to the saved PDF copy.</param>
/// <param name="UnknownTags">Any <c>##TAG##</c> pattern still present after replacement, in case the template references a tag this app does not know about.</param>
public sealed record InvoiceFiles(string DocxPath, string PdfPath, IReadOnlyList<string> UnknownTags);

/// <summary>
/// Fills the invoice template for an approved candidate and saves a .docx and PDF copy.
/// </summary>
public interface IInvoiceBuilder : IDisposable
{
    /// <summary>
    /// Builds and saves one invoice.
    /// </summary>
    /// <param name="candidate">The approved candidate, with its invoice number already assigned.</param>
    /// <param name="templatePath">Full path to the Word template.</param>
    /// <param name="invoiceFolder">Folder that finished invoices are saved under.</param>
    /// <returns>The paths of the files that were created, and any tag left unreplaced.</returns>
    InvoiceFiles CreateInvoice(InvoiceCandidate candidate, string templatePath, string invoiceFolder);
}

/// <summary>
/// Automates a private Word instance to fill the invoice template described in section 10 of the
/// specification. One Word instance is started lazily on the first invoice and reused for every
/// subsequent one in the same session, then quit once when this builder is disposed.
///
/// Like <see cref="ExcelRepository"/>, Word is automated through late-bound <c>dynamic</c> COM
/// calls against a ProgID (<c>Word.Application</c>) rather than the typed
/// <c>Microsoft.Office.Interop.Word</c> PIA, to avoid that assembly's dependency on the old
/// "office.dll" core PIA that a Click-to-Run Office install does not register.
/// </summary>
public sealed class WordInvoiceBuilder : IInvoiceBuilder
{
    // WdSaveFormat.wdFormatXMLDocument (.docx) and WdExportFormat.wdExportFormatPDF, used as
    // plain ints so this class never has to reference the Office Interop PIA.
    private const int WdFormatXmlDocument = 12;
    private const int WdFormatXmlTemplate = 14;
    private const int WdExportFormatPdf = 17;
    private const int WdReplaceAll = 2;
    private const int WdReplaceNone = 0;
    private const int WdFindStop = 0;
    private const int WdAlertsNone = 0;
    private const int WdHeaderFooterPrimary = 1;

    private static readonly char[] InvalidFileNameChars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|'];

    private readonly IAppLogger _logger;
    private dynamic? _word;
    private int? _wordProcessId;

    /// <summary>
    /// Creates an invoice builder. Word is not started until the first call to
    /// <see cref="CreateInvoice"/>.
    /// </summary>
    /// <param name="logger">Where to log leftover-tag warnings.</param>
    public WordInvoiceBuilder(IAppLogger logger)
    {
        _logger = logger;
    }

    /// <inheritdoc />
    public InvoiceFiles CreateInvoice(InvoiceCandidate candidate, string templatePath, string invoiceFolder)
    {
        if (!File.Exists(templatePath))
        {
            throw new InvoiceBuildException($"Template not found: {templatePath}");
        }

        if (string.IsNullOrWhiteSpace(candidate.InvoiceNumber))
        {
            throw new InvoiceBuildException($"Candidate {candidate.Client.ClientId} has no invoice number assigned.");
        }

        var yearFolder = Path.Combine(invoiceFolder, candidate.InvoiceDate.Year.ToString());
        Directory.CreateDirectory(yearFolder);

        var fileBaseName = SanitizeFileName($"{candidate.InvoiceNumber} {candidate.Client.ClientName}");
        var docxPath = Path.Combine(yearFolder, fileBaseName + ".docx");
        var pdfPath = Path.Combine(yearFolder, fileBaseName + ".pdf");

        if (File.Exists(docxPath))
        {
            throw new InvoiceBuildException($"'{docxPath}' already exists. Not overwriting.");
        }

        if (File.Exists(pdfPath))
        {
            throw new InvoiceBuildException($"'{pdfPath}' already exists. Not overwriting.");
        }

        EnsureWordStarted();

        dynamic documents = _word!.Documents;
        dynamic? document = null;
        try
        {
            document = documents.Add(templatePath);

            FillLineItemTable(document, candidate);
            ReplaceScalarTags(document, candidate);

            var unknownTags = FindLeftoverTags(document);
            foreach (var tag in unknownTags)
            {
                _logger.Warn($"Invoice {candidate.InvoiceNumber}: leftover tag '{tag}' was not replaced.");
            }

            document.SaveAs2(docxPath, WdFormatXmlDocument);
            document.ExportAsFixedFormat(pdfPath, WdExportFormatPdf);

            return new InvoiceFiles(docxPath, pdfPath, unknownTags);
        }
        finally
        {
            try
            {
                document?.Close(false);
            }
            catch (COMException)
            {
                // Best effort; we still need to release the document below.
            }

            if (document is not null)
            {
                Marshal.FinalReleaseComObject(document);
            }

            Marshal.FinalReleaseComObject(documents);
            ComRelease.CollectTwice();
        }
    }

    /// <summary>
    /// Creates a sample invoice template at <paramref name="templatePath"/>, with the Normal
    /// style set to Trebuchet MS and every tag from section 10 of the specification present in
    /// the body, header and footer, plus a two-row item table so the row-duplication logic has
    /// something real to exercise.
    /// </summary>
    /// <param name="templatePath">Full path to save the sample template to. Saved as a .dotx template if the path ends in ".dotx", otherwise as a .docx.</param>
    public void CreateSampleTemplate(string templatePath)
    {
        var folder = Path.GetDirectoryName(templatePath);
        if (!string.IsNullOrEmpty(folder))
        {
            Directory.CreateDirectory(folder);
        }

        EnsureWordStarted();

        dynamic documents = _word!.Documents;
        dynamic? document = null;
        try
        {
            document = documents.Add();

            SetDefaultFont(document, "Trebuchet MS");
            WriteSampleHeaderAndFooter(document);
            WriteSampleBody(document);

            if (File.Exists(templatePath))
            {
                File.Delete(templatePath);
            }

            var format = templatePath.EndsWith(".dotx", StringComparison.OrdinalIgnoreCase)
                ? WdFormatXmlTemplate
                : WdFormatXmlDocument;

            document.SaveAs2(templatePath, format);
        }
        finally
        {
            try
            {
                document?.Close(false);
            }
            catch (COMException)
            {
            }

            if (document is not null)
            {
                Marshal.FinalReleaseComObject(document);
            }

            Marshal.FinalReleaseComObject(documents);
            ComRelease.CollectTwice();
        }
    }

    private static void SetDefaultFont(dynamic document, string fontName)
    {
        dynamic styles = document.Styles;
        try
        {
            dynamic normal = styles["Normal"];
            try
            {
                dynamic font = normal.Font;
                try
                {
                    font.Name = fontName;
                }
                finally
                {
                    Marshal.FinalReleaseComObject(font);
                }
            }
            finally
            {
                Marshal.FinalReleaseComObject(normal);
            }
        }
        finally
        {
            Marshal.FinalReleaseComObject(styles);
        }
    }

    private static void WriteSampleHeaderAndFooter(dynamic document)
    {
        dynamic sections = document.Sections;
        try
        {
            dynamic section = sections[1];
            try
            {
                SetHeaderFooterText(section, isHeader: true, text: "Renewal Billing Ltd");
                SetHeaderFooterText(section, isHeader: false, text: "Invoice ##INVOICENO##");
            }
            finally
            {
                Marshal.FinalReleaseComObject(section);
            }
        }
        finally
        {
            Marshal.FinalReleaseComObject(sections);
        }
    }

    private static void SetHeaderFooterText(dynamic section, bool isHeader, string text)
    {
        dynamic collection = isHeader ? section.Headers : section.Footers;
        try
        {
            dynamic headerFooter = collection[WdHeaderFooterPrimary];
            try
            {
                dynamic range = headerFooter.Range;
                try
                {
                    range.Text = text;
                }
                finally
                {
                    Marshal.FinalReleaseComObject(range);
                }
            }
            finally
            {
                Marshal.FinalReleaseComObject(headerFooter);
            }
        }
        finally
        {
            Marshal.FinalReleaseComObject(collection);
        }
    }

    private static void WriteSampleBody(dynamic document)
    {
        AppendParagraph(document, "INVOICE");
        AppendParagraph(document, string.Empty);
        AppendParagraph(document, "##CLIENTNAME##");
        AppendParagraph(document, "##CONTACTNAME##");
        AppendParagraph(document, "##ADDRESS1##");
        AppendParagraph(document, "##ADDRESS2##");
        AppendParagraph(document, "##TOWN##");
        AppendParagraph(document, "##POSTCODE##");
        AppendParagraph(document, "##EMAIL##");
        AppendParagraph(document, string.Empty);
        AppendParagraph(document, "Invoice number: ##INVOICENO##");
        AppendParagraph(document, "Invoice date: ##INVOICEDATE##");
        AppendParagraph(document, "Pay by: ##PAYBYDATE##");
        AppendParagraph(document, "Period: ##PERIODSTART## to ##PERIODEND##");
        AppendParagraph(document, "RPI used: ##RPI##%");
        AppendParagraph(document, string.Empty);

        InsertSampleItemTable(document);

        AppendParagraph(document, string.Empty);
        AppendParagraph(document, "Subtotal: ##SUBTOTAL##");
        AppendParagraph(document, "VAT: ##VAT##");
        AppendParagraph(document, "Total: ##TOTAL##");
    }

    private static void AppendParagraph(dynamic document, string text)
    {
        dynamic content = document.Content;
        try
        {
            content.InsertAfter(text + "\r");
        }
        finally
        {
            Marshal.FinalReleaseComObject(content);
        }
    }

    private static void InsertSampleItemTable(dynamic document)
    {
        dynamic content = document.Content;
        int endPosition;
        try
        {
            endPosition = (int)content.End;
        }
        finally
        {
            Marshal.FinalReleaseComObject(content);
        }

        dynamic insertionPoint = document.Range(endPosition - 1, endPosition - 1);
        dynamic tables = document.Tables;
        try
        {
            dynamic table = tables.Add(insertionPoint, 2, 4);
            try
            {
                dynamic rows = table.Rows;
                try
                {
                    dynamic headerRow = rows[1];
                    try
                    {
                        SetCellText(headerRow, 1, "Description");
                        SetCellText(headerRow, 2, "Qty");
                        SetCellText(headerRow, 3, "Unit cost");
                        SetCellText(headerRow, 4, "Line total");
                    }
                    finally
                    {
                        Marshal.FinalReleaseComObject(headerRow);
                    }

                    dynamic bodyRow = rows[2];
                    try
                    {
                        SetCellText(bodyRow, 1, "##ITEM##");
                        SetCellText(bodyRow, 2, "##QTY##");
                        SetCellText(bodyRow, 3, "##UNITCOST##");
                        SetCellText(bodyRow, 4, "##LINETOTAL##");
                    }
                    finally
                    {
                        Marshal.FinalReleaseComObject(bodyRow);
                    }
                }
                finally
                {
                    Marshal.FinalReleaseComObject(rows);
                }
            }
            finally
            {
                Marshal.FinalReleaseComObject(table);
            }
        }
        finally
        {
            Marshal.FinalReleaseComObject(tables);
            Marshal.FinalReleaseComObject(insertionPoint);
        }
    }

    /// <inheritdoc />
    public void Dispose()
    {
        if (_word is null)
        {
            return;
        }

        ComRelease.QuitAndEnsureProcessExits(_word, _wordProcessId);
        _word = null;
        _wordProcessId = null;
    }

    private void EnsureWordStarted()
    {
        if (_word is not null)
        {
            return;
        }

        (_word, _wordProcessId) = ComRelease.StartOfficeApplication("Word.Application", "WINWORD");
        _word.Visible = false;
        _word.DisplayAlerts = WdAlertsNone;
    }

    private static void ReplaceScalarTags(dynamic document, InvoiceCandidate candidate)
    {
        foreach (var (tag, value) in TagMapBuilder.BuildScalarTags(candidate))
        {
            ReplaceTag(document, tag, value);
        }
    }

    /// <summary>
    /// Replaces every occurrence of a tag across the whole document: the main body, tables,
    /// headers, footers and text boxes, by walking every story range and its
    /// <c>NextStoryRange</c> chain.
    /// </summary>
    private static void ReplaceTag(dynamic document, string tag, string value)
    {
        dynamic storyRanges = document.StoryRanges;
        try
        {
            // StoryRanges is indexed by WdStoryType, not by a contiguous 1..Count position: a
            // document without footnotes, comments, or a particular header/footer simply has no
            // entry at that story type's index. Enumerating with foreach visits only the story
            // types actually present, which manual 1..Count indexing does not.
            foreach (dynamic story in storyRanges)
            {
                try
                {
                    ReplaceInChain(story, tag, value);
                }
                finally
                {
                    Marshal.FinalReleaseComObject(story);
                }
            }
        }
        finally
        {
            Marshal.FinalReleaseComObject(storyRanges);
        }
    }

    private static void ReplaceInChain(dynamic firstRange, string tag, string value)
    {
        dynamic current = firstRange;
        var ownsCurrent = false;

        while (true)
        {
            ReplaceInRange(current, tag, value);

            dynamic? next = null;
            try
            {
                next = current.NextStoryRange;
            }
            catch (COMException)
            {
                // No further linked story; this chain is done.
            }

            if (ownsCurrent)
            {
                Marshal.FinalReleaseComObject(current);
            }

            if (next is null)
            {
                return;
            }

            current = next;
            ownsCurrent = true;
        }
    }

    private static void ReplaceInRange(dynamic range, string tag, string value)
    {
        if (value.Length <= 255)
        {
            dynamic find = range.Find;
            try
            {
                find.Execute(FindText: tag, MatchWildcards: false, Forward: true, Wrap: WdFindStop, Replace: WdReplaceAll, ReplaceWith: value);
            }
            finally
            {
                Marshal.FinalReleaseComObject(find);
            }

            return;
        }

        // Find's own Replacement fails for text over 255 characters, so locate each occurrence
        // one at a time (Replace:=wdReplaceNone) and set the found range's .Text directly. Find
        // mutates `range` in place to the matched text, so the next Execute call continues
        // searching forward from there.
        while (true)
        {
            dynamic find = range.Find;
            bool found;
            try
            {
                found = (bool)find.Execute(FindText: tag, MatchWildcards: false, Forward: true, Wrap: WdFindStop, Replace: WdReplaceNone);
            }
            finally
            {
                Marshal.FinalReleaseComObject(find);
            }

            if (!found)
            {
                return;
            }

            range.Text = value;
        }
    }

    private static IReadOnlyList<string> FindLeftoverTags(dynamic document)
    {
        var text = new StringBuilder();
        dynamic storyRanges = document.StoryRanges;
        try
        {
            // See the matching comment in ReplaceTag: StoryRanges is sparse, so this must
            // enumerate rather than index 1..Count.
            foreach (dynamic story in storyRanges)
            {
                try
                {
                    AppendChainText(story, text);
                }
                finally
                {
                    Marshal.FinalReleaseComObject(story);
                }
            }
        }
        finally
        {
            Marshal.FinalReleaseComObject(storyRanges);
        }

        return TagMapBuilder.FindUnknownTags(text.ToString());
    }

    private static void AppendChainText(dynamic firstRange, StringBuilder text)
    {
        dynamic current = firstRange;
        var ownsCurrent = false;

        while (true)
        {
            text.Append((string)current.Text).Append(' ');

            dynamic? next = null;
            try
            {
                next = current.NextStoryRange;
            }
            catch (COMException)
            {
            }

            if (ownsCurrent)
            {
                Marshal.FinalReleaseComObject(current);
            }

            if (next is null)
            {
                return;
            }

            current = next;
            ownsCurrent = true;
        }
    }

    /// <summary>
    /// The line item table's location: which table, which row holds the <c>##ITEM##</c> cell,
    /// and which column each of the four tags is in (detected from the template rather than
    /// assumed, since a template author can reorder the columns).
    /// </summary>
    private readonly record struct TableLocation(int TableIndex, int RowIndex, int ItemCol, int QtyCol, int UnitCostCol, int LineTotalCol);

    private void FillLineItemTable(dynamic document, InvoiceCandidate candidate)
    {
        var lineItems = TagMapBuilder.BuildLineItemTags(candidate);
        if (lineItems.Count == 0)
        {
            throw new InvoiceBuildException($"Candidate {candidate.Client.ClientId} has no line items to invoice.");
        }

        var location = FindItemTemplateRow(document);

        if (lineItems.Count > 1)
        {
            dynamic tables = document.Tables;
            try
            {
                dynamic table = tables[location.TableIndex];
                try
                {
                    for (var i = 1; i < lineItems.Count; i++)
                    {
                        AddItemRow(table, location);
                    }
                }
                finally
                {
                    Marshal.FinalReleaseComObject(table);
                }
            }
            finally
            {
                Marshal.FinalReleaseComObject(tables);
            }
        }

        FillItemRows(document, location, lineItems);
    }

    private static TableLocation FindItemTemplateRow(dynamic document)
    {
        dynamic tables = document.Tables;
        try
        {
            var tableCount = (int)tables.Count;
            for (var t = 1; t <= tableCount; t++)
            {
                dynamic table = tables[t];
                try
                {
                    dynamic rows = table.Rows;
                    try
                    {
                        var rowCount = (int)rows.Count;
                        for (var r = 1; r <= rowCount; r++)
                        {
                            dynamic row = rows[r];
                            try
                            {
                                var (itemCol, qtyCol, unitCostCol, lineTotalCol) = FindTagColumns((object)row);
                                if (itemCol > 0)
                                {
                                    return new TableLocation(t, r, itemCol, qtyCol, unitCostCol, lineTotalCol);
                                }
                            }
                            finally
                            {
                                Marshal.FinalReleaseComObject(row);
                            }
                        }
                    }
                    finally
                    {
                        Marshal.FinalReleaseComObject(rows);
                    }
                }
                finally
                {
                    Marshal.FinalReleaseComObject(table);
                }
            }
        }
        finally
        {
            Marshal.FinalReleaseComObject(tables);
        }

        throw new InvoiceBuildException("Template is missing a line item table row containing ##ITEM##.");
    }

    private static (int ItemCol, int QtyCol, int UnitCostCol, int LineTotalCol) FindTagColumns(dynamic row)
    {
        int itemCol = 0, qtyCol = 0, unitCostCol = 0, lineTotalCol = 0;

        dynamic cells = row.Cells;
        try
        {
            var cellCount = (int)cells.Count;
            for (var c = 1; c <= cellCount; c++)
            {
                dynamic cell = cells[c];
                try
                {
                    switch (CleanCellText(cell))
                    {
                        case "##ITEM##": itemCol = c; break;
                        case "##QTY##": qtyCol = c; break;
                        case "##UNITCOST##": unitCostCol = c; break;
                        case "##LINETOTAL##": lineTotalCol = c; break;
                    }
                }
                finally
                {
                    Marshal.FinalReleaseComObject(cell);
                }
            }
        }
        finally
        {
            Marshal.FinalReleaseComObject(cells);
        }

        return (itemCol, qtyCol, unitCostCol, lineTotalCol);
    }

    private static string CleanCellText(dynamic cell)
    {
        dynamic range = cell.Range;
        try
        {
            // Word always appends a cell-mark (\a) and paragraph-mark (\r) to a cell's text.
            return ((string)range.Text).Replace("\r", string.Empty).Replace("\a", string.Empty).Trim();
        }
        finally
        {
            Marshal.FinalReleaseComObject(range);
        }
    }

    private static void AddItemRow(dynamic table, TableLocation location)
    {
        dynamic rows = table.Rows;
        try
        {
            // Rows.Add takes an optional BeforeRow; omitted entirely (not a named AfterRow,
            // which this method does not have), it appends a new row at the end of the table.
            dynamic newRow = rows.Add();
            try
            {
                SetCellText(newRow, location.ItemCol, "##ITEM##");
                SetCellText(newRow, location.QtyCol, "##QTY##");
                SetCellText(newRow, location.UnitCostCol, "##UNITCOST##");
                SetCellText(newRow, location.LineTotalCol, "##LINETOTAL##");
            }
            finally
            {
                Marshal.FinalReleaseComObject(newRow);
            }
        }
        finally
        {
            Marshal.FinalReleaseComObject(rows);
        }
    }

    private static void FillItemRows(dynamic document, TableLocation location, IReadOnlyList<LineItemTags> lineItems)
    {
        dynamic tables = document.Tables;
        try
        {
            dynamic table = tables[location.TableIndex];
            try
            {
                dynamic rows = table.Rows;
                try
                {
                    for (var i = 0; i < lineItems.Count; i++)
                    {
                        dynamic row = rows[location.RowIndex + i];
                        try
                        {
                            SetCellText(row, location.ItemCol, lineItems[i].Item);
                            SetCellText(row, location.QtyCol, lineItems[i].Qty);
                            SetCellText(row, location.UnitCostCol, lineItems[i].UnitCost);
                            SetCellText(row, location.LineTotalCol, lineItems[i].LineTotal);
                        }
                        finally
                        {
                            Marshal.FinalReleaseComObject(row);
                        }
                    }
                }
                finally
                {
                    Marshal.FinalReleaseComObject(rows);
                }
            }
            finally
            {
                Marshal.FinalReleaseComObject(table);
            }
        }
        finally
        {
            Marshal.FinalReleaseComObject(tables);
        }
    }

    private static void SetCellText(dynamic row, int column, string text)
    {
        if (column <= 0)
        {
            return;
        }

        dynamic cells = row.Cells;
        try
        {
            dynamic cell = cells[column];
            try
            {
                dynamic range = cell.Range;
                try
                {
                    range.Text = text;
                }
                finally
                {
                    Marshal.FinalReleaseComObject(range);
                }
            }
            finally
            {
                Marshal.FinalReleaseComObject(cell);
            }
        }
        finally
        {
            Marshal.FinalReleaseComObject(cells);
        }
    }

    /// <summary>
    /// Strips the characters Windows does not allow in a file name.
    /// </summary>
    /// <param name="name">The proposed name, before the file extension.</param>
    /// <returns>The name with every invalid character removed.</returns>
    public static string SanitizeFileName(string name)
    {
        var builder = new StringBuilder(name.Length);
        foreach (var ch in name)
        {
            if (Array.IndexOf(InvalidFileNameChars, ch) < 0)
            {
                builder.Append(ch);
            }
        }

        return builder.ToString();
    }
}
