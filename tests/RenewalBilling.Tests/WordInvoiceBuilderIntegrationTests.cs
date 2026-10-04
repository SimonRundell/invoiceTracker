using System.IO.Compression;
using RenewalBilling.Models;
using RenewalBilling.Services;

namespace RenewalBilling.Tests;

/// <summary>
/// Exercises <see cref="WordInvoiceBuilder"/> against a real, locally installed Word instance.
/// Unlike the rest of the test project, these are integration tests, not pure unit tests: they
/// require Microsoft Word to be installed, which this project already assumes (see section 1 of
/// the specification: "COM (Excel and Word installed locally)").
/// </summary>
public class WordInvoiceBuilderIntegrationTests : IDisposable
{
    private readonly string _workDir;

    public WordInvoiceBuilderIntegrationTests()
    {
        _workDir = Path.Combine(Path.GetTempPath(), "RenewalBillingWordTests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_workDir);
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(_workDir, recursive: true);
        }
        catch (IOException)
        {
        }
    }

    private static InvoiceCandidate MakeCandidate(string invoiceNumber, IReadOnlyList<ProposedLineItem> items)
    {
        var client = new Client
        {
            ClientId = "C1",
            ClientName = "Acme Ltd",
            ContactName = "Jo Bloggs",
            Address1 = "1 High Street",
            Address2 = "Suite 2",
            Town = "Anytown",
            Postcode = "AB1 2CD",
            Email = "jo@acme.test",
            DueDate = new DateTime(2026, 10, 20),
        };

        return new InvoiceCandidate
        {
            Client = client,
            Status = CandidateStatus.Due,
            Items = items,
            RpiPercentUsed = 3.4m,
            InvoiceDate = new DateTime(2026, 10, 4),
            PeriodStart = new DateTime(2026, 10, 20),
            PeriodEnd = new DateTime(2027, 10, 19),
            PayByDate = new DateTime(2026, 10, 20),
            Vat = 0m,
            InvoiceNumber = invoiceNumber,
        };
    }

    [Fact]
    public void CreateInvoice_TwoLineItems_FillsTemplateAndSavesDocxAndPdf()
    {
        var templatePath = Path.Combine(_workDir, "Template.dotx");
        var invoiceFolder = Path.Combine(_workDir, "Invoices");
        var logger = new FileLogger(Path.Combine(_workDir, "Logs"));

        using var builder = new WordInvoiceBuilder(logger);
        builder.CreateSampleTemplate(templatePath);

        var items = new List<ProposedLineItem>
        {
            new(new BillingItem { ClientId = "C1", Description = "Annual software licence", Quantity = 1, UnitCost = 250m }, 258m, 3.4m),
            new(new BillingItem { ClientId = "C1", Description = "Support contract", Quantity = 1, UnitCost = 500m }, 517m, 3.4m),
        };

        var candidate = MakeCandidate("INV-2026-0001", items);

        var result = builder.CreateInvoice(candidate, templatePath, invoiceFolder);

        Assert.True(File.Exists(result.DocxPath), $"Expected docx at {result.DocxPath}");
        Assert.True(File.Exists(result.PdfPath), $"Expected pdf at {result.PdfPath}");
        Assert.Empty(result.UnknownTags);

        var bodyText = ReadDocxMainText(result.DocxPath);
        Assert.Contains("Acme Ltd", bodyText);
        Assert.Contains("Jo Bloggs", bodyText);
        Assert.Contains("INV-2026-0001", bodyText);
        Assert.Contains("Annual software licence", bodyText);
        Assert.Contains("Support contract", bodyText);
        Assert.Contains("£258.00", bodyText);
        Assert.Contains("£517.00", bodyText);
        Assert.Contains("£775.00", bodyText); // net = 258 + 517
        Assert.DoesNotContain("##", bodyText);
    }

    [Fact]
    public void CreateInvoice_FileAlreadyExists_ThrowsAndDoesNotOverwrite()
    {
        var templatePath = Path.Combine(_workDir, "Template.dotx");
        var invoiceFolder = Path.Combine(_workDir, "Invoices");
        var logger = new FileLogger(Path.Combine(_workDir, "Logs"));

        using var builder = new WordInvoiceBuilder(logger);
        builder.CreateSampleTemplate(templatePath);

        var items = new List<ProposedLineItem>
        {
            new(new BillingItem { ClientId = "C1", Description = "Annual software licence", Quantity = 1, UnitCost = 250m }, 258m, 3.4m),
        };

        var candidate = MakeCandidate("INV-2026-0002", items);
        var first = builder.CreateInvoice(candidate, templatePath, invoiceFolder);
        var originalLength = new FileInfo(first.DocxPath).Length;

        Assert.Throws<InvoiceBuildException>(() => builder.CreateInvoice(candidate, templatePath, invoiceFolder));

        Assert.Equal(originalLength, new FileInfo(first.DocxPath).Length);
    }

    private static string ReadDocxMainText(string docxPath)
    {
        using var archive = ZipFile.OpenRead(docxPath);
        var entry = archive.GetEntry("word/document.xml") ?? throw new InvalidOperationException("document.xml not found");
        using var stream = entry.Open();
        using var reader = new StreamReader(stream);
        var xml = reader.ReadToEnd();

        // Crude but sufficient for assertions: strip tags, leaving the visible text behind.
        return System.Text.RegularExpressions.Regex.Replace(xml, "<[^>]+>", " ");
    }
}
