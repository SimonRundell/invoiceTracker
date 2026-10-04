using RenewalBilling;
using RenewalBilling.Services;

namespace RenewalBilling.Tests;

public class ExcelRepositoryTests
{
    private static ClientColumnsConfig ClientColumns() => new()
    {
        Id = "ClientId", Name = "ClientName", Contact = "ContactName",
        Address1 = "Address1", Address2 = "Address2", Town = "Town",
        Postcode = "Postcode", Email = "Email",
        DueDate = "DueDate", LastPaidDate = "LastPaidDate", Active = "Active",
    };

    private static ItemColumnsConfig ItemColumns() => new()
    {
        ClientId = "ClientId", Description = "Description", Quantity = "Quantity",
        UnitCost = "UnitCost", PreviousUnitCost = "PreviousUnitCost", LastUpliftPercent = "LastUpliftPercent",
    };

    private static InvoiceColumnsConfig InvoiceColumns() => new()
    {
        InvoiceNumber = "InvoiceNumber", ClientId = "ClientId", InvoiceDate = "InvoiceDate",
        DueDate = "DueDate", PeriodStart = "PeriodStart", PeriodEnd = "PeriodEnd",
        Net = "Net", Vat = "Vat", Total = "Total", RpiPercent = "RpiPercent",
        DocxPath = "DocxPath", PdfPath = "PdfPath",
    };

    // Builds a 1-based object[,] the way Excel's Range.Value2 returns one: row 1 is headers,
    // row 2.. is data, both dimensions starting at index 1.
    private static object[,] BuildSheet(string[] headers, params object?[][] rows)
    {
        var array = (object[,])Array.CreateInstance(typeof(object), new[] { rows.Length + 1, headers.Length }, new[] { 1, 1 });

        for (var c = 0; c < headers.Length; c++)
        {
            array[1, c + 1] = headers[c];
        }

        for (var r = 0; r < rows.Length; r++)
        {
            for (var c = 0; c < headers.Length; c++)
            {
                array[r + 2, c + 1] = rows[r][c] ?? string.Empty;
            }
        }

        return array;
    }

    [Fact]
    public void BuildHeaderMap_IsCaseInsensitiveAndIgnoresBlankHeaders()
    {
        var data = BuildSheet(["ClientId", "clientname", "", "DueDate"]);

        var map = ExcelRepository.BuildHeaderMap(data);

        Assert.Equal(1, map["ClientId"]);
        Assert.Equal(2, map["CLIENTNAME"]); // case-insensitive match
        Assert.Equal(4, map["DueDate"]);
        Assert.Equal(3, map.Count); // the blank header at column 3 is not mapped
    }

    [Fact]
    public void ValidateHeaders_MissingColumn_ThrowsNamingIt()
    {
        var headerMap = new Dictionary<string, int> { ["ClientId"] = 1 };

        var ex = Assert.Throws<WorkbookException>(() =>
            ExcelRepository.ValidateHeaders(headerMap, ["ClientId", "DueDate", "Active"], "Clients"));

        Assert.Contains("DueDate", ex.Message);
        Assert.Contains("Active", ex.Message);
        Assert.Contains("Clients", ex.Message);
    }

    [Fact]
    public void ValidateHeaders_AllPresent_DoesNotThrow()
    {
        var headerMap = new Dictionary<string, int> { ["ClientId"] = 1, ["DueDate"] = 2 };

        ExcelRepository.ValidateHeaders(headerMap, ["ClientId", "DueDate"], "Clients");
    }

    [Fact]
    public void ParseClients_SkipsFullyBlankRowsAndTrimsStrings()
    {
        var headers = new[] { "ClientId", "ClientName", "ContactName", "Address1", "Address2", "Town", "Postcode", "Email", "DueDate", "LastPaidDate", "Active" };
        var data = BuildSheet(headers,
            ["C1", "  Acme Ltd  ", "Jo", "1 St", "", "Town", "AB1 2CD", "jo@test", 45900.0, "", "Y"],
            ["", "", "", "", "", "", "", "", "", "", ""]); // fully blank row must be skipped

        var headerMap = ExcelRepository.BuildHeaderMap(data);
        var result = ExcelRepository.ParseClients(data, headerMap, ClientColumns());

        Assert.Single(result);
        Assert.Equal("Acme Ltd", result[0].Client.ClientName); // trimmed
        Assert.Equal(2, result[0].RowNumber);
    }

    [Fact]
    public void ParseClients_OleDateDueDate_IsConvertedFromDouble()
    {
        var headers = new[] { "ClientId", "ClientName", "ContactName", "Address1", "Address2", "Town", "Postcode", "Email", "DueDate", "LastPaidDate", "Active" };
        var oleDate = new DateTime(2026, 10, 20).ToOADate();
        var data = BuildSheet(headers, ["C1", "Acme", "", "", "", "", "", "", oleDate, "", "Y"]);

        var headerMap = ExcelRepository.BuildHeaderMap(data);
        var result = ExcelRepository.ParseClients(data, headerMap, ClientColumns());

        Assert.Equal(new DateTime(2026, 10, 20), result[0].Client.DueDate);
    }

    [Fact]
    public void ParseClients_StringDueDate_IsParsedTolerantly()
    {
        var headers = new[] { "ClientId", "ClientName", "ContactName", "Address1", "Address2", "Town", "Postcode", "Email", "DueDate", "LastPaidDate", "Active" };
        var data = BuildSheet(headers, ["C1", "Acme", "", "", "", "", "", "", "20/10/2026", "", "Y"]);

        var headerMap = ExcelRepository.BuildHeaderMap(data);
        var result = ExcelRepository.ParseClients(data, headerMap, ClientColumns());

        Assert.Equal(new DateTime(2026, 10, 20), result[0].Client.DueDate);
    }

    [Fact]
    public void ParseClients_BlankDueDate_IsNull()
    {
        var headers = new[] { "ClientId", "ClientName", "ContactName", "Address1", "Address2", "Town", "Postcode", "Email", "DueDate", "LastPaidDate", "Active" };
        var data = BuildSheet(headers, ["C1", "Acme", "", "", "", "", "", "", "", "", "Y"]);

        var headerMap = ExcelRepository.BuildHeaderMap(data);
        var result = ExcelRepository.ParseClients(data, headerMap, ClientColumns());

        Assert.Null(result[0].Client.DueDate);
    }

    [Theory]
    [InlineData("Y", true)]
    [InlineData("y", true)]
    [InlineData("", true)]
    [InlineData("N", false)]
    [InlineData("n", false)]
    [InlineData("No", false)]
    public void ParseClients_ActiveFlag_HandlesYNAndBlank(string active, bool expected)
    {
        var headers = new[] { "ClientId", "ClientName", "ContactName", "Address1", "Address2", "Town", "Postcode", "Email", "DueDate", "LastPaidDate", "Active" };
        var data = BuildSheet(headers, ["C1", "Acme", "", "", "", "", "", "", "", "", active]);

        var headerMap = ExcelRepository.BuildHeaderMap(data);
        var result = ExcelRepository.ParseClients(data, headerMap, ClientColumns());

        Assert.Equal(expected, result[0].Client.Active);
    }

    [Fact]
    public void ParseItems_MixedNumberFormats_ParsesDoubleAndStringUnitCost()
    {
        var headers = new[] { "ClientId", "Description", "Quantity", "UnitCost", "PreviousUnitCost", "LastUpliftPercent" };
        var data = BuildSheet(headers,
            ["C1", "Licence", 1.0, 100.0, 95.0, 5.26],      // numeric (double) cells, as Excel normally stores them
            ["C2", "Support", "2", "50.5", "", ""]);         // the same column read back as text, with blanks

        var headerMap = ExcelRepository.BuildHeaderMap(data);
        var result = ExcelRepository.ParseItems(data, headerMap, ItemColumns());

        Assert.Equal(2, result.Count);
        Assert.Equal(100.0m, result[0].Item.UnitCost);
        Assert.Equal(95.0m, result[0].Item.PreviousUnitCost);
        Assert.Equal(2m, result[1].Item.Quantity);
        Assert.Equal(50.5m, result[1].Item.UnitCost);
        Assert.Null(result[1].Item.PreviousUnitCost);
    }

    [Fact]
    public void ParseItems_SkipsBlankRowAndRowWithNoClientId()
    {
        var headers = new[] { "ClientId", "Description", "Quantity", "UnitCost", "PreviousUnitCost", "LastUpliftPercent" };
        var data = BuildSheet(headers,
            ["", "", "", "", "", ""],
            ["", "Orphan line with no client id", 1.0, 10.0, "", ""],
            ["C1", "Licence", 1.0, 100.0, "", ""]);

        var headerMap = ExcelRepository.BuildHeaderMap(data);
        var result = ExcelRepository.ParseItems(data, headerMap, ItemColumns());

        Assert.Single(result);
        Assert.Equal("C1", result[0].Item.ClientId);
    }

    [Fact]
    public void ParseInvoices_ReadsEveryColumn()
    {
        var headers = new[] { "InvoiceNumber", "ClientId", "InvoiceDate", "DueDate", "PeriodStart", "PeriodEnd", "Net", "Vat", "Total", "RpiPercent", "DocxPath", "PdfPath" };
        var invoiceDate = new DateTime(2026, 10, 4).ToOADate();
        var periodStart = new DateTime(2026, 10, 20).ToOADate();
        var periodEnd = new DateTime(2027, 10, 19).ToOADate();
        var data = BuildSheet(headers,
            ["INV-2026-0001", "C3", invoiceDate, periodStart, periodStart, periodEnd, 324.0, 0.0, 324.0, 3.5, @"Invoices\2026\x.docx", @"Invoices\2026\x.pdf"]);

        var headerMap = ExcelRepository.BuildHeaderMap(data);
        var result = ExcelRepository.ParseInvoices(data, headerMap, InvoiceColumns());

        Assert.Single(result);
        var record = result[0];
        Assert.Equal("INV-2026-0001", record.InvoiceNumber);
        Assert.Equal("C3", record.ClientId);
        Assert.Equal(new DateTime(2026, 10, 20), record.PeriodStart);
        Assert.Equal(new DateTime(2027, 10, 19), record.PeriodEnd);
        Assert.Equal(324.0m, record.Net);
        Assert.Equal(3.5m, record.RpiPercent);
    }

    [Fact]
    public void ParseInvoices_SkipsRowsWithNoInvoiceNumber()
    {
        var headers = new[] { "InvoiceNumber", "ClientId", "InvoiceDate", "DueDate", "PeriodStart", "PeriodEnd", "Net", "Vat", "Total", "RpiPercent", "DocxPath", "PdfPath" };
        var data = BuildSheet(headers, ["", "C1", "", "", "", "", "", "", "", "", "", ""]);

        var headerMap = ExcelRepository.BuildHeaderMap(data);
        var result = ExcelRepository.ParseInvoices(data, headerMap, InvoiceColumns());

        Assert.Empty(result);
    }
}
