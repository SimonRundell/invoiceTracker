using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace RenewalBilling.Services;

/// <summary>
/// Where an <see cref="RpiResult"/> came from.
/// </summary>
public enum RpiSource
{
    /// <summary>Fetched live from ONS this run.</summary>
    Online,

    /// <summary>ONS could not be reached; the last successful fetch's cache file was used.</summary>
    Cache,

    /// <summary>ONS could not be reached and there was no cache; the configured fallback percentage was used.</summary>
    Fallback,
}

/// <summary>
/// The RPI percentage to use for this run, and where it came from.
/// </summary>
/// <param name="Percent">The RPI percentage.</param>
/// <param name="PeriodLabel">The period the percentage covers, for example <c>2026 AUG</c>.</param>
/// <param name="Source">Where this value came from.</param>
public sealed record RpiResult(decimal Percent, string PeriodLabel, RpiSource Source);

/// <summary>
/// Thrown when the ONS response cannot be parsed into an RPI value.
/// </summary>
public sealed class RpiParseException : Exception
{
    /// <summary>Creates a new <see cref="RpiParseException"/> with a message describing the problem.</summary>
    /// <param name="message">A clear, human readable description of what is wrong.</param>
    public RpiParseException(string message) : base(message)
    {
    }
}

/// <summary>
/// Fetches the latest RPI (all items, 12 month percentage change, ONS series CZBH, dataset MM23)
/// figure, with a cache and fallback for when ONS cannot be reached.
/// </summary>
public interface IRpiService
{
    /// <summary>
    /// Gets the latest RPI percentage: fetched online if possible, otherwise the last cached
    /// value, otherwise the configured fallback percentage.
    /// </summary>
    Task<RpiResult> GetLatestRpiAsync();
}

/// <summary>
/// One entry in the ONS "months" array. Only the fields this app uses are mapped; everything
/// else in the response is ignored by <see cref="JsonSerializer"/> automatically.
/// </summary>
internal sealed class OnsMonthEntry
{
    /// <summary>The period label, for example <c>2026 AUG</c>.</summary>
    [JsonPropertyName("date")]
    public string Date { get; set; } = string.Empty;

    /// <summary>The percentage change, as a string (ONS encodes every value as text).</summary>
    [JsonPropertyName("value")]
    public string Value { get; set; } = string.Empty;
}

/// <summary>
/// The parts of the ONS timeseries response this app reads.
/// </summary>
internal sealed class OnsResponse
{
    /// <summary>Every monthly data point, oldest first.</summary>
    [JsonPropertyName("months")]
    public List<OnsMonthEntry> Months { get; set; } = new();
}

/// <summary>
/// Default <see cref="IRpiService"/>: fetches the configured ONS URL with <see cref="HttpClient"/>,
/// caches the result to a local file, and falls back to that cache or a configured percentage
/// when ONS cannot be reached. The JSON parsing below is pure and unit tested directly; only
/// <see cref="GetLatestRpiAsync"/> itself touches the network or the file system.
/// </summary>
public sealed class RpiService : IRpiService
{
    private readonly string _url;
    private readonly int _timeoutSeconds;
    private readonly decimal _fallbackPercent;
    private readonly string _cacheFilePath;
    private readonly IAppLogger _logger;

    /// <summary>
    /// Creates an RPI service bound to one set of settings.
    /// </summary>
    /// <param name="config">The RPI section of the application config.</param>
    /// <param name="logger">Where to log the outcome of each fetch attempt.</param>
    public RpiService(RpiConfig config, IAppLogger logger)
    {
        _url = config.Url;
        _timeoutSeconds = config.TimeoutSeconds;
        _fallbackPercent = config.FallbackPercent;
        _cacheFilePath = config.CacheFile;
        _logger = logger;
    }

    /// <inheritdoc />
    public async Task<RpiResult> GetLatestRpiAsync()
    {
        try
        {
            using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(_timeoutSeconds) };

            // ONS returns 403 Forbidden to requests with no User-Agent header, which is
            // HttpClient's default. A plain, honest identifying string is enough to pass.
            client.DefaultRequestHeaders.UserAgent.ParseAdd("RenewalBilling/1.0");

            var json = await client.GetStringAsync(_url);
            var (percent, periodLabel) = ParseMonthsJson(json);

            WriteCache(percent, periodLabel);
            _logger.Info($"RPI fetched online: {percent}% ({periodLabel}).");
            return new RpiResult(percent, periodLabel, RpiSource.Online);
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or RpiParseException)
        {
            _logger.Warn($"RPI online fetch failed: {ex.Message}");
            return ReadCacheOrFallback();
        }
    }

    private RpiResult ReadCacheOrFallback()
    {
        var cached = TryReadCache();
        if (cached is not null)
        {
            _logger.Warn($"Using cached RPI: {cached.Percent}% ({cached.PeriodLabel}), cached {cached.FetchedAtUtc:yyyy-MM-dd HH:mm} UTC.");
            return new RpiResult(cached.Percent, cached.PeriodLabel, RpiSource.Cache);
        }

        _logger.Warn($"No RPI cache available. Using the configured fallback percentage: {_fallbackPercent}%.");
        return new RpiResult(_fallbackPercent, "Fallback", RpiSource.Fallback);
    }

    private RpiCacheData? TryReadCache()
    {
        if (!File.Exists(_cacheFilePath))
        {
            return null;
        }

        try
        {
            var json = File.ReadAllText(_cacheFilePath);
            return JsonSerializer.Deserialize<RpiCacheData>(json);
        }
        catch (Exception ex) when (ex is IOException or JsonException)
        {
            _logger.Warn($"RPI cache file could not be read: {ex.Message}");
            return null;
        }
    }

    private void WriteCache(decimal percent, string periodLabel)
    {
        try
        {
            var data = new RpiCacheData
            {
                Percent = percent,
                PeriodLabel = periodLabel,
                FetchedAtUtc = DateTime.UtcNow,
            };

            var folder = Path.GetDirectoryName(_cacheFilePath);
            if (!string.IsNullOrEmpty(folder))
            {
                Directory.CreateDirectory(folder);
            }

            File.WriteAllText(_cacheFilePath, JsonSerializer.Serialize(data));
        }
        catch (IOException ex)
        {
            _logger.Warn($"Could not write the RPI cache file: {ex.Message}");
        }
    }

    /// <summary>
    /// Parses an ONS timeseries response, returning the most recent entry in the "months" array
    /// that has a numeric value. Entries are scanned from the end backwards, since ONS can
    /// publish a placeholder entry for a period that has not been measured yet.
    /// </summary>
    /// <param name="json">The raw JSON body returned by the ONS endpoint.</param>
    /// <returns>The percentage and its period label, for example <c>2026 AUG</c>.</returns>
    /// <exception cref="RpiParseException">
    /// Thrown when the JSON cannot be parsed, has no "months" array, or no entry in it has a
    /// numeric value.
    /// </exception>
    public static (decimal Percent, string PeriodLabel) ParseMonthsJson(string json)
    {
        OnsResponse? response;
        try
        {
            response = JsonSerializer.Deserialize<OnsResponse>(json);
        }
        catch (JsonException ex)
        {
            throw new RpiParseException($"ONS response was not valid JSON: {ex.Message}");
        }

        if (response is null || response.Months.Count == 0)
        {
            throw new RpiParseException("ONS response did not contain a 'months' array.");
        }

        for (var i = response.Months.Count - 1; i >= 0; i--)
        {
            var entry = response.Months[i];
            if (decimal.TryParse(entry.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out var percent))
            {
                return (percent, entry.Date);
            }
        }

        throw new RpiParseException("ONS response's 'months' array had no entry with a numeric value.");
    }
}

/// <summary>
/// The shape of the local RPI cache file.
/// </summary>
internal sealed class RpiCacheData
{
    /// <summary>The cached RPI percentage.</summary>
    public decimal Percent { get; set; }

    /// <summary>The period the cached percentage covers, for example <c>2026 AUG</c>.</summary>
    public string PeriodLabel { get; set; } = string.Empty;

    /// <summary>When this value was fetched from ONS, in UTC.</summary>
    public DateTime FetchedAtUtc { get; set; }
}
