namespace RenewalBilling.Services;

/// <summary>
/// Severity of a logged message.
/// </summary>
public enum LogLevel
{
    /// <summary>Routine information about normal progress.</summary>
    Info,

    /// <summary>Something unexpected that the run can still continue past.</summary>
    Warn,

    /// <summary>A failure that stopped a run, or stopped processing for one client.</summary>
    Error,
}

/// <summary>
/// Writes time stamped log lines to a plain text file.
/// </summary>
public interface IAppLogger
{
    /// <summary>Logs a routine informational message.</summary>
    /// <param name="message">The message to write.</param>
    void Info(string message);

    /// <summary>Logs a warning that does not stop the run.</summary>
    /// <param name="message">The message to write.</param>
    void Warn(string message);

    /// <summary>Logs an error.</summary>
    /// <param name="message">The message to write.</param>
    void Error(string message);
}

/// <summary>
/// An <see cref="IAppLogger"/> that appends to one file per calendar day in a configured folder.
/// File names follow the pattern <c>RenewalBilling_yyyyMMdd.log</c>.
/// </summary>
public sealed class FileLogger : IAppLogger
{
    private readonly string _logFolder;
    private readonly object _writeLock = new();

    /// <summary>
    /// Creates a logger that writes into <paramref name="logFolder"/>, creating the folder if needed.
    /// </summary>
    /// <param name="logFolder">Full path to the folder that log files are written into.</param>
    public FileLogger(string logFolder)
    {
        _logFolder = logFolder;
        Directory.CreateDirectory(_logFolder);
    }

    /// <inheritdoc />
    public void Info(string message) => Write(LogLevel.Info, message);

    /// <inheritdoc />
    public void Warn(string message) => Write(LogLevel.Warn, message);

    /// <inheritdoc />
    public void Error(string message) => Write(LogLevel.Error, message);

    private void Write(LogLevel level, string message)
    {
        var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {level.ToString().ToUpperInvariant()} {message}";
        var filePath = Path.Combine(_logFolder, $"RenewalBilling_{DateTime.Now:yyyyMMdd}.log");

        lock (_writeLock)
        {
            File.AppendAllLines(filePath, new[] { line });
        }
    }
}
