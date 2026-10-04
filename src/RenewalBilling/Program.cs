using RenewalBilling.Services;

namespace RenewalBilling;

/// <summary>
/// Process exit codes returned by the application.
/// </summary>
public enum ExitCode
{
    /// <summary>The run completed normally.</summary>
    Ok = 0,

    /// <summary>The config file or the client workbook could not be used.</summary>
    ConfigOrWorkbookError = 1,

    /// <summary>Excel or Word could not be automated.</summary>
    OfficeNotAvailable = 2,

    /// <summary>An unexpected failure stopped the run.</summary>
    UnexpectedFailure = 3,
}

/// <summary>
/// Which of the command line modes the run should operate in.
/// </summary>
public enum RunMode
{
    /// <summary>Normal run: find candidates and show the review form only if there are any.</summary>
    Scheduled,

    /// <summary>Log what would be invoiced. Shows nothing and writes nothing.</summary>
    DryRun,

    /// <summary>Create a sample workbook and sample Word template, then exit.</summary>
    MakeSamples,
}

/// <summary>
/// The parsed command line arguments for one run.
/// </summary>
public sealed class RunOptions
{
    /// <summary>Which mode the run should operate in.</summary>
    public RunMode Mode { get; init; } = RunMode.Scheduled;

    /// <summary>An override for "today", used for testing. Null means use the real date.</summary>
    public DateTime? DateOverride { get; init; }

    /// <summary>An override for the config file path. Null means use the default next to the exe.</summary>
    public string? ConfigPath { get; init; }
}

/// <summary>
/// Thrown when the command line arguments cannot be understood.
/// </summary>
public sealed class ArgsException : Exception
{
    /// <summary>Creates a new <see cref="ArgsException"/> with a message describing the problem.</summary>
    /// <param name="message">A clear, human readable description of what is wrong.</param>
    public ArgsException(string message) : base(message)
    {
    }
}

/// <summary>
/// Parses the command line arguments described in the project specification:
/// <c>/scheduled</c> (default), <c>/dryrun</c>, <c>/date yyyy-MM-dd</c>, <c>/config &lt;path&gt;</c>,
/// <c>/makesamples</c>.
/// </summary>
public static class RunOptionsParser
{
    /// <summary>
    /// Parses raw command line arguments into a <see cref="RunOptions"/>.
    /// </summary>
    /// <param name="args">The raw arguments passed to the process.</param>
    /// <returns>The parsed options.</returns>
    /// <exception cref="ArgsException">Thrown when an argument is unknown or malformed.</exception>
    public static RunOptions Parse(string[] args)
    {
        RunMode mode = RunMode.Scheduled;
        DateTime? dateOverride = null;
        string? configPath = null;
        bool modeSet = false;

        for (var i = 0; i < args.Length; i++)
        {
            var arg = args[i];

            switch (arg.ToLowerInvariant())
            {
                case "/scheduled":
                    mode = RunMode.Scheduled;
                    modeSet = true;
                    break;

                case "/dryrun":
                    mode = RunMode.DryRun;
                    modeSet = true;
                    break;

                case "/makesamples":
                    mode = RunMode.MakeSamples;
                    modeSet = true;
                    break;

                case "/date":
                    i++;
                    if (i >= args.Length)
                    {
                        throw new ArgsException("/date requires a value in the form yyyy-MM-dd.");
                    }

                    if (!DateTime.TryParseExact(args[i], "yyyy-MM-dd",
                            System.Globalization.CultureInfo.InvariantCulture,
                            System.Globalization.DateTimeStyles.None, out var parsedDate))
                    {
                        throw new ArgsException($"'{args[i]}' is not a valid date. Use the form yyyy-MM-dd.");
                    }

                    dateOverride = parsedDate;
                    break;

                case "/config":
                    i++;
                    if (i >= args.Length)
                    {
                        throw new ArgsException("/config requires a file path.");
                    }

                    configPath = args[i];
                    break;

                default:
                    throw new ArgsException($"Unrecognised argument '{arg}'.");
            }
        }

        _ = modeSet;

        return new RunOptions
        {
            Mode = mode,
            DateOverride = dateOverride,
            ConfigPath = configPath,
        };
    }
}

/// <summary>
/// Application entry point.
/// </summary>
static class Program
{
    /// <summary>
    /// The main entry point. Parses arguments, loads and validates config, sets up logging,
    /// then hands off to the rest of the application.
    /// </summary>
    /// <param name="args">Command line arguments. See <see cref="RunOptionsParser"/>.</param>
    /// <returns>An <see cref="ExitCode"/> value.</returns>
    [STAThread]
    static int Main(string[] args)
    {
        RunOptions options;
        try
        {
            options = RunOptionsParser.Parse(args);
        }
        catch (ArgsException ex)
        {
            Console.Error.WriteLine($"Argument error: {ex.Message}");
            return (int)ExitCode.ConfigOrWorkbookError;
        }

        var baseDirectory = AppContext.BaseDirectory;
        var configPath = ResolveConfigPath(options.ConfigPath, baseDirectory);

        AppConfig config;
        try
        {
            config = AppConfigLoader.Load(configPath, baseDirectory);
        }
        catch (ConfigException ex)
        {
            WriteStartupError(baseDirectory, ex.Message);
            return (int)ExitCode.ConfigOrWorkbookError;
        }

        var logger = new FileLogger(config.LogFolder);

        try
        {
            var runDate = options.DateOverride ?? DateTime.Today;
            logger.Info($"Run start. Mode={options.Mode} RunDate={runDate:dd MMM yyyy} Args=[{string.Join(' ', args)}]");

            switch (options.Mode)
            {
                case RunMode.MakeSamples:
                    logger.Info("/makesamples requested. Sample workbook and template generation will be added in a later milestone.");
                    break;

                case RunMode.DryRun:
                case RunMode.Scheduled:
                default:
                    logger.Info("Milestone 1 scaffold only: eligibility checks, RPI lookup, and invoice generation are not implemented yet.");
                    break;
            }

            logger.Info("Run end.");
            return (int)ExitCode.Ok;
        }
        catch (Exception ex)
        {
            logger.Error($"Unexpected failure: {ex}");
            return (int)ExitCode.UnexpectedFailure;
        }
    }

    private static string ResolveConfigPath(string? configPathArg, string baseDirectory)
    {
        var path = string.IsNullOrWhiteSpace(configPathArg) ? ".config.json" : configPathArg;
        return Path.IsPathRooted(path) ? path : Path.GetFullPath(Path.Combine(baseDirectory, path));
    }

    private static void WriteStartupError(string baseDirectory, string message)
    {
        var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} ERROR {message}";
        Console.Error.WriteLine(line);

        try
        {
            var fallbackFolder = Path.Combine(baseDirectory, "Logs");
            Directory.CreateDirectory(fallbackFolder);
            var fallbackFile = Path.Combine(fallbackFolder, $"RenewalBilling_{DateTime.Now:yyyyMMdd}.log");
            File.AppendAllLines(fallbackFile, new[] { line });
        }
        catch (IOException)
        {
            // The config is already broken; if we cannot even write a fallback log there is
            // nothing more we can tell the user beyond the console line written above.
        }
    }
}
