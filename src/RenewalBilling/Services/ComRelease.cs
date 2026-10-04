using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

namespace RenewalBilling.Services;

/// <summary>
/// Shared helpers for releasing COM objects deterministically, so Excel and Word never leave an
/// orphan process behind. Used by every service that automates Office.
/// </summary>
public static class ComRelease
{
    /// <summary>
    /// Releases every given COM object with <see cref="Marshal.FinalReleaseComObject"/>, in the
    /// reverse of the order they are passed in, then forces two rounds of garbage collection so
    /// no pending finalizer keeps Excel or Word alive. Pass objects in the order they were
    /// created; null entries are ignored.
    /// </summary>
    /// <param name="comObjectsInCreationOrder">The COM objects to release, in creation order.</param>
    public static void ReleaseAll(params object?[] comObjectsInCreationOrder)
    {
        for (var i = comObjectsInCreationOrder.Length - 1; i >= 0; i--)
        {
            var comObject = comObjectsInCreationOrder[i];
            if (comObject is not null && Marshal.IsComObject(comObject))
            {
                Marshal.FinalReleaseComObject(comObject);
            }
        }

        CollectTwice();
    }

    /// <summary>
    /// Forces two rounds of garbage collection and finalization. Office COM wrappers often need
    /// a second pass before the underlying RCW is actually released.
    /// </summary>
    public static void CollectTwice()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    /// <summary>
    /// Starts a new out-of-process COM server for the given ProgID (for example
    /// <c>Excel.Application</c> or <c>Word.Application</c>) and identifies the Windows process
    /// id it started, by diffing the named process list before and after activation. This avoids
    /// relying on a COM property for the window handle: Excel's <c>Application.Hwnd</c> works,
    /// but Word's <c>Application</c> object has no equivalent, so a single approach that works
    /// for both is simpler than two different ones.
    /// </summary>
    /// <param name="progId">The ProgID to activate, for example <c>Excel.Application</c>.</param>
    /// <param name="processName">The process name to look for, without ".exe", for example <c>EXCEL</c>.</param>
    /// <returns>The new COM object and, if exactly one new matching process appeared, its process id.</returns>
    public static (dynamic App, int? ProcessId) StartOfficeApplication(string progId, string processName)
    {
        var type = Type.GetTypeFromProgID(progId)
            ?? throw new COMException($"{progId} is not registered on this machine.");

        var before = Process.GetProcessesByName(processName).Select(p => p.Id).ToHashSet();

        dynamic app = Activator.CreateInstance(type)
            ?? throw new COMException($"Could not start a new {progId} instance.");

        var newProcessIds = Process.GetProcessesByName(processName)
            .Select(p => p.Id)
            .Where(id => !before.Contains(id))
            .ToList();

        return (app, newProcessIds.Count == 1 ? newProcessIds[0] : null);
    }

    /// <summary>
    /// Quits an Excel or Word application object and releases it, then confirms the underlying
    /// process actually exits, force-killing it after a few seconds if it has not. Late-bound
    /// Office automation easily leaves one stray COM reference somewhere in a long property
    /// chain, and a single outstanding reference is enough for Office to stay resident
    /// indefinitely after <c>Quit()</c>. This is the safety net that guarantees no orphan
    /// process survives a run regardless of any one leak slipping through.
    /// </summary>
    /// <param name="app">The Excel.Application or Word.Application COM object to quit.</param>
    /// <param name="processId">The process id captured by <see cref="StartOfficeApplication"/> when this app was created, or null if it could not be determined.</param>
    public static void QuitAndEnsureProcessExits(dynamic app, int? processId)
    {
        try
        {
            app.Quit();
        }
        catch (COMException)
        {
            // Best effort; we still release and, if needed, kill the process below.
        }

        Marshal.FinalReleaseComObject(app);
        CollectTwice();

        if (processId is null)
        {
            return;
        }

        for (var attempt = 0; attempt < 10; attempt++)
        {
            if (!TryGetProcess(processId.Value, out var process))
            {
                return;
            }

            using (process)
            {
                if (process.HasExited)
                {
                    return;
                }
            }

            Thread.Sleep(300);
        }

        if (TryGetProcess(processId.Value, out var finalProcess))
        {
            using (finalProcess)
            {
                if (!finalProcess.HasExited)
                {
                    finalProcess.Kill();
                }
            }
        }
    }

    private static bool TryGetProcess(int processId, out Process process)
    {
        try
        {
            process = Process.GetProcessById(processId);
            return true;
        }
        catch (ArgumentException)
        {
            process = null!;
            return false;
        }
    }
}
