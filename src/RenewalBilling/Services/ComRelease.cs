using System.Runtime.InteropServices;

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
}
