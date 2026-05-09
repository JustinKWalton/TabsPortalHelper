// ActiveFileTracker.cs
//
// Records the most recent PDF path the user worked with through a
// TABSportal flow (currently: /open). Used by BluebeamColumnsFixOrchestrator
// to know which file to scrub when the user clicks "Fix Active File
// Columns..." without having to parse Bluebeam's window title.
//
// Why this approach (vs. parsing Revu's MainWindowTitle):
//   * Title parsing only gives a filename, not a full path. Resolving back
//     to a path requires a Drive search and is fragile when the same name
//     exists in multiple project folders.
//   * Title strings vary across Revu versions (v21 vs. v2019).
//   * The user's intent is to fix the file they JUST opened from TABSportal
//     -- which is exactly what /open tells us. Trust the most recent signal.
//
// Thread safety:
//   HttpServer dispatches requests on multiple threads. Writes (from
//   HandleOpen) and reads (from the orchestrator on the tray UI thread or
//   from HandleScrubActive) can race. Single object lock around two fields
//   is plenty.

using System;

namespace TabsPortalHelper
{
    public static class ActiveFileTracker
    {
        private static readonly object _lock = new();
        private static string? _lastFilePath;
        private static DateTime _lastUpdatedUtc;

        /// <summary>
        /// Records that this file path was most recently the focus of a
        /// TABSportal flow. Idempotent. Safe to call from any thread.
        /// Whitespace / null inputs are silently ignored so callers don't
        /// have to guard.
        /// </summary>
        public static void SetLastFile(string? filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath)) return;
            lock (_lock)
            {
                _lastFilePath  = filePath;
                _lastUpdatedUtc = DateTime.UtcNow;
            }
        }

        /// <summary>
        /// Returns the most recently tracked file path and when it was
        /// recorded, or (null, default) if no file has been tracked yet.
        /// The returned path is NOT guaranteed to still exist on disk;
        /// callers must validate before use.
        /// </summary>
        public static (string? FilePath, DateTime UtcWhen) GetLastFile()
        {
            lock (_lock)
            {
                return (_lastFilePath, _lastUpdatedUtc);
            }
        }

        /// <summary>
        /// Forgets the tracked file. Useful in tests, or if a caller wants
        /// to force the next "fix active file" operation to prompt the user
        /// for a file path.
        /// </summary>
        public static void Clear()
        {
            lock (_lock)
            {
                _lastFilePath   = null;
                _lastUpdatedUtc = default;
            }
        }
    }
}