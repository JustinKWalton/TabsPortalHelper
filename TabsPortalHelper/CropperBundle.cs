// CropperBundle.cs
//
// Extracts the embedded Node.js + cropper bundle to %LOCALAPPDATA%\TabsPortalHelper\
// on install (and defensively on first /process-plan call).
//
// The bundle is a zip containing:
//   node\node.exe
//   cropper\cropper.js
//   cropper\package.json
//   cropper\node_modules\... (pdf-lib, pako, etc.)
//
// It's embedded as TabsPortalHelper.cropper-bundle.zip â€” see .csproj for the
// EmbeddedResource entry.
//
// Extraction is gated by a version marker file so we don't pay the ~5-second
// extraction cost on every launch. The marker is written after a successful
// extraction; if BundleVersion changes (next helper release ships a new
// node version or new cropper.js), the marker mismatches and extraction
// re-runs automatically.

using System;
using System.IO;
using System.IO.Compression;
using System.Reflection;

namespace TabsPortalHelper
{
    public static class CropperBundle
    {
        // Bump this every time the embedded bundle contents change.
        // The version marker file under TabsPortalHelper\ records the last-
        // extracted version; if this constant changes, ExtractIfNeeded
        // overwrites the on-disk bundle.
        public const string BundleVersion = "2.6.2";

        // Embedded resource name â€” must match the LogicalName in .csproj.
        private const string EmbeddedResourceName = "TabsPortalHelper.cropper-bundle.zip";

        // Marker file written after successful extraction.
        private const string VersionMarkerFile = "cropper-bundle.version";

        /// <summary>
        /// Root install directory: %LOCALAPPDATA%\TabsPortalHelper\
        /// </summary>
        public static string InstallRoot =>
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "TabsPortalHelper");

        /// <summary>
        /// Path to the bundled node.exe â€” used by CropperEndpoint to spawn the cropper.
        /// </summary>
        public static string NodeExePath =>
            Path.Combine(InstallRoot, "node", "node.exe");

        /// <summary>
        /// Path to cropper.js â€” used by CropperEndpoint to spawn the cropper.
        /// </summary>
        public static string CropperJsPath =>
            Path.Combine(InstallRoot, "cropper", "cropper.js");

        /// <summary>
        /// Returns true if both node.exe and cropper.js are on disk and the
        /// version marker matches BundleVersion. Used by CropperEndpoint to
        /// decide whether to call ExtractIfNeeded before running.
        /// </summary>
        public static bool IsAvailable()
        {
            if (!File.Exists(NodeExePath)) return false;
            if (!File.Exists(CropperJsPath)) return false;
            return ReadOnDiskVersion() == BundleVersion;
        }

        /// <summary>
        /// Idempotent. Extracts the embedded bundle if the version marker is
        /// missing or doesn't match. Safe to call on every /process-plan
        /// request â€” does nothing on the hot path once extraction has happened.
        /// </summary>
        /// <returns>true if extracted (or already present); false on failure.</returns>
        public static bool ExtractIfNeeded()
        {
            try
            {
                Directory.CreateDirectory(InstallRoot);

                // Fast path: already extracted, version matches.
                if (IsAvailable()) return true;

                // Extract from embedded resource into a temp dir, then swap into place.
                // Two-stage to avoid leaving the user with a half-extracted bundle if
                // we crash mid-extract.
                var tempDir = Path.Combine(InstallRoot, $".cropper-extract-{Guid.NewGuid():N}");
                Directory.CreateDirectory(tempDir);

                try
                {
                    using (var src = GetEmbeddedStream())
                    using (var zip = new ZipArchive(src, ZipArchiveMode.Read, leaveOpen: false))
                    {
                        zip.ExtractToDirectory(tempDir, overwriteFiles: true);
                    }

                    // Move node\ and cropper\ into final place. Each is a top-level
                    // folder inside the zip. Wipe any existing copies first so a
                    // version bump doesn't leave stale files behind.
                    foreach (var name in new[] { "node", "cropper" })
                    {
                        var src = Path.Combine(tempDir, name);
                        var dst = Path.Combine(InstallRoot, name);

                        if (!Directory.Exists(src))
                        {
                            throw new InvalidOperationException(
                                $"Bundle missing expected folder '{name}'.");
                        }

                        if (Directory.Exists(dst))
                        {
                            // Best-effort delete â€” if a file is in use (e.g. a previous
                            // node.exe still has a handle from a stuck cropper run),
                            // ExtractIfNeeded fails gracefully and the next call retries.
                            Directory.Delete(dst, recursive: true);
                        }

                        Directory.Move(src, dst);
                    }
                }
                finally
                {
                    try { Directory.Delete(tempDir, recursive: true); }
                    catch { /* best-effort cleanup */ }
                }

                // Stamp the version marker last â€” if anything above threw, the marker
                // stays stale and we'll retry on next call.
                File.WriteAllText(GetVersionMarkerPath(), BundleVersion);
                return true;
            }
            catch (Exception ex)
            {
                // Don't take down install or /process-plan over this â€” log and report
                // back to caller. Helper-cropping degrades gracefully: the edge function
                // falls back to placeholder thumbnails if the manifest isn't there.
                LastError = ex;
                return false;
            }
        }

        /// <summary>
        /// Captures the most recent extraction failure (if any). Useful for the
        /// /process-plan response when ExtractIfNeeded returns false.
        /// </summary>
        public static Exception? LastError { get; private set; }

        // â”€â”€â”€ helpers â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

        private static string GetVersionMarkerPath() =>
            Path.Combine(InstallRoot, VersionMarkerFile);

        private static string? ReadOnDiskVersion()
        {
            var path = GetVersionMarkerPath();
            if (!File.Exists(path)) return null;
            try { return File.ReadAllText(path).Trim(); }
            catch { return null; }
        }

        private static Stream GetEmbeddedStream()
        {
            var asm = Assembly.GetExecutingAssembly();
            var stream = asm.GetManifestResourceStream(EmbeddedResourceName)
                ?? throw new InvalidOperationException(
                    $"Embedded resource '{EmbeddedResourceName}' not found. " +
                    "Add cropper-bundle.zip to .csproj as <EmbeddedResource> with " +
                    $"LogicalName='{EmbeddedResourceName}'.");
            return stream;
        }
    }
}
