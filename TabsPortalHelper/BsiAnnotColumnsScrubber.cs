// BsiAnnotColumnsScrubber.cs
//
// Strips the /BSIAnnotColumns key from a PDF's document Catalog.
//
// Background
// ----------
// Bluebeam Revu stores PDF-level custom column definitions in the document
// Catalog under the /BSIAnnotColumns key. Per Bluebeam's own docs, when a
// PDF carries its own column definitions they OVERRIDE the active Profile's
// columns in the Markups List. This is the root cause of the "my TABS
// columns disappeared" symptom users hit when opening PDFs marked up by
// other firms (engineering drawings from LAN, plan reviews from outside
// reviewers, corrective mods from owner's design firms, etc.).
//
// The fix
// -------
// Remove the /BSIAnnotColumns key. With it absent, Bluebeam falls back to
// the active Profile's column definitions — which is exactly what we want
// for the TABSportal profile. No other PDF content is touched: page content
// streams, annotations, signatures, metadata, bookmarks, and form fields
// are all preserved.
//
// Used by
// -------
//   - "Bluebeam ▸ Scrub PDF File…" tray menu item (closed-file scrub)
//   - Active-file fix orchestrator (close → scrub → reopen flow)
//   - /scrub-file and /scrub-active HTTP endpoints
//
// Idempotency
// -----------
// All operations are idempotent. Scrubbing a file that is already clean
// is a no-op (returns Status.AlreadyClean) and never touches the file.

using System;
using System.IO;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace TabsPortalHelper
{
    public static class BsiAnnotColumnsScrubber
    {
        /// <summary>
        /// The PDF Catalog key that holds Bluebeam's PDF-scoped custom
        /// column definitions. Presence of this key causes Revu to ignore
        /// the active Profile's columns in the Markups List for this file.
        /// </summary>
        public const string CatalogKey = "/BSIAnnotColumns";

        public enum Status
        {
            /// <summary>Removed /BSIAnnotColumns and saved the file.</summary>
            Scrubbed,

            /// <summary>File did not contain /BSIAnnotColumns. No write occurred.</summary>
            AlreadyClean,

            /// <summary>An error occurred while reading or writing.</summary>
            Failed,
        }

        public sealed class Result
        {
            public Status Status { get; init; }
            public string FilePath { get; init; } = "";
            public string? Message { get; init; }
            public Exception? Exception { get; init; }

            public bool Success => Status != Status.Failed;
        }

        /// <summary>
        /// Strips /BSIAnnotColumns from <paramref name="pdfPath"/> in place.
        ///
        /// Atomic: writes to a sibling .scrub.tmp file first, then
        /// File.Replace swaps it in. If anything fails partway, the
        /// original file is untouched.
        ///
        /// Idempotent: scrubbing a clean file returns AlreadyClean without
        /// writing.
        /// </summary>
        /// <param name="pdfPath">Absolute path to the PDF.</param>
        public static Result Scrub(string pdfPath)
        {
            if (string.IsNullOrWhiteSpace(pdfPath))
            {
                return new Result
                {
                    Status   = Status.Failed,
                    FilePath = pdfPath ?? "",
                    Message  = "PDF path is empty.",
                };
            }

            if (!File.Exists(pdfPath))
            {
                return new Result
                {
                    Status   = Status.Failed,
                    FilePath = pdfPath,
                    Message  = "File not found: " + pdfPath,
                };
            }

            string tmpPath = pdfPath + ".scrub.tmp";

            try
            {
                // Open in Modify mode so PdfSharp gives us a writable Catalog.
                using (var doc = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Modify))
                {
                    var catalog = doc.Internals.Catalog;

                    if (!catalog.Elements.ContainsKey(CatalogKey))
                    {
                        // Already clean. Nothing to write.
                        return new Result
                        {
                            Status   = Status.AlreadyClean,
                            FilePath = pdfPath,
                            Message  = "/BSIAnnotColumns not present.",
                        };
                    }

                    catalog.Elements.Remove(CatalogKey);

                    // Save to .tmp first for atomic replace.
                    doc.Save(tmpPath);
                } // doc disposed here, file handle on tmpPath released

                // Atomic replace on NTFS: original goes away, tmp becomes
                // original. We don't keep .bak files (clutters Drive folders).
                File.Replace(tmpPath, pdfPath, destinationBackupFileName: null);

                return new Result
                {
                    Status   = Status.Scrubbed,
                    FilePath = pdfPath,
                    Message  = "Removed /BSIAnnotColumns and saved.",
                };
            }
            catch (Exception ex)
            {
                // Best-effort cleanup of the .tmp file.
                try { if (File.Exists(tmpPath)) File.Delete(tmpPath); } catch { }

                return new Result
                {
                    Status    = Status.Failed,
                    FilePath  = pdfPath,
                    Message   = "Scrub failed: " + ex.Message,
                    Exception = ex,
                };
            }
        }

        /// <summary>
        /// Returns true if the PDF currently contains /BSIAnnotColumns
        /// (i.e. would be modified by Scrub). Read-only; does not modify
        /// the file.
        ///
        /// Used by ambient detection — e.g. checking the active Revu file
        /// before showing the "Fix it" toast on clipboard markup send.
        ///
        /// Returns false on any read failure (corrupt PDF, locked file,
        /// encrypted, etc.) — better to skip a toast than to spam one on
        /// a file we can't actually read.
        /// </summary>
        public static bool HasBsiAnnotColumns(string pdfPath)
        {
            if (string.IsNullOrWhiteSpace(pdfPath) || !File.Exists(pdfPath))
                return false;

            try
            {
                using var doc = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Import);
                return doc.Internals.Catalog.Elements.ContainsKey(CatalogKey);
            }
            catch
            {
                return false;
            }
        }
    }
}
