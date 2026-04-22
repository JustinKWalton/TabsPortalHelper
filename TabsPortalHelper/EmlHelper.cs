using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace TabsPortalHelper
{
    // ════════════════════════════════════════════════════════════════════════
    // EML Compose Helper
    // Builds a proper MIME .eml file and opens it in the default mail client.
    //
    // Advantages over Simple MAPI:
    //   - CC and BCC recipients work correctly in all mail clients
    //   - HTML body with proper line breaks
    //   - File attachments as MIME parts
    //   - Works with Mailbird, Outlook, Thunderbird, Windows Mail
    // ════════════════════════════════════════════════════════════════════════

    static class EmlHelper
    {
        public class ComposeRequest
        {
            public List<string> To        { get; set; } = new();
            public List<string> Cc        { get; set; } = new();
            public List<string> Bcc       { get; set; } = new();
            public string       Subject   { get; set; } = "";
            public string       Body      { get; set; } = "";
            public List<string> FilePaths { get; set; } = new();
        }

        public class ComposeResult
        {
            public bool    Success       { get; set; }
            public string? Error         { get; set; }
            public string? EmlPath       { get; set; }
        }

        public static ComposeResult Compose(ComposeRequest request)
        {
            string? emlPath = null;

            try
            {
                // ── Build MIME message ───────────────────────────────────────
                var boundary = $"----=_Part_{Guid.NewGuid():N}";
                var sb = new StringBuilder();

                // ── Headers ──────────────────────────────────────────────────
                sb.AppendLine("MIME-Version: 1.0");
                sb.AppendLine($"Date: {DateTime.UtcNow:R}");
                sb.AppendLine($"Subject: {EncodeMimeHeader(request.Subject)}");

                if (request.To.Count > 0)
                    sb.AppendLine($"To: {string.Join(", ", request.To)}");

                if (request.Cc.Count > 0)
                    sb.AppendLine($"Cc: {string.Join(", ", request.Cc)}");

                if (request.Bcc.Count > 0)
                    sb.AppendLine($"Bcc: {string.Join(", ", request.Bcc)}");

                if (request.FilePaths.Count > 0)
                {
                    // Multipart/mixed for body + attachments
                    sb.AppendLine($"Content-Type: multipart/mixed; boundary=\"{boundary}\"");
                    sb.AppendLine();
                    sb.AppendLine($"--{boundary}");
                }

                // ── HTML body part ───────────────────────────────────────────
                // Convert plain text newlines to HTML <br> tags
                var htmlBody = request.Body
                    .Replace("\r\n", "\n")
                    .Replace("\n", "<br>\r\n");

                var fullHtml = $"<html><body><p style=\"font-family:Arial,sans-serif;font-size:10pt;\">{htmlBody}</p></body></html>";

                if (request.FilePaths.Count > 0)
                {
                    sb.AppendLine("Content-Type: text/html; charset=utf-8");
                    sb.AppendLine("Content-Transfer-Encoding: base64");
                    sb.AppendLine();
                    sb.AppendLine(Convert.ToBase64String(Encoding.UTF8.GetBytes(fullHtml)));
                    sb.AppendLine();
                }
                else
                {
                    // No attachments — simple single-part HTML message
                    sb.AppendLine("Content-Type: text/html; charset=utf-8");
                    sb.AppendLine("Content-Transfer-Encoding: base64");
                    sb.AppendLine();
                    sb.AppendLine(Convert.ToBase64String(Encoding.UTF8.GetBytes(fullHtml)));
                }

                // ── Attachment parts ─────────────────────────────────────────
                foreach (var filePath in request.FilePaths)
                {
                    if (!File.Exists(filePath)) continue;

                    var fileName = Path.GetFileName(filePath);
                    var fileBytes = File.ReadAllBytes(filePath);
                    var contentType = GetContentType(filePath);
                    var base64 = Convert.ToBase64String(fileBytes, Base64FormattingOptions.InsertLineBreaks);

                    sb.AppendLine($"--{boundary}");
                    sb.AppendLine($"Content-Type: {contentType}; name=\"{fileName}\"");
                    sb.AppendLine("Content-Transfer-Encoding: base64");
                    sb.AppendLine($"Content-Disposition: attachment; filename=\"{fileName}\"");
                    sb.AppendLine();
                    sb.AppendLine(base64);
                    sb.AppendLine();
                }

                // ── Close boundary if multipart ──────────────────────────────
                if (request.FilePaths.Count > 0)
                    sb.AppendLine($"--{boundary}--");

                // ── Write .eml to temp file ──────────────────────────────────
                emlPath = Path.Combine(Path.GetTempPath(), $"TABS_{Guid.NewGuid():N}.eml");
                File.WriteAllText(emlPath, sb.ToString(), Encoding.UTF8);

                // ── Open in default mail client ──────────────────────────────
                Process.Start(new ProcessStartInfo(emlPath) { UseShellExecute = true });

                // ── Schedule temp file cleanup after 60 seconds ──────────────
                var pathToDelete = emlPath;
                System.Threading.Tasks.Task.Delay(60_000).ContinueWith(_ =>
                {
                    try { if (File.Exists(pathToDelete)) File.Delete(pathToDelete); }
                    catch { /* ignore cleanup errors */ }
                });

                return new ComposeResult { Success = true, EmlPath = emlPath };
            }
            catch (Exception ex)
            {
                // Clean up temp file if something went wrong
                try { if (emlPath != null && File.Exists(emlPath)) File.Delete(emlPath); }
                catch { /* ignore */ }

                return new ComposeResult { Success = false, Error = ex.Message };
            }
        }

        // ── Encode subject header for non-ASCII characters ───────────────────
        static string EncodeMimeHeader(string value)
        {
            // Check if encoding is needed
            foreach (char c in value)
                if (c > 127)
                    return $"=?utf-8?B?{Convert.ToBase64String(Encoding.UTF8.GetBytes(value))}?=";
            return value;
        }

        static string GetContentType(string filePath) =>
            Path.GetExtension(filePath).ToLowerInvariant() switch
            {
                ".pdf"  => "application/pdf",
                ".png"  => "image/png",
                ".jpg"  => "image/jpeg",
                ".jpeg" => "image/jpeg",
                ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                ".txt"  => "text/plain",
                _       => "application/octet-stream"
            };
    }
}
