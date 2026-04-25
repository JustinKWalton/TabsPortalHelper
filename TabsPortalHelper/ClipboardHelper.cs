using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace TabsPortalHelper
{
    /// <summary>
    /// Writes a Bluebeam markup to the Windows clipboard under the custom
    /// format name "Bluebeam.Windows.View.Input.BBCopyItem", so Ctrl+V
    /// inside Bluebeam Revu pastes a live, editable annotation.
    ///
    /// BSI column mapping (MUST match the TABSportal.bpx profile's UserDefined column ordering
    /// and the generate-markup-summary edge function's extractMarkups):
    ///     BSI[0] = Note            (the user's free-form note)
    ///     BSI[1] = Comment Status
    ///     BSI[2] = Element         (was: ItemCategory / "Subject" pre-v2.4)
    ///     BSI[3] = Location
    ///     BSI[4] = Issue           (NEW in v2.4 — was: Element pre-v2.4)
    /// DO NOT reorder these without updating both TABSportal.bpx
    /// and the edge function's bsiStrs[0..4] mapping.
    ///
    /// /Subj (PDF built-in Subject field) is no longer written. The Subject
    /// column was removed from the TABSportal profile in v2.3 and the helper
    /// stopped populating it in v2.4.
    /// </summary>
    static class ClipboardHelper
    {
        // ── Win32 P/Invoke ───────────────────────────────────────────────────
        const uint GMEM_MOVEABLE = 0x0002;

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool CloseClipboard();

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool EmptyClipboard();

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr SetClipboardData(uint uFormat, IntPtr hMem);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        static extern uint RegisterClipboardFormatW(string lpszFormat);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern IntPtr GlobalAlloc(uint uFlags, UIntPtr dwBytes);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern IntPtr GlobalLock(IntPtr hMem);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool GlobalUnlock(IntPtr hMem);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern IntPtr GlobalFree(IntPtr hMem);

        const string BluebeamClipboardFormat = "Bluebeam.Windows.View.Input.BBCopyItem";

        // ── Public API ──────────────────────────────────────────────────────

        public class MarkupRequest
        {
            public string? Note     { get; set; }   // → BSI[0] — user's free-form note
            public string? Status   { get; set; }   // → BSI[1] — Comment Status
            public string? Element  { get; set; }   // → BSI[2] (was: ItemCategory)
            public string? Location { get; set; }   // → BSI[3]
            public string? Issue    { get; set; }   // → BSI[4] — NEW in v2.4
            public string? Contents { get; set; }   // plain text — /Contents + /RC body
            public string? RcHtml   { get; set; }   // optional XHTML; auto-generated from Contents if null
        }

        public class MarkupResult
        {
            public bool    Success          { get; set; }
            public int     BytesWritten     { get; set; }
            public int     DataStringLength { get; set; }
            public string? Error            { get; set; }
        }

        public static MarkupResult PutMarkupOnClipboard(MarkupRequest req)
        {
            try
            {
                byte[] templateBytes = LoadTemplate();
                byte[] modifiedBytes = BuildModifiedBlob(templateBytes, req, out int dataStringLen);
                // Clipboard requires an STA thread. We create one per call so
                // this works even from thread-pool contexts like HttpListener.
                Exception? err = null;
                var t = new Thread(() =>
                {
                    try { WriteToClipboard(modifiedBytes); }
                    catch (Exception ex) { err = ex; }
                });
                t.SetApartmentState(ApartmentState.STA);
                t.Start();
                t.Join();
                if (err != null) throw err;

                return new MarkupResult
                {
                    Success          = true,
                    BytesWritten     = modifiedBytes.Length,
                    DataStringLength = dataStringLen,
                };
            }
            catch (Exception ex)
            {
                return new MarkupResult { Success = false, Error = ex.Message };
            }
        }

        // ── Template loading ────────────────────────────────────────────────

        static byte[] LoadTemplate()
        {
            const string resourceName = "TabsPortalHelper.template.bin";
            var asm = Assembly.GetExecutingAssembly();
            using var s = asm.GetManifestResourceStream(resourceName)
                ?? throw new InvalidOperationException(
                    $"Embedded resource '{resourceName}' not found. Add template.bin to the .csproj as <EmbeddedResource> with LogicalName set to '{resourceName}'.");
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            return ms.ToArray();
        }

        // ── Blob surgery ────────────────────────────────────────────────────

        // PDF annotation dict start marker — unique within the template.
        static readonly byte[] DataStartMarker = Encoding.ASCII.GetBytes("<</IT/PolygonCloud");

        static byte[] BuildModifiedBlob(byte[] template, MarkupRequest req, out int dataStringLen)
        {
            int markerPos = IndexOf(template, DataStartMarker, 0);
            if (markerPos < 0)
                throw new InvalidOperationException(
                    "Could not locate m_sData content in template — template file may be damaged.");

            // Walk backwards to find the BinaryFormatter 7-bit length prefix.
            int prefixEnd = markerPos - 1;
            if ((template[prefixEnd] & 0x80) != 0)
                throw new InvalidOperationException("Malformed length prefix: terminal byte has continuation bit set.");
            int prefixStart = prefixEnd;
            while (prefixStart > 0 && (template[prefixStart - 1] & 0x80) != 0)
                prefixStart--;

            int oldStringLen = DecodeVarInt(template, prefixStart, out int _);

            string oldData = Encoding.UTF8.GetString(template, markerPos, oldStringLen);

            string newData = RewritePdfDict(oldData, req);

            byte[] newDataBytes = Encoding.UTF8.GetBytes(newData);
            int    newStringLen = newDataBytes.Length;
            byte[] newPrefix    = EncodeVarInt(newStringLen);

            int suffixStart = markerPos + oldStringLen;
            int suffixLen   = template.Length - suffixStart;
            int newTotal    = prefixStart + newPrefix.Length + newStringLen + suffixLen;

            byte[] result = new byte[newTotal];
            Buffer.BlockCopy(template, 0,           result, 0,                                           prefixStart);
            Buffer.BlockCopy(newPrefix, 0,          result, prefixStart,                                 newPrefix.Length);
            Buffer.BlockCopy(newDataBytes, 0,       result, prefixStart + newPrefix.Length,              newStringLen);
            Buffer.BlockCopy(template, suffixStart, result, prefixStart + newPrefix.Length + newStringLen, suffixLen);

            dataStringLen = newStringLen;
            return result;
        }

        // ── PDF annotation dict rewriter ────────────────────────────────────

        static string RewritePdfDict(string pdfDict, MarkupRequest req)
        {
            string contents = req.Contents ?? "";
            string note     = req.Note     ?? "";   // BSI[0]
            string status   = req.Status   ?? "";   // BSI[1]
            string element  = req.Element  ?? "";   // BSI[2] (was: ItemCategory)
            string location = req.Location ?? "";   // BSI[3]
            string issue    = req.Issue    ?? "";   // BSI[4] — NEW in v2.4

            // Normalize newlines to CR — Bluebeam's native Contents hex
            // uses 000D (CR) for paragraph breaks.
            contents = contents.Replace("\r\n", "\r").Replace("\n", "\r");

            string rcHtml = req.RcHtml ?? BuildXhtmlFromPlainText(contents);

            // /Subj is no longer written. The Subject column was removed from
            // the TABSportal profile in v2.3 and the helper stopped populating
            // it in v2.4. The template's /Subj() entry stays empty.

            // BSI column order is contractual -- MUST match TABSportal.bpx UserDefined[0..4]
            // and the edge function's extractMarkups bsiStrs[0..4] mapping.
            pdfDict = ReplaceBsiColumnData(pdfDict,
                note:     note,      // BSI[0]
                status:   status,    // BSI[1]
                element:  element,   // BSI[2]
                location: location,  // BSI[3]
                issue:    issue);    // BSI[4]
            pdfDict = ReplacePdfLiteralString(pdfDict, "/RC", rcHtml);
            pdfDict = ReplacePdfHexString(pdfDict, "/Contents", EncodeContentsHex(contents));

            return pdfDict;
        }

        // --- PDF literal string helpers ------------------------------------

        static string EscapePdfLiteral(string s)
        {
            var sb = new StringBuilder(s.Length + 16);
            foreach (char c in s)
            {
                switch (c)
                {
                    case '(':  sb.Append("\\("); break;
                    case ')':  sb.Append("\\)"); break;
                    case '\\': sb.Append("\\\\"); break;
                    default:   sb.Append(c); break;
                }
            }
            return sb.ToString();
        }

        static string ReplacePdfLiteralString(string dict, string key, string newValue)
        {
            int keyIdx = dict.IndexOf(key + "(", StringComparison.Ordinal);
            if (keyIdx < 0) return dict;

            int openParen  = keyIdx + key.Length;                  // position of '('
            int closeParen = FindMatchingCloseParen(dict, openParen);
            if (closeParen < 0) return dict;

            return dict.Substring(0, openParen + 1)
                 + EscapePdfLiteral(newValue)
                 + dict.Substring(closeParen);
        }

        static int FindMatchingCloseParen(string s, int openParen)
        {
            int depth = 1;
            int i = openParen + 1;
            while (i < s.Length)
            {
                char c = s[i];
                if (c == '\\')  { i += 2; continue; }   // skip escaped byte
                if (c == '(')   depth++;
                else if (c == ')') { depth--; if (depth == 0) return i; }
                i++;
            }
            return -1;
        }

        static string ReplacePdfHexString(string dict, string key, string newHex)
        {
            int keyIdx = dict.IndexOf(key + "<", StringComparison.Ordinal);
            if (keyIdx < 0) return dict;

            int open  = keyIdx + key.Length;           // position of '<'
            int close = dict.IndexOf('>', open + 1);
            if (close < 0) return dict;

            return dict.Substring(0, open + 1)
                 + newHex
                 + dict.Substring(close);
        }

        /// <summary>
        /// Writes a 5-element PDF array into the /BSIColumnData field.
        /// Order MUST match the canonical TABS column schema:
        ///   BSI[0] = note, BSI[1] = status, BSI[2] = element,
        ///   BSI[3] = location, BSI[4] = issue
        /// </summary>
        static string ReplaceBsiColumnData(string dict,
            string note, string status, string element, string location, string issue)
        {
            int keyIdx = dict.IndexOf("/BSIColumnData[", StringComparison.Ordinal);
            if (keyIdx < 0) return dict;

            int open  = keyIdx + "/BSIColumnData".Length;   // position of '['
            int close = FindMatchingCloseBracket(dict, open);
            if (close < 0) return dict;

            string newArray = "["
                + "(" + EscapePdfLiteral(note)     + ")"   // BSI[0] — user note
                + "(" + EscapePdfLiteral(status)   + ")"   // BSI[1] — comment status
                + "(" + EscapePdfLiteral(element)  + ")"   // BSI[2] — element (was: ItemCategory)
                + "(" + EscapePdfLiteral(location) + ")"   // BSI[3] — location
                + "(" + EscapePdfLiteral(issue)    + ")"   // BSI[4] — issue (NEW)
                + "]";

            return dict.Substring(0, open)
                 + newArray
                 + dict.Substring(close + 1);
        }

        static int FindMatchingCloseBracket(string s, int openBracket)
        {
            int depth = 1;
            int i = openBracket + 1;
            while (i < s.Length)
            {
                char c = s[i];
                if (c == '\\') { i += 2; continue; }
                if (c == '(')                              // skip whole literal string
                {
                    int cp = FindMatchingCloseParen(s, i);
                    if (cp < 0) return -1;
                    i = cp + 1;
                    continue;
                }
                if (c == '[') depth++;
                else if (c == ']') { depth--; if (depth == 0) return i; }
                i++;
            }
            return -1;
        }

        // --- /Contents UTF-16BE hex encoding ------------------------------

        static string EncodeContentsHex(string text)
        {
            var sb = new StringBuilder(4 + text.Length * 4);
            sb.Append("feff");                                 // BOM
            foreach (char c in text)
            {
                sb.Append(((byte)(c >> 8)).ToString("x2"));
                sb.Append(((byte)(c & 0xFF)).ToString("x2"));
            }
            return sb.ToString();
        }

        // --- XHTML fallback (when /RC not supplied) -----------------------

        static string BuildXhtmlFromPlainText(string text)
        {
            var sb = new StringBuilder(
                "<?xml version=\"1.0\"?>"
              + "<body xmlns:xfa=\"http://www.xfa.org/schema/xfa-data/1.0/\" "
              + "xfa:contentType=\"text/html\" xfa:APIVersion=\"BluebeamPDFRevu:2018\" "
              + "xfa:spec=\"2.2.0\" xmlns=\"http://www.w3.org/1999/xhtml\">");

            if (string.IsNullOrEmpty(text)) { sb.Append("<p /></body>"); return sb.ToString(); }

            string[] paragraphs = Regex.Split(text, @"\r\r|\n\n|\r\n\r\n");
            for (int i = 0; i < paragraphs.Length; i++)
            {
                if (i > 0) sb.Append("<p />");
                sb.Append("<p>");
                sb.Append(EscapeXml(paragraphs[i].Replace('\r', ' ').Replace('\n', ' ')));
                sb.Append("</p>");
            }
            sb.Append("</body>");
            return sb.ToString();
        }

        static string EscapeXml(string s) => s
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;");

        // --- BinaryFormatter 7-bit LEB128 --------------------------------

        static int DecodeVarInt(byte[] buf, int offset, out int consumed)
        {
            int value = 0, shift = 0, start = offset;
            while (true)
            {
                byte b = buf[offset++];
                value |= (b & 0x7F) << shift;
                if ((b & 0x80) == 0) break;
                shift += 7;
                if (shift > 28) throw new InvalidOperationException("VarInt too long");
            }
            consumed = offset - start;
            return value;
        }

        static byte[] EncodeVarInt(int value)
        {
            if (value < 0) throw new ArgumentOutOfRangeException(nameof(value));
            var bytes = new List<byte>(5);
            while (value >= 0x80)
            {
                bytes.Add((byte)((value & 0x7F) | 0x80));
                value >>= 7;
            }
            bytes.Add((byte)value);
            return bytes.ToArray();
        }

        // --- Byte array search -------------------------------------------

        static int IndexOf(byte[] haystack, byte[] needle, int startIndex)
        {
            int nlen = needle.Length;
            int hmax = haystack.Length - nlen;
            for (int i = startIndex; i <= hmax; i++)
            {
                bool match = true;
                for (int j = 0; j < nlen; j++)
                    if (haystack[i + j] != needle[j]) { match = false; break; }
                if (match) return i;
            }
            return -1;
        }

        // ── Win32 clipboard write ───────────────────────────────────────────

        static void WriteToClipboard(byte[] data)
        {
            uint format = RegisterClipboardFormatW(BluebeamClipboardFormat);
            if (format == 0)
                throw new InvalidOperationException(
                    $"RegisterClipboardFormat failed ({Marshal.GetLastWin32Error()})");

            // Retry OpenClipboard briefly — other apps may hold it transiently.
            bool opened = false;
            for (int i = 0; i < 10; i++)
            {
                if (OpenClipboard(IntPtr.Zero)) { opened = true; break; }
                Thread.Sleep(50);
            }
            if (!opened)
                throw new InvalidOperationException(
                    $"OpenClipboard failed after retries ({Marshal.GetLastWin32Error()})");

            IntPtr hGlobal = IntPtr.Zero;
            try
            {
                if (!EmptyClipboard())
                    throw new InvalidOperationException(
                        $"EmptyClipboard failed ({Marshal.GetLastWin32Error()})");

                hGlobal = GlobalAlloc(GMEM_MOVEABLE, (UIntPtr)data.Length);
                if (hGlobal == IntPtr.Zero)
                    throw new InvalidOperationException(
                        $"GlobalAlloc failed for {data.Length} bytes ({Marshal.GetLastWin32Error()})");

                IntPtr ptr = GlobalLock(hGlobal);
                if (ptr == IntPtr.Zero)
                    throw new InvalidOperationException(
                        $"GlobalLock failed ({Marshal.GetLastWin32Error()})");

                try   { Marshal.Copy(data, 0, ptr, data.Length); }
                finally { GlobalUnlock(hGlobal); }

                // On success, the OS takes ownership of hGlobal — DO NOT GlobalFree.
                if (SetClipboardData(format, hGlobal) == IntPtr.Zero)
                    throw new InvalidOperationException(
                        $"SetClipboardData failed ({Marshal.GetLastWin32Error()})");

                hGlobal = IntPtr.Zero;   // ownership transferred
            }
            finally
            {
                if (hGlobal != IntPtr.Zero) GlobalFree(hGlobal);
                CloseClipboard();
            }
        }
    }
}
