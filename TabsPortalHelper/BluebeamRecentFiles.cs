using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace TabsPortalHelper
{
    /// <summary>
    /// Reads Bluebeam Revu's recent-files database so the TABS app can offer
    /// "the plan set you just had open in Bluebeam" without a file dialog.
    ///
    /// Revu (21 / 20 / 2019) keeps %APPDATA%\Bluebeam Software\Revu\{ver}\RdbRecentFiles.dat.
    /// It is an undocumented binary, but every record carries:
    ///   [... 8-byte .NET ticks (open time) ...] [len][folder] [len][filename] [len][view state]
    /// Strings are ASCII with a single length byte (lengths &lt; 128). We scan for
    /// "X:\..." runs, read the length-prefixed folder + filename, and take the
    /// newest .NET ticks value in the 80 bytes before the folder as the open time.
    /// Read-only; Revu holds the file with shared access so it's safe while running.
    /// </summary>
    public static class BluebeamRecentFiles
    {
        public record RecentFile(string Path, string Folder, string Name, DateTime OpenedAt, bool Exists, long SizeBytes, DateTime? ModifiedAt);

        static readonly long TicksMin = new DateTime(2015, 1, 1).Ticks;

        public static string? FindDatabase()
        {
            var root = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Bluebeam Software", "Revu");
            if (!Directory.Exists(root)) return null;
            // Newest Revu version first (21 > 20 > 2019).
            var candidates = Directory.GetDirectories(root)
                .Select(d => System.IO.Path.Combine(d, "RdbRecentFiles.dat"))
                .Where(File.Exists)
                .OrderByDescending(p => File.GetLastWriteTimeUtc(p))
                .ToList();
            return candidates.FirstOrDefault();
        }

        public static List<RecentFile> Read(int limit = 15, bool pdfOnly = true)
        {
            var db = FindDatabase();
            if (db == null) return new List<RecentFile>();
            byte[] b;
            using (var fs = new FileStream(db, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (var ms = new MemoryStream())
            {
                fs.CopyTo(ms);
                b = ms.ToArray();
            }
            var ticksMax = DateTime.UtcNow.AddDays(1).Ticks;
            var best = new Dictionary<string, RecentFile>(StringComparer.OrdinalIgnoreCase);

            for (int i = 2; i < b.Length - 4; i++)
            {
                // "X:\" at i, preceded by its length byte
                if (b[i + 1] != (byte)':' || b[i + 2] != (byte)'\\') continue;
                char drive = (char)b[i];
                if (!char.IsLetter(drive)) continue;
                int folderLen = b[i - 1];
                if (folderLen < 4 || folderLen > 120 || i + folderLen > b.Length) continue;
                if (!IsPrintable(b, i, folderLen)) continue;
                string folder = Encoding.ASCII.GetString(b, i, folderLen);
                int p = i + folderLen;
                if (p >= b.Length) continue;
                int nameLen = b[p];
                if (nameLen < 5 || nameLen > 120 || p + 1 + nameLen > b.Length) continue;
                if (!IsPrintable(b, p + 1, nameLen)) continue;
                string name = Encoding.ASCII.GetString(b, p + 1, nameLen);
                if (pdfOnly && !name.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) continue;

                // Open time: newest plausible .NET ticks in the 80 bytes before the folder.
                long bestTicks = 0;
                int from = Math.Max(0, i - 1 - 80);
                for (int o = from; o <= i - 1 - 8; o++)
                {
                    long t = BitConverter.ToInt64(b, o);
                    if (t > TicksMin && t < ticksMax && t > bestTicks) bestTicks = t;
                }
                if (bestTicks == 0) continue;
                var when = new DateTime(bestTicks, DateTimeKind.Unspecified);

                string full = System.IO.Path.Combine(folder, name);
                if (!best.TryGetValue(full, out var existing) || existing.OpenedAt < when)
                {
                    bool exists = File.Exists(full);
                    long size = 0;
                    DateTime? modified = null;
                    if (exists) { try { var fi = new FileInfo(full); size = fi.Length; modified = fi.LastWriteTime; } catch { } }
                    best[full] = new RecentFile(full, folder, name, when, exists, size, modified);
                }
                i = p + nameLen; // skip past this record
            }

            return best.Values.OrderByDescending(r => r.OpenedAt).Take(limit).ToList();
        }

        static bool IsPrintable(byte[] b, int start, int len)
        {
            for (int k = start; k < start + len; k++) if (b[k] < 0x20 || b[k] > 0x7e) return false;
            return true;
        }
    }
}
