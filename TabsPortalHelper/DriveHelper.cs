using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Data.Sqlite;
using Microsoft.Win32;

namespace TabsPortalHelper
{
    static class DriveHelper
    {
        // ════════════════════════════════════════════════════════════════════════
        // FILE ID → LOCAL PATH via DriveFS SQLite database
        //
        // Confirmed schema from live DB inspection:
        //   stable_ids  (stable_id INTEGER, cloud_id TEXT)
        //   items       (stable_id INTEGER, local_title TEXT, is_folder INTEGER)
        //   stable_parents (item_stable_id INTEGER, parent_stable_id INTEGER)
        // ════════════════════════════════════════════════════════════════════════
        public static string? FindLocalPathByFileId(string driveFileId)
        {
            var driveFsRoot = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Google", "DriveFS");

            if (!Directory.Exists(driveFsRoot)) return null;

            foreach (var accountDir in Directory.GetDirectories(driveFsRoot))
            {
                var dbPath = Path.Combine(accountDir, "metadata_sqlite_db");
                if (!File.Exists(dbPath)) continue;

                try
                {
                    var result = ResolvePathFromDb(dbPath, driveFileId);
                    if (result != null) return result;
                }
                catch { /* DB locked or schema mismatch — try next account */ }
            }

            return null;
        }

        static string? ResolvePathFromDb(string dbPath, string cloudFileId)
        {
            var connStr = new SqliteConnectionStringBuilder
            {
                DataSource = dbPath,
                Mode = SqliteOpenMode.ReadOnly,
                Cache = SqliteCacheMode.Shared,
            }.ToString();

            using var conn = new SqliteConnection(connStr);
            conn.Open();

            // Step 1: cloud_id → internal stable_id integer
            long stableId;
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT stable_id FROM stable_ids WHERE cloud_id = @id LIMIT 1";
                cmd.Parameters.AddWithValue("@id", cloudFileId);
                var result = cmd.ExecuteScalar();
                if (result == null) return null;
                stableId = Convert.ToInt64(result);
            }

            // Step 2: get local filename
            string? localTitle;
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT local_title FROM items WHERE stable_id = @id LIMIT 1";
                cmd.Parameters.AddWithValue("@id", stableId);
                localTitle = cmd.ExecuteScalar() as string;
            }

            if (localTitle == null) return null;

            // Step 3: walk parent chain to build full path
            var pathParts = new List<string> { localTitle };
            long currentId = stableId;

            for (int depth = 0; depth < 30; depth++)
            {
                long? parentId = null;
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "SELECT parent_stable_id FROM stable_parents WHERE item_stable_id = @id LIMIT 1";
                    cmd.Parameters.AddWithValue("@id", currentId);
                    var result = cmd.ExecuteScalar();
                    if (result == null) break;
                    parentId = Convert.ToInt64(result);
                }

                string? parentTitle;
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "SELECT local_title FROM items WHERE stable_id = @id LIMIT 1";
                    cmd.Parameters.AddWithValue("@id", parentId!.Value);
                    parentTitle = cmd.ExecuteScalar() as string;
                }

                if (parentTitle == null) break;

                // "My Drive" is the root — stop here, Drive root prepended below
                if (parentTitle.Equals("My Drive", StringComparison.OrdinalIgnoreCase)) break;

                pathParts.Add(parentTitle);
                currentId = parentId!.Value;
            }

            // Step 4: prepend Drive root
            var driveRoot = FindDriveRoot();
            if (driveRoot == null) return null;

            pathParts.Reverse();
            return Path.Combine(new[] { driveRoot }.Concat(pathParts).ToArray());
        }

        // ════════════════════════════════════════════════════════════════════════
        // DRIVE ROOT DISCOVERY
        // ════════════════════════════════════════════════════════════════════════
        public static string? FindDriveRoot()
        {
            // 0. Manual override
            var envOverride = Environment.GetEnvironmentVariable("TABS_DRIVE_ROOT");
            if (!string.IsNullOrWhiteSpace(envOverride) && Directory.Exists(envOverride))
                return envOverride;

            // 1. DriveFS registry (modern Drive for Desktop)
            try
            {
                using var prefsKey = Registry.CurrentUser.OpenSubKey(
                    @"Software\Google\DriveFS\PerAccountPreferences");
                if (prefsKey != null)
                {
                    foreach (var accountId in prefsKey.GetSubKeyNames())
                    {
                        using var accountKey = prefsKey.OpenSubKey(accountId);
                        var mountPoint = accountKey?.GetValue("mount_point_path") as string;
                        if (!string.IsNullOrWhiteSpace(mountPoint))
                        {
                            var myDrive = Path.Combine(mountPoint, "My Drive");
                            if (Directory.Exists(myDrive)) return myDrive;
                            if (Directory.Exists(mountPoint)) return mountPoint;
                        }
                    }
                }
            }
            catch { }

            // 2. Older Backup and Sync
            try
            {
                using var shareKey = Registry.CurrentUser.OpenSubKey(@"Software\Google\Drive\Preferences");
                var rootPath = shareKey?.GetValue("mount_point_path") as string;
                if (!string.IsNullOrWhiteSpace(rootPath) && Directory.Exists(rootPath))
                    return rootPath;
            }
            catch { }

            // 3. Common default paths
            var userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            var candidates = new[]
            {
                @"G:\My Drive",
                @"H:\My Drive",
                @"I:\My Drive",
                Path.Combine(userProfile, "Google Drive", "My Drive"),
                Path.Combine(userProfile, "Google Drive"),
                Path.Combine(userProfile, "My Drive"),
            };

            var found = candidates.FirstOrDefault(Directory.Exists);
            if (found != null) return found;

            // 4. Scan all drive letters
            foreach (var drive in DriveInfo.GetDrives())
            {
                try
                {
                    if (drive.IsReady)
                    {
                        var myDrive = Path.Combine(drive.RootDirectory.FullName, "My Drive");
                        if (Directory.Exists(myDrive)) return myDrive;
                    }
                }
                catch { }
            }

            return null;
        }

        public static void ShowDriveRootDialog()
        {
            var root = FindDriveRoot();
            if (root != null)
            {
                MessageBox.Show(
                    $"Google Drive root detected:\n\n{root}",
                    "TABS Portal Helper — Drive Root",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(
                    "Could not detect Google Drive root.\n\n" +
                    "Make sure Google Drive for Desktop is installed and signed in.\n\n" +
                    "To override manually, set the environment variable:\n" +
                    "TABS_DRIVE_ROOT=C:\\Users\\you\\My Drive",
                    "TABS Portal Helper — Drive Root Not Found",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }
    }
}
