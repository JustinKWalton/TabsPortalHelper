using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Win32;

namespace TabsPortalHelper
{
    static class Installer
    {
        const string AppName = "TABS Portal Helper";
        const string AppVersion = "2.0.0";
        const string UninstallRegKey = @"Software\Microsoft\Windows\CurrentVersion\Uninstall\TabsPortalHelper";
        const string StartupRegKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
        const string TabsRegKey = @"Software\TabsPortalHelper";

        static string ExePath =>
            Path.Combine(AppContext.BaseDirectory, "TabsPortalHelper.exe");

        // ════════════════════════════════════════════════════════════════════════
        // FULL INSTALL
        // ════════════════════════════════════════════════════════════════════════
        public static void Install()
        {
            try
            {
                RegisterStartup();
                RegisterAddRemovePrograms();
                SaveDiagnostics();

                var driveRoot = DriveHelper.FindDriveRoot();
                var driveMsg = driveRoot != null
                    ? $"\n\nGoogle Drive detected at:\n{driveRoot}"
                    : "\n\n⚠ Google Drive root not detected.\nMake sure Drive for Desktop is installed and signed in.";

                // ── Bluebeam custom columns (Note, Comment Status, Element, Location) ──
                // Idempotent: if columns already present, status is NotNeeded and nothing is touched.
                // If Bluebeam is running, we surface a message telling the user how to finish setup.
                string columnMsg;
                try
                {
                    var columnResult = ColumnInstaller.CheckAndInstall();
                    switch (columnResult.Status)
                    {
                        case ColumnInstaller.InstallStatus.Installed:
                            columnMsg = $"\n\n✓ Bluebeam columns configured ({columnResult.TouchedFiles.Count} profile(s)).";
                            break;
                        case ColumnInstaller.InstallStatus.NotNeeded:
                            columnMsg = "\n\n✓ Bluebeam columns already configured.";
                            break;
                        case ColumnInstaller.InstallStatus.BluebeamRunning:
                            columnMsg = "\n\n⚠ Bluebeam Revu is running. Close it, then right-click the TABS " +
                                        "tray icon and choose 'Install Bluebeam Columns...' to finish setup.";
                            break;
                        case ColumnInstaller.InstallStatus.NoProfileFound:
                            columnMsg = "\n\n⚠ No Bluebeam Revu 21/2024/2025 profile found. Launch Bluebeam once " +
                                        "to create one, then use the tray menu 'Install Bluebeam Columns...'.";
                            break;
                        case ColumnInstaller.InstallStatus.ConflictDetected:
                            columnMsg = "\n\n⚠ Bluebeam column conflict: " + columnResult.Message;
                            break;
                        default:
                            columnMsg = $"\n\n⚠ Bluebeam column setup: {columnResult.Status} — {columnResult.Message}";
                            break;
                    }
                }
                catch (Exception cex)
                {
                    // Never fail the whole install because columns couldn't be set up.
                    columnMsg = "\n\n⚠ Bluebeam column setup skipped: " + cex.Message;
                }

                MessageBox.Show(
                    $"{AppName} v{AppVersion} installed successfully!\n\n" +
                    $"The helper is now running in your system tray (look for the TABS icon " +
                    $"near the clock) and will start automatically with Windows.\n\n" +
                    $"If you downloaded an installer file, that download can now be deleted — " +
                    $"the helper has been copied to its permanent location.{driveMsg}{columnMsg}",
                    AppName,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                // Leave the user in a working state: launch the tray app so they don't
                // have to wait until next Windows login for the icon to appear.
                LaunchTrayApp();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Installation failed:\n{ex.Message}",
                    $"{AppName} — Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        // ════════════════════════════════════════════════════════════════════════
        // FULL UNINSTALL
        // ════════════════════════════════════════════════════════════════════════
        public static void Uninstall()
        {
            try
            {
                // Remove from startup
                using (var key = Registry.CurrentUser.OpenSubKey(StartupRegKey, writable: true))
                    key?.DeleteValue("TabsPortalHelper", throwOnMissingValue: false);

                // Remove from Add/Remove Programs
                Registry.CurrentUser.DeleteSubKeyTree(UninstallRegKey, throwOnMissingSubKey: false);

                // Remove diagnostics key
                Registry.CurrentUser.DeleteSubKeyTree(TabsRegKey, throwOnMissingSubKey: false);

                MessageBox.Show(
                    $"{AppName} has been uninstalled.\n\nYou can delete the application folder manually.",
                    AppName,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Uninstall failed:\n{ex.Message}",
                    $"{AppName} — Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        // ════════════════════════════════════════════════════════════════════════
        // REGISTER STARTUP (run at Windows login)
        // ════════════════════════════════════════════════════════════════════════
        public static void RegisterStartup()
        {
            using var key = Registry.CurrentUser.OpenSubKey(StartupRegKey, writable: true);
            key?.SetValue("TabsPortalHelper", $"\"{ExePath}\"");
        }

        // ════════════════════════════════════════════════════════════════════════
        // ADD/REMOVE PROGRAMS ENTRY
        // Writes to HKCU so no admin required
        // ════════════════════════════════════════════════════════════════════════
        static void RegisterAddRemovePrograms()
        {
            using var key = Registry.CurrentUser.CreateSubKey(UninstallRegKey);
            key.SetValue("DisplayName", AppName);
            key.SetValue("DisplayVersion", AppVersion);
            key.SetValue("Publisher", "Texas Accessibility Solutions");
            key.SetValue("DisplayIcon", ExePath);
            key.SetValue("UninstallString", $"\"{ExePath}\" --uninstall");
            key.SetValue("QuietUninstallString", $"\"{ExePath}\" --uninstall");
            key.SetValue("InstallLocation", AppContext.BaseDirectory);
            key.SetValue("InstallDate", DateTime.Now.ToString("yyyyMMdd"));
            key.SetValue("NoModify", 1, RegistryValueKind.DWord);
            key.SetValue("NoRepair", 1, RegistryValueKind.DWord);
        }

        static void SaveDiagnostics()
        {
            using var key = Registry.CurrentUser.CreateSubKey(TabsRegKey);
            key.SetValue("InstalledVersion", AppVersion);
            key.SetValue("InstalledAt", DateTime.UtcNow.ToString("o"));
            key.SetValue("ExePath", ExePath);
            key.SetValue("DriveRootDetected", DriveHelper.FindDriveRoot() ?? "not found");
        }

        // ════════════════════════════════════════════════════════════════════════
        // POST-INSTALL TRAY LAUNCH
        // Starts the tray app so the user has a working system immediately after
        // --install completes, instead of waiting until next Windows login.
        // ════════════════════════════════════════════════════════════════════════
        static void LaunchTrayApp()
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = ExePath,
                    UseShellExecute = false,
                    WorkingDirectory = AppContext.BaseDirectory,
                });
            }
            catch
            {
                // Best-effort. If this fails the user can launch manually, and
                // Windows will start it on next login via the registered startup key.
            }
        }
    }
}
