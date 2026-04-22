using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace TabsPortalHelper
{
    static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            // ── Handle CLI commands (install/uninstall run from PowerShell) ───────
            if (args.Length == 1)
            {
                switch (args[0].ToLowerInvariant())
                {
                    case "--install":
                        Installer.Install();
                        return;
                    case "--uninstall":
                        Installer.Uninstall();
                        return;
                    case "--drive-root":
                        DriveHelper.ShowDriveRootDialog();
                        return;
                }
            }

            // ── Install bootstrap ────────────────────────────────────────────────
            // If the .exe is being run from anywhere other than %LOCALAPPDATA%\TabsPortalHelper
            // (e.g., double-clicked out of Downloads), offer to install it there.
            if (!IsRunningFromInstallDir())
            {
                if (PromptAndBootstrapInstall())
                {
                    // Bootstrap launched the installed copy — this process is done.
                    return;
                }
                // User declined; fall through and run as a one-off tray instance
                // from wherever they launched us. Harmless — no registry touched.
            }

            // ── Single instance guard — don't run two tray apps ──────────────────
            using var mutex = new Mutex(true, "TabsPortalHelper_SingleInstance", out bool isNew);
            if (!isNew)
            {
                MessageBox.Show(
                    "TABS Portal Helper is already running.\n\nCheck the system tray.",
                    "TABS Portal Helper",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new TrayApp());
        }

        // ── Install bootstrap helpers ────────────────────────────────────────────

        static string GetCanonicalInstallPath()
        {
            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "TabsPortalHelper");
            return Path.Combine(dir, "TabsPortalHelper.exe");
        }

        static bool IsRunningFromInstallDir()
        {
            try
            {
                var running = Path.GetFullPath(Application.ExecutablePath);
                var canonical = Path.GetFullPath(GetCanonicalInstallPath());
                return string.Equals(running, canonical, StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Prompts the user to install, copies self to %LOCALAPPDATA%\TabsPortalHelper,
        /// launches the installed copy with --install, and returns true if bootstrap
        /// succeeded (caller should exit). Returns false if user declined or install failed.
        /// </summary>
        static bool PromptAndBootstrapInstall()
        {
            var installedExe = GetCanonicalInstallPath();
            bool isUpgrade = File.Exists(installedExe);

            var prompt = isUpgrade
                ? "TABS Portal Helper is already installed on this computer.\n\n" +
                  "Replace it with this version?\n\n" +
                  "The running tray app will be closed automatically."
                : "Install TABS Portal Helper?\n\n" +
                  "It will be installed to your user profile and start automatically " +
                  "when you sign in to Windows. Administrator rights are not required.";

            var title = isUpgrade
                ? "TABS Portal Helper — Upgrade"
                : "TABS Portal Helper — Install";

            var result = MessageBox.Show(
                prompt,
                title,
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button1);

            if (result != DialogResult.Yes) return false;

            try
            {
                // Stop any running tray instance so we can overwrite its .exe.
                int myPid = Process.GetCurrentProcess().Id;
                var running = Process.GetProcessesByName("TabsPortalHelper")
                    .Where(p => p.Id != myPid)
                    .ToArray();
                foreach (var p in running)
                {
                    try { p.Kill(entireProcessTree: true); p.WaitForExit(3000); }
                    catch { /* best-effort */ }
                }
                if (running.Length > 0) Thread.Sleep(700); // let file handles drain

                // Copy self into %LOCALAPPDATA%\TabsPortalHelper\
                var installDir = Path.GetDirectoryName(installedExe)!;
                Directory.CreateDirectory(installDir);
                File.Copy(Application.ExecutablePath, installedExe, overwrite: true);

                // Launch the installed copy with --install so it registers
                // startup + Add/Remove Programs and starts the tray.
                Process.Start(new ProcessStartInfo
                {
                    FileName = installedExe,
                    Arguments = "--install",
                    UseShellExecute = false,
                    WorkingDirectory = installDir,
                });

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Install failed:\n\n" + ex.Message + "\n\n" +
                    "You can try running the .exe again, or contact support.",
                    "TABS Portal Helper",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return false;
            }
        }
    }
}
