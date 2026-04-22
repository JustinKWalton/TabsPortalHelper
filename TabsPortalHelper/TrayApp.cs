using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace TabsPortalHelper
{
    class TrayApp : ApplicationContext
    {
        private readonly NotifyIcon _trayIcon;
        private readonly HttpServer _server;
        const string Version = "2.0.0";
        const int HttpPort = 52874;

        public TrayApp()
        {
            // ── Load tray icon from embedded resource ────────────────────────────
            var icon = LoadEmbeddedIcon();

            // ── Build context menu ───────────────────────────────────────────────
            var menu = new ContextMenuStrip();

            var header = new ToolStripMenuItem("TABS Portal Helper  v" + Version)
            {
                Enabled = false,
                Font = new Font(SystemFonts.MenuFont!, FontStyle.Bold)
            };
            menu.Items.Add(header);
            menu.Items.Add(new ToolStripSeparator());

            var driveRootItem = new ToolStripMenuItem("Show Drive Root...");
            driveRootItem.Click += (s, e) => DriveHelper.ShowDriveRootDialog();
            menu.Items.Add(driveRootItem);

            var reinstallItem = new ToolStripMenuItem("Re-register on Startup...");
            reinstallItem.Click += (s, e) => Installer.RegisterStartup();
            menu.Items.Add(reinstallItem);

            var columnsItem = new ToolStripMenuItem("Install Bluebeam Columns...");
            columnsItem.Click += (s, e) => InstallBluebeamColumns();
            menu.Items.Add(columnsItem);

            menu.Items.Add(new ToolStripSeparator());

            var uninstallItem = new ToolStripMenuItem("Uninstall...");
            uninstallItem.Click += (s, e) => PromptUninstall();
            menu.Items.Add(uninstallItem);

            menu.Items.Add(new ToolStripSeparator());

            var exitItem = new ToolStripMenuItem("Exit");
            exitItem.Click += (s, e) => ExitApp();
            menu.Items.Add(exitItem);

            // ── Create tray icon ─────────────────────────────────────────────────
            _trayIcon = new NotifyIcon
            {
                Icon = icon,
                Text = "TABS Portal Helper",
                ContextMenuStrip = menu,
                Visible = true
            };

            _trayIcon.DoubleClick += (s, e) => ShowStatus();

            // ── Start HTTP server ────────────────────────────────────────────────
            _server = new HttpServer(HttpPort);
            _server.Start();

            // ── Show startup balloon ─────────────────────────────────────────────
            _trayIcon.ShowBalloonTip(
                3000,
                "TABS Portal Helper",
                $"Running on port {HttpPort} — ready to open files in Bluebeam.",
                ToolTipIcon.Info);
        }

        void ShowStatus()
        {
            var driveRoot = DriveHelper.FindDriveRoot();
            var bluebeam = BluebeamHelper.FindBluebeam();

            MessageBox.Show(
                $"TABS Portal Helper  v{Version}\n\n" +
                $"HTTP Server:  localhost:{HttpPort}  ✓\n\n" +
                $"Google Drive:  {driveRoot ?? "⚠ Not detected"}\n\n" +
                $"Bluebeam Revu:  {bluebeam ?? "⚠ Not found"}",
                "TABS Portal Helper — Status",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        void InstallBluebeamColumns()
        {
            ColumnInstaller.InstallResult result;
            try
            {
                result = ColumnInstaller.CheckAndInstall();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Unexpected error while setting up Bluebeam columns:\n\n" + ex.Message,
                    "TABS — Bluebeam Columns",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            string msg;
            MessageBoxIcon icon;
            switch (result.Status)
            {
                case ColumnInstaller.InstallStatus.Installed:
                    msg = $"TABS columns installed in {result.TouchedFiles.Count} Bluebeam profile(s).\n\n" +
                          "A .tabsbackup sidecar of the original was saved alongside each modified profile.";
                    icon = MessageBoxIcon.Information;
                    break;
                case ColumnInstaller.InstallStatus.NotNeeded:
                    msg = "TABS columns are already set up in Bluebeam. No changes needed.";
                    icon = MessageBoxIcon.Information;
                    break;
                case ColumnInstaller.InstallStatus.BluebeamRunning:
                    msg = "Bluebeam Revu is currently running.\n\nPlease close Bluebeam completely and try again.";
                    icon = MessageBoxIcon.Warning;
                    break;
                case ColumnInstaller.InstallStatus.NoProfileFound:
                    msg = "No Bluebeam Revu profile was found for version 21, 2024, or 2025.\n\n" +
                          "Launch Bluebeam once to let it create a default profile, then try again.";
                    icon = MessageBoxIcon.Warning;
                    break;
                case ColumnInstaller.InstallStatus.ConflictDetected:
                    msg = "Cannot install TABS columns automatically:\n\n" + result.Message +
                          "\n\nContact support if you need help resolving the conflict.";
                    icon = MessageBoxIcon.Warning;
                    break;
                default:
                    msg = "TABS column setup failed:\n\n" + (result.Message ?? "unknown error");
                    icon = MessageBoxIcon.Error;
                    break;
            }
            MessageBox.Show(msg, "TABS — Bluebeam Columns", MessageBoxButtons.OK, icon);
        }

        void PromptUninstall()
        {
            var result = MessageBox.Show(
                "This will uninstall TABS Portal Helper and remove it from startup.\n\nContinue?",
                "Uninstall TABS Portal Helper",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                Installer.Uninstall();
                ExitApp();
            }
        }

        void ExitApp()
        {
            _server.Stop();
            _trayIcon.Visible = false;
            _trayIcon.Dispose();
            Application.Exit();
        }

        static Icon LoadEmbeddedIcon()
        {
            try
            {
                var asm = Assembly.GetExecutingAssembly();
                // Embedded resource name: TabsPortalHelper.tray.ico
                using var stream = asm.GetManifestResourceStream("TabsPortalHelper.tray.ico");
                if (stream != null) return new Icon(stream);
            }
            catch { }

            // Fallback to system default
            return SystemIcons.Application;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _server?.Stop();
                _trayIcon?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
