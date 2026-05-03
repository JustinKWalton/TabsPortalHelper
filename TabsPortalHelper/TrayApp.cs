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

        const string Version  = "2.5.4";
        const int    HttpPort = 52874;

        public TrayApp()
        {
            // -- Load tray icon from embedded resource --
            var icon = LoadEmbeddedIcon();

            // -- Build context menu --
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

            var profileItem = new ToolStripMenuItem("Install Bluebeam Profile...");
            profileItem.Click += (s, e) => InstallBluebeamProfile();
            menu.Items.Add(profileItem);

            menu.Items.Add(new ToolStripSeparator());

            var uninstallItem = new ToolStripMenuItem("Uninstall...");
            uninstallItem.Click += (s, e) => PromptUninstall();
            menu.Items.Add(uninstallItem);

            menu.Items.Add(new ToolStripSeparator());

            var exitItem = new ToolStripMenuItem("Exit");
            exitItem.Click += (s, e) => ExitApp();
            menu.Items.Add(exitItem);

            // -- Create tray icon --
            _trayIcon = new NotifyIcon
            {
                Icon             = icon,
                Text             = "TABS Portal Helper",
                ContextMenuStrip = menu,
                Visible          = true
            };
            _trayIcon.DoubleClick += (s, e) => ShowStatus();

            // -- Start HTTP server --
            _server = new HttpServer(HttpPort);
            _server.Start();

            // -- Show startup balloon --
            _trayIcon.ShowBalloonTip(
                3000,
                "TABS Portal Helper",
                $"Running on port {HttpPort} \u2014 ready to open files in Bluebeam.",
                ToolTipIcon.Info);
        }

        void ShowStatus()
        {
            var driveRoot = DriveHelper.FindDriveRoot();
            var bluebeam  = BluebeamHelper.FindBluebeam();

            MessageBox.Show(
                $"TABS Portal Helper  v{Version}\n\n" +
                $"HTTP Server:  localhost:{HttpPort}  \u2713\n\n" +
                $"Google Drive:  {driveRoot ?? "\u26A0 Not detected"}\n\n" +
                $"Bluebeam Revu:  {bluebeam ?? "\u26A0 Not found"}",
                "TABS Portal Helper \u2014 Status",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        void InstallBluebeamProfile()
        {
            var confirm = MessageBox.Show(
                "About to install the TABSportal Bluebeam profile.\n\n" +
                "This adds a new profile alongside your existing Bluebeam profiles \u2014 " +
                "your current profiles will not be modified. After install, Bluebeam will " +
                "open with TABSportal as the active profile. You can switch back to any " +
                "other profile any time via Revu \u2192 Profiles.\n\n" +
                "Continue?",
                "TABS \u2014 Install Bluebeam Profile",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button1);

            if (confirm != DialogResult.Yes)
                return;

            ProfileInstaller.InstallResult result;
            try
            {
                result = ProfileInstaller.CheckAndInstall();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Unexpected error while installing the Bluebeam profile:\n\n" + ex.Message,
                    "TABS \u2014 Bluebeam Profile",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            using var dlg = new ProfileInstallDialog(
                "TABS \u2014 Bluebeam Profile",
                preamble: string.Empty,
                result);
            dlg.ShowDialog();
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
