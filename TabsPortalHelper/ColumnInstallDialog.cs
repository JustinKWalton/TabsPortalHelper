using System;
using System.Diagnostics;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TabsPortalHelper
{
    /// <summary>
    /// Unified post-install / column-install dialog.
    ///
    /// Renders the current <see cref="ColumnInstaller.InstallResult"/>, and when the
    /// status is <see cref="ColumnInstaller.InstallStatus.BluebeamRunning"/> shows a
    /// Retry button that re-runs <see cref="ColumnInstaller.CheckAndInstall"/> in place
    /// instead of spawning a new dialog. All other statuses show an OK button.
    ///
    /// Called from two places:
    ///   • Installer.Install()  — pass the full install-success text as <paramref name="preamble"/>.
    ///   • TrayApp.InstallBluebeamColumns() — pass an empty preamble; the dialog auto-shrinks.
    /// </summary>
    public sealed class ColumnInstallDialog : Form
    {
        private readonly string _preamble;

        private readonly PictureBox _iconBox;
        private readonly Label      _messageLabel;
        private readonly Button     _primaryButton;
        private readonly Button     _secondaryButton;

        private readonly Point _primaryAlonePos;
        private readonly Point _primaryWithSecondaryPos;

        private ColumnInstaller.InstallResult _result;
        private bool _terminalError;

        public ColumnInstallDialog(
            string windowTitle,
            string preamble,
            ColumnInstaller.InstallResult initialResult)
        {
            _preamble = preamble ?? string.Empty;
            _result   = initialResult;

            bool hasPreamble = _preamble.Length > 0;

            Text            = windowTitle;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox     = false;
            MinimizeBox     = false;
            ShowInTaskbar   = false;
            StartPosition   = FormStartPosition.CenterScreen;
            ClientSize      = hasPreamble ? new Size(520, 360) : new Size(480, 200);

            const int Pad    = 16;
            const int BtnW   = 100;
            const int BtnH   = 28;
            const int BtnGap = 6;

            int btnY = ClientSize.Height - BtnH - Pad;
            _primaryAlonePos         = new Point(ClientSize.Width - Pad - BtnW, btnY);
            _primaryWithSecondaryPos = new Point(ClientSize.Width - Pad - (BtnW * 2) - BtnGap, btnY);
            var secondaryPos         = new Point(ClientSize.Width - Pad - BtnW, btnY);

            _iconBox = new PictureBox
            {
                Location = new Point(Pad, 20),
                Size     = new Size(32, 32),
                SizeMode = PictureBoxSizeMode.StretchImage,
            };

            _messageLabel = new Label
            {
                Location = new Point(60, 16),
                Size     = new Size(ClientSize.Width - 60 - Pad, btnY - 36),
                AutoSize = false,
            };

            _primaryButton = new Button { Size = new Size(BtnW, BtnH) };
            _primaryButton.Click += OnPrimaryClick;

            _secondaryButton = new Button
            {
                Text         = "Skip for now",
                Size         = new Size(BtnW, BtnH),
                Location     = secondaryPos,
                DialogResult = DialogResult.Cancel,
            };

            AcceptButton = _primaryButton;
            CancelButton = _secondaryButton;

            Controls.Add(_iconBox);
            Controls.Add(_messageLabel);
            Controls.Add(_primaryButton);
            Controls.Add(_secondaryButton);

            RenderFromResult();
        }

        private void RenderFromResult()
        {
            string body;
            Icon   icon;
            bool   retryMode;

            switch (_result.Status)
            {
                case ColumnInstaller.InstallStatus.Installed:
                    body = $"✓ Bluebeam columns installed ({_result.TouchedFiles.Count} profile(s)).\r\n\r\n"
                         + "A .tabsbackup sidecar of the original was saved alongside each modified profile.";
                    icon = SystemIcons.Information;
                    retryMode = false;
                    break;

                case ColumnInstaller.InstallStatus.NotNeeded:
                    body = "✓ Bluebeam columns are already set up. No changes needed.";
                    icon = SystemIcons.Information;
                    retryMode = false;
                    break;

                case ColumnInstaller.InstallStatus.BluebeamRunning:
                    body = "⚠ Bluebeam Revu is currently running.\r\n\r\n"
                         + "Please close Bluebeam Revu completely (check the taskbar for multiple "
                         + "windows), then click Retry to complete the installation.";
                    icon = SystemIcons.Warning;
                    retryMode = true;
                    break;

                case ColumnInstaller.InstallStatus.NoProfileFound:
                    body = "⚠ No Bluebeam Revu 21, 2024, or 2025 profile was found.\r\n\r\n"
                         + "Launch Bluebeam once to let it create a default profile, then use the "
                         + "tray menu 'Install Bluebeam Columns...' to finish setup.";
                    icon = SystemIcons.Warning;
                    retryMode = false;
                    break;

                case ColumnInstaller.InstallStatus.ConflictDetected:
                    body = "⚠ Cannot install TABS columns automatically:\r\n\r\n"
                         + (_result.Message ?? "conflict detected")
                         + "\r\n\r\nContact support if you need help resolving the conflict.";
                    icon = SystemIcons.Warning;
                    retryMode = false;
                    break;

                default:
                    body = "⚠ Column setup: " + _result.Status + " — " + (_result.Message ?? "unknown");
                    icon = SystemIcons.Warning;
                    retryMode = false;
                    break;
            }

            ApplyMessage(icon, body, retryMode);
        }

        private void ApplyMessage(Icon icon, string body, bool retryMode)
        {
            _iconBox.Image = icon.ToBitmap();
            _messageLabel.Text = _preamble.Length == 0
                ? body
                : _preamble + "\r\n\r\n" + body;

            if (retryMode)
            {
                _primaryButton.Text      = "Retry";
                _primaryButton.Location  = _primaryWithSecondaryPos;
                _secondaryButton.Visible = true;
            }
            else
            {
                _primaryButton.Text      = "OK";
                _primaryButton.Location  = _primaryAlonePos;
                _secondaryButton.Visible = false;
            }

            _primaryButton.Enabled   = true;
            _secondaryButton.Enabled = true;
        }

        private async void OnPrimaryClick(object? sender, EventArgs e)
        {
            // Non-retry states → primary is an OK/Close button, just dismiss.
            if (_terminalError || _result.Status != ColumnInstaller.InstallStatus.BluebeamRunning)
            {
                DialogResult = DialogResult.OK;
                Close();
                return;
            }

            // Retry flow.
            _primaryButton.Enabled   = false;
            _secondaryButton.Enabled = false;
            UseWaitCursor            = true;
            _iconBox.Image           = SystemIcons.Information.ToBitmap();
            _messageLabel.Text       = _preamble.Length == 0
                ? "Installing column sets…"
                : _preamble + "\r\n\r\nInstalling column sets…";

            // Give Revu's on-exit config flush time to finish so we don't race its final write.
            await Task.Delay(500);

            try
            {
                _result = await Task.Run(ColumnInstaller.CheckAndInstall);
                if (IsDisposed) return;
                UseWaitCursor = false;
                RenderFromResult();
            }
            catch (Exception ex)
            {
                if (IsDisposed) return;
                UseWaitCursor = false;
                Debug.WriteLine("Column retry install threw: " + ex);
                _terminalError = true;
                ApplyMessage(
                    SystemIcons.Error,
                    "⚠ Unexpected error while installing columns:\r\n\r\n" + ex.Message,
                    retryMode: false);
            }
        }

        /// <summary>Utility: is Bluebeam Revu running right now?</summary>
        public static bool IsRevuRunning()
        {
            try { return Process.GetProcessesByName("Revu").Length > 0; }
            catch { return false; }
        }
    }
}
