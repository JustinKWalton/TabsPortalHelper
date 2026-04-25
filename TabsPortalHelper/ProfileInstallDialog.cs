using System;
using System.Drawing;
using System.Windows.Forms;

namespace TabsPortalHelper
{
    /// <summary>
    /// Post-install / tray-menu dialog for TABSportal profile installation.
    ///
    /// Much simpler than the v2.2 ColumnInstallDialog — no retry loop, because
    /// Revu.exe /s /bpximport /bpxactive works whether Bluebeam is running
    /// or not. Single OK button, three possible states.
    ///
    /// Called from:
    ///   • Installer.Install()               — preamble = install success text
    ///   • TrayApp.InstallBluebeamProfile()  — preamble = empty (compact size)
    /// </summary>
    public sealed class ProfileInstallDialog : Form
    {
        public ProfileInstallDialog(
            string windowTitle,
            string preamble,
            ProfileInstaller.InstallResult result)
        {
            preamble ??= string.Empty;
            bool hasPreamble = preamble.Length > 0;

            Text            = windowTitle;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox     = false;
            MinimizeBox     = false;
            ShowInTaskbar   = false;
            StartPosition   = FormStartPosition.CenterScreen;
            ClientSize      = hasPreamble ? new Size(520, 340) : new Size(480, 220);

            const int Pad  = 16;
            const int BtnW = 100;
            const int BtnH = 28;
            int btnY = ClientSize.Height - BtnH - Pad;

            var iconBox = new PictureBox
            {
                Location = new Point(Pad, 20),
                Size     = new Size(32, 32),
                SizeMode = PictureBoxSizeMode.StretchImage,
                Image    = IconFor(result.Status).ToBitmap(),
            };

            var label = new Label
            {
                Location = new Point(60, 16),
                Size     = new Size(ClientSize.Width - 60 - Pad, btnY - 36),
                AutoSize = false,
                Text     = BuildMessage(preamble, result),
            };

            var okButton = new Button
            {
                Text         = "OK",
                Size         = new Size(BtnW, BtnH),
                Location     = new Point(ClientSize.Width - Pad - BtnW, btnY),
                DialogResult = DialogResult.OK,
            };

            AcceptButton = okButton;
            CancelButton = okButton;

            Controls.Add(iconBox);
            Controls.Add(label);
            Controls.Add(okButton);
        }

        private static Icon IconFor(ProfileInstaller.InstallStatus status) => status switch
        {
            ProfileInstaller.InstallStatus.Installed => SystemIcons.Information,
            _                                        => SystemIcons.Warning,
        };

        private static string BuildMessage(string preamble, ProfileInstaller.InstallResult r)
        {
            string body = r.Status switch
            {
                ProfileInstaller.InstallStatus.Installed =>
                    "\u2713 TABSportal Bluebeam profile installed.\r\n\r\n" +
                    "TABSportal is now your active Bluebeam profile and Bluebeam " +
                    "is open. Your existing profiles are untouched \u2014 " +
                    "TABSportal is added alongside them, and you can switch back " +
                    "any time from Revu \u2192 Profiles.",

                ProfileInstaller.InstallStatus.NoRevuFound =>
                    "\u26A0 Bluebeam Revu was not found on this computer.\r\n\r\n" +
                    "Install Bluebeam Revu, then use the tray menu " +
                    "\u201CInstall Bluebeam Profile\u2026\u201D to finish setup.",

                ProfileInstaller.InstallStatus.Failed =>
                    "\u26A0 Couldn't install the TABSportal profile:\r\n\r\n" +
                    (r.Message ?? "unknown error"),

                _ =>
                    "\u26A0 Profile setup: " + r.Status + " \u2014 " + (r.Message ?? ""),
            };

            return preamble.Length == 0 ? body : preamble + "\r\n\r\n" + body;
        }
    }
}
