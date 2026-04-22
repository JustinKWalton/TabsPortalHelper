using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TabsPortalHelper
{
    static class BluebeamHelper
    {
        static readonly string[] KnownPaths =
        {
            @"C:\Program Files\Bluebeam Software\Bluebeam Revu\21\Revu\Revu.exe",
            @"C:\Program Files\Bluebeam Software\Bluebeam Revu\2019\Revu\Revu.exe",
            @"C:\Program Files\Bluebeam Software\Bluebeam Revu\20\Revu\Revu.exe",
            @"C:\Program Files (x86)\Bluebeam Software\Bluebeam Revu\21\Revu\Revu.exe",
            @"C:\Program Files (x86)\Bluebeam Software\Bluebeam Revu\2019\Revu\Revu.exe",
            @"C:\Program Files (x86)\Bluebeam Software\Bluebeam Revu\20\Revu\Revu.exe",
        };

        const string RevuProcessName = "Revu";

        // ─── Win32 for foreground focus ──────────────────────────────────────
        const int  SW_RESTORE        = 9;
        const int  SW_SHOW           = 5;
        const byte VK_MENU           = 0x12;   // Alt
        const uint KEYEVENTF_KEYUP   = 0x0002;

        [DllImport("user32.dll")]   static extern bool   SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")]   static extern bool   ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")]   static extern bool   IsIconic(IntPtr hWnd);
        [DllImport("user32.dll")]   static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]   static extern uint   GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
        [DllImport("user32.dll")]   static extern bool   BringWindowToTop(IntPtr hWnd);
        [DllImport("user32.dll")]   static extern void   keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
        [DllImport("kernel32.dll")] static extern uint   GetCurrentThreadId();

        public static string? FindBluebeam() =>
            KnownPaths.FirstOrDefault(File.Exists);

        /// <summary>
        /// Opens the file in Bluebeam if available, otherwise prompts for default app.
        /// If Bluebeam is already running, the existing instance is brought to the
        /// foreground after the file-open command is dispatched.
        /// Returns true if opened in Bluebeam, false if opened in default app.
        /// </summary>
        public static bool OpenFile(string filePath)
        {
            var bluebeamExe = FindBluebeam();

            if (bluebeamExe != null)
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = bluebeamExe,
                    Arguments = $"\"{filePath}\"",
                    UseShellExecute = false
                });

                // Fire-and-forget: let the HTTP handler return immediately
                // while we wait for Revu to process the command-line arg
                // and then force its window to the foreground.
                Task.Run(async () =>
                {
                    await Task.Delay(300);
                    BringRevuToFront(timeoutMs: 3000);
                });

                return true;
            }

            // Bluebeam not found — offer fallback
            var result = MessageBox.Show(
                $"Bluebeam Revu was not found on this computer.\n\n" +
                $"Open with the Windows default app instead?\n\n{filePath}",
                "TABS Portal Helper",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                });
            }

            return false;
        }

        // ─── Foreground helpers ──────────────────────────────────────────────

        static bool BringRevuToFront(int timeoutMs)
        {
            var deadline = Environment.TickCount + timeoutMs;
            while (Environment.TickCount < deadline)
            {
                foreach (var p in Process.GetProcessesByName(RevuProcessName))
                {
                    try
                    {
                        p.Refresh();
                        var hWnd = p.MainWindowHandle;
                        if (hWnd != IntPtr.Zero)
                        {
                            ForceForeground(hWnd);
                            return true;
                        }
                    }
                    catch { /* process exited mid-iteration */ }
                    finally { p.Dispose(); }
                }
                Thread.Sleep(100);
            }
            return false;
        }

        static void ForceForeground(IntPtr hWnd)
        {
            if (IsIconic(hWnd)) ShowWindow(hWnd, SW_RESTORE);

            // Windows blocks SetForegroundWindow from background processes
            // (a tray app responding to HTTP has no foreground rights, so
            // AllowSetForegroundWindow is a no-op in this context).
            // The workaround: synthesise an Alt key press on our own thread.
            // Windows treats the thread that most recently received input as
            // eligible to change the foreground, so after these two
            // keybd_event calls SetForegroundWindow actually works instead
            // of getting demoted to a taskbar flash.
            keybd_event(VK_MENU, 0, 0,               UIntPtr.Zero);
            keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);

            BringWindowToTop(hWnd);
            ShowWindow(hWnd, SW_SHOW);
            SetForegroundWindow(hWnd);
        }
    }
}