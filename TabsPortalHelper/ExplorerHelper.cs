using System;
using System.Diagnostics;
using System.IO;

namespace TabsPortalHelper
{
    static class ExplorerHelper
    {
        // ============================================================
        // OpenFolder
        //
        // Opens a Windows Explorer window at folderPath and brings it
        // to the foreground.
        //
        // Implementation note: rather than calling Process.Start +
        // SetForegroundWindow directly from this long-running tray
        // process (which Windows aggressively blocks from stealing
        // focus when called from a busy browser), we shell out to a
        // hidden powershell.exe. The PS process is *freshly spawned*
        // by us, so for its first ~1s of life Windows grants it
        // foreground-change rights that the tray app does not have.
        // Inside PS we use SendKeys '%' (synthetic Alt) to top off
        // the eligibility heuristic, then VB's AppActivate to bring
        // the Explorer window with matching title to the front.
        //
        // Why this works when raw Win32 didn't: the foreground-grab
        // restrictions that bit us across v2.5.0..2.5.3 apply per
        // *process*. A new process gets a clean slate. Shelling out
        // is essentially the "proxy process" pattern made cheap.
        // ============================================================
        public static bool OpenFolder(string folderPath)
        {
            try
            {
                var folderName = Path.GetFileName(
                    folderPath.TrimEnd(Path.DirectorySeparatorChar));

                // Escape single quotes for PS single-quoted strings.
                var pathArg = folderPath.Replace("'", "''");
                var nameArg = folderName.Replace("'", "''");

                // One-line PowerShell:
                //   1. Start-Process Explorer at the target folder.
                //   2. Load Forms + VB assemblies.
                //   3. Loop up to ~1.2s polling the window into focus:
                //        - SendKeys '%' synthesizes Alt for the PS
                //          process (which is briefly foreground-eligible
                //          as a freshly spawned process).
                //        - AppActivate finds the window by partial title
                //          match and activates it. Throws if not yet
                //          present, so we retry.
                var script = string.Join("; ", new[]
                {
                    $"Start-Process explorer.exe -ArgumentList '\"{pathArg}\"'",
                    "Add-Type -AssemblyName System.Windows.Forms",
                    "Add-Type -AssemblyName Microsoft.VisualBasic",
                    $"$folderName = '{nameArg}'",
                    "for ($i = 0; $i -lt 8; $i++) {" +
                    "  Start-Sleep -Milliseconds 150;" +
                    "  [System.Windows.Forms.SendKeys]::SendWait('%');" +
                    "  try {" +
                    "    [Microsoft.VisualBasic.Interaction]::AppActivate($folderName);" +
                    "    break" +
                    "  } catch {}" +
                    "}"
                });

                Process.Start(new ProcessStartInfo
                {
                    FileName  = "powershell.exe",
                    Arguments = $"-NoProfile -WindowStyle Hidden -Command \"{script}\"",
                    UseShellExecute = false,
                    CreateNoWindow  = true,
                });
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
