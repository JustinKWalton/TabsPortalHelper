// ProfileInstaller.cs
//
// Idempotently installs the TABSportal Bluebeam Revu profile.
//
// Supersedes ColumnInstaller (v2.0–2.2), which edited the user's existing
// .bpx XML directly to graft in TABS columns. The new approach ships a
// complete TABSportal.bpx as an embedded resource and tells Revu.exe to
// import + activate it via:
//
//     Revu.exe /s /bpximport:"<path>" /bpxactive:"TABSportal"
//
// Launch sequence:
//
//   • Bluebeam already running:
//       Just fire /bpximport. It hands off to the running instance,
//       switches the active profile live. No flicker.
//
//   • Bluebeam not running:
//       1. Launch Revu.exe visibly with no args.
//       2. Poll until its main window has a non-empty title — Bluebeam
//          updates its title from "" to "Revu" once the UI is fully
//          ready to accept commands. /bpximport against a not-yet-ready
//          instance falls through to "do silent import + exit" mode
//          and kills the new process.
//       3. Once truly ready, fire /bpximport. Hands off to the running
//          instance, switches the active profile live. No flicker.
//
//   • Bluebeam not running, but ready-poll times out:
//       Fall back to two-step launch: silent /bpximport (kills the
//       in-flight Bluebeam if any), then launch Bluebeam normally so
//       the user lands in a working window. User sees a brief
//       open/close/reopen flicker, but the end state is correct.
//
// Benefits over the old approach (v2.0–2.2 ColumnInstaller):
//   • Works whether Bluebeam is running or not. No "close Bluebeam
//     first" retry-loop.
//   • Controls the full Markups List view, not just the column set —
//     only TABSportal columns show, nothing else clutters the list.
//   • Zero risk of corrupting user profiles, since we never parse or
//     rewrite their XML.
//   • Profile can be iterated and redeployed by shipping a new helper
//     version (or later, by fetching from Supabase).

using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading;

namespace TabsPortalHelper
{
    public static class ProfileInstaller
    {
        /// <summary>
        /// Profile name inside the .bpx — must match the Name attribute on
        /// the &lt;RevuProfile&gt; root. Used with the /bpxactive switch.
        /// </summary>
        public const string ProfileName = "TABSportal";

        /// <summary>
        /// Embedded resource name — must match the &lt;LogicalName&gt; in .csproj.
        /// </summary>
        private const string EmbeddedResourceName = "TabsPortalHelper.TABSportal.bpx";

        /// <summary>
        /// Max time we'll wait for /bpximport to complete before returning
        /// to the caller. Normally sub-second; 15s is a generous cap.
        /// </summary>
        private const int ImportTimeoutMs = 15_000;

        /// <summary>
        /// Max time we'll wait for a freshly-launched Bluebeam to become
        /// "ready" (main window visible with a non-empty title) before
        /// falling back to the two-step launch path.
        /// </summary>
        private const int BluebeamReadyTimeoutMs = 30_000;

        /// <summary>
        /// How often we poll the launched Bluebeam process for readiness.
        /// </summary>
        private const int BluebeamReadyPollMs = 250;

        /// <summary>
        /// Extra grace after we detect a non-empty MainWindowTitle. Bluebeam
        /// keeps loading toolset / profile state for a beat after the title
        /// appears, and /bpximport can race that work if we fire too soon.
        /// </summary>
        private const int BluebeamReadyBufferMs = 1_500;

        public enum InstallStatus
        {
            Installed,     // Revu.exe was told to import + activate the profile
            NoRevuFound,   // Bluebeam Revu is not installed on this machine
            Failed,        // Extraction or process launch failed
        }

        public sealed class InstallResult
        {
            public InstallStatus Status { get; set; }
            public string? Message { get; set; }
            public string? RevuExePath { get; set; }
            public string? BpxPath { get; set; }
        }

        public static InstallResult CheckAndInstall()
        {
            // 1. Locate Revu.exe.
            var revu = BluebeamHelper.FindBluebeam();
            if (revu == null)
            {
                return new InstallResult
                {
                    Status = InstallStatus.NoRevuFound,
                    Message = "Bluebeam Revu was not found on this computer."
                };
            }

            // 2. Extract the bundled profile to a stable path the helper owns.
            string bpxPath;
            try
            {
                bpxPath = ExtractBundledProfile();
            }
            catch (Exception ex)
            {
                return new InstallResult
                {
                    Status = InstallStatus.Failed,
                    Message = "Couldn't extract the bundled TABSportal profile: " + ex.Message,
                    RevuExePath = revu,
                };
            }

            // 3. Decide whether we need to launch Bluebeam ourselves and
            //    wait for it to be ready before firing /bpximport.
            bool wasRunning = IsBluebeamRunning();
            bool readyForLiveImport = wasRunning;

            try
            {
                if (!wasRunning)
                {
                    // Cold start: launch Bluebeam visibly so the import has
                    // a running instance to hand off to.
                    var openPsi = new ProcessStartInfo
                    {
                        FileName        = revu,
                        UseShellExecute = false,
                    };
                    var openProc = Process.Start(openPsi);

                    // WaitForInputIdle gets us past the very first message-
                    // pump idle, which is well before Bluebeam's UI is
                    // actually responsive to /bpximport. So we additionally
                    // poll until MainWindowTitle is non-empty, which is
                    // Bluebeam's signal that the main UI is up.
                    try
                    {
                        openProc?.WaitForInputIdle(BluebeamReadyTimeoutMs);
                    }
                    catch
                    {
                        // WaitForInputIdle can throw on unsupported process
                        // types. Harmless — the polling loop below picks up
                        // the slack.
                    }

                    readyForLiveImport = WaitForBluebeamReady(BluebeamReadyTimeoutMs);

                    if (readyForLiveImport)
                    {
                        // Buffer to let Bluebeam finish loading toolsets
                        // and profile state after the window appears.
                        Thread.Sleep(BluebeamReadyBufferMs);
                    }
                }

                // 4. /bpximport behavior depends on whether a fully-loaded
                //    Bluebeam instance is currently running:
                //      • Yes  → live profile switch, no flicker.
                //      • No   → silent import process exits, taking any
                //               racing Bluebeam with it. We then launch
                //               Bluebeam visibly as a fallback.
                var importPsi = new ProcessStartInfo
                {
                    FileName        = revu,
                    UseShellExecute = false,
                    CreateNoWindow  = true,
                };
                // ArgumentList quotes each arg independently — safe with
                // colons, spaces, and '/' prefixes.
                importPsi.ArgumentList.Add("/s");
                importPsi.ArgumentList.Add("/bpximport:" + bpxPath);
                importPsi.ArgumentList.Add("/bpxactive:" + ProfileName);

                using (var importProc = Process.Start(importPsi))
                {
                    importProc?.WaitForExit(ImportTimeoutMs);
                }

                // 5. Fallback: if Bluebeam wasn't ready in time, the
                //    /bpximport call almost certainly killed the in-flight
                //    instance. Launch Bluebeam normally so the user lands
                //    in a working window with the profile already active.
                if (!wasRunning && !readyForLiveImport)
                {
                    var openAgain = new ProcessStartInfo
                    {
                        FileName        = revu,
                        UseShellExecute = false,
                    };
                    Process.Start(openAgain);
                }
            }
            catch (Exception ex)
            {
                return new InstallResult
                {
                    Status      = InstallStatus.Failed,
                    Message     = "Couldn't launch Bluebeam Revu to import the profile: " + ex.Message,
                    RevuExePath = revu,
                    BpxPath     = bpxPath,
                };
            }

            return new InstallResult
            {
                Status      = InstallStatus.Installed,
                RevuExePath = revu,
                BpxPath     = bpxPath,
            };
        }

        /// <summary>
        /// Returns true if at least one Bluebeam Revu process is currently
        /// running on this machine. Doesn't say anything about whether the
        /// running instance is ready to accept commands.
        /// </summary>
        private static bool IsBluebeamRunning()
        {
            try
            {
                return Process.GetProcessesByName("Revu").Length > 0;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Polls until at least one Revu process has a non-empty
        /// MainWindowTitle (signaling the UI is fully up and able to
        /// receive /bpximport), or the timeout elapses.
        /// </summary>
        /// <returns>true if Bluebeam became ready; false on timeout.</returns>
        private static bool WaitForBluebeamReady(int timeoutMs)
        {
            var deadline = Environment.TickCount + timeoutMs;
            while (Environment.TickCount < deadline)
            {
                try
                {
                    var procs = Process.GetProcessesByName("Revu");
                    foreach (var p in procs)
                    {
                        try
                        {
                            // p.MainWindowTitle is computed each access; we
                            // need to refresh to pick up changes.
                            p.Refresh();
                            if (!string.IsNullOrEmpty(p.MainWindowTitle))
                                return true;
                        }
                        catch
                        {
                            // Process may have exited mid-loop; ignore.
                        }
                    }
                }
                catch
                {
                    // Process enumeration failed; retry next tick.
                }

                Thread.Sleep(BluebeamReadyPollMs);
            }

            return false;
        }

        /// <summary>
        /// Extracts the embedded TABSportal.bpx into
        /// %LOCALAPPDATA%\TabsPortalHelper\profile\TABSportal.bpx,
        /// overwriting any previous copy. Returns the absolute path.
        /// </summary>
        public static string ExtractBundledProfile()
        {
            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "TabsPortalHelper",
                "profile");
            Directory.CreateDirectory(dir);

            var path = Path.Combine(dir, "TABSportal.bpx");

            var asm = Assembly.GetExecutingAssembly();
            using var src = asm.GetManifestResourceStream(EmbeddedResourceName)
                ?? throw new InvalidOperationException(
                    $"Embedded resource '{EmbeddedResourceName}' not found. " +
                    "Add TABSportal.bpx to .csproj as <EmbeddedResource> with " +
                    $"LogicalName='{EmbeddedResourceName}'.");
            using var dst = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None);
            src.CopyTo(dst);

            return path;
        }
    }
}
