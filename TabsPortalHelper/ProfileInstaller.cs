// ProfileInstaller.cs
//
// Idempotently installs the TABSportal Bluebeam Revu profile.
//
// Supersedes ColumnInstaller (v2.0–2.2), which edited the user's existing
// .bpx XML directly to graft in TABS columns. The new approach ships a
// complete TABSportal.bpx as an embedded resource and tells Revu.exe to
// import + activate it:
//
//     Revu.exe /s /bpximport:"<path>" /bpxactive:"TABSportal"
//
// Benefits over the old approach:
//   • Works whether Bluebeam is running (columns update live) or not
//     (Bluebeam launches momentarily to apply the change). No "close
//     Bluebeam first" retry-loop.
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

            // 3. Fire-and-forget: Revu.exe /s /bpximport:... /bpxactive:TABSportal.
            //    Works whether or not Bluebeam is already running. If running,
            //    the active profile switches live; if not, Revu launches and
            //    applies the profile before the user ever interacts with it.
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = revu,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                };
                // ArgumentList quotes each arg independently — safe with
                // colons, spaces, and '/' prefixes.
                psi.ArgumentList.Add("/s");
                psi.ArgumentList.Add("/bpximport:" + bpxPath);
                psi.ArgumentList.Add("/bpxactive:" + ProfileName);

                Process.Start(psi);
            }
            catch (Exception ex)
            {
                return new InstallResult
                {
                    Status = InstallStatus.Failed,
                    Message = "Couldn't launch Bluebeam Revu to import the profile: " + ex.Message,
                    RevuExePath = revu,
                    BpxPath = bpxPath,
                };
            }

            return new InstallResult
            {
                Status = InstallStatus.Installed,
                RevuExePath = revu,
                BpxPath = bpxPath,
            };
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
