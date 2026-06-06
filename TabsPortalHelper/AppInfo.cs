using System.Reflection;

namespace TabsPortalHelper
{
    /// <summary>
    /// Single source of truth for the app version. Derives from the assembly
    /// version stamped by &lt;Version&gt; in TabsPortalHelper.csproj, so the tray
    /// menu, the installer, and the Add/Remove Programs entry can never drift
    /// apart again. To release a new version, bump ONLY the .csproj &lt;Version&gt;.
    /// </summary>
    static class AppInfo
    {
        public static string Version { get; } = ComputeVersion();

        static string ComputeVersion()
        {
            var v = Assembly.GetExecutingAssembly().GetName().Version;
            return v == null ? "0.0.0" : $"{v.Major}.{v.Minor}.{v.Build}";
        }
    }
}
