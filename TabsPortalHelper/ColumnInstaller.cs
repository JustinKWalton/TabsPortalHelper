// ColumnInstaller.cs
//
// Idempotently installs the TABS canonical custom columns (Note, Comment Status,
// Element, Location) into the user's Bluebeam Revu profile .bpx file(s).
//
// Wire-up in Program.cs / ApplicationContext:
//
//     const string COLUMN_SCHEMA_VERSION = "1.0.0";  // bump when canonical schema changes
//     var cfg = HelperConfig.Load();
//     if (cfg.LastColumnInstallVersion != COLUMN_SCHEMA_VERSION)
//     {
//         var r = ColumnInstaller.CheckAndInstall();
//         switch (r.Status)
//         {
//             case ColumnInstaller.InstallStatus.NotNeeded:
//             case ColumnInstaller.InstallStatus.Installed:
//             case ColumnInstaller.InstallStatus.NoProfileFound:
//                 cfg.LastColumnInstallVersion = COLUMN_SCHEMA_VERSION;
//                 cfg.Save();
//                 if (r.Status == ColumnInstaller.InstallStatus.Installed)
//                     TrayIcon.ShowToast("TABS columns installed in Bluebeam.");
//                 break;
//             case ColumnInstaller.InstallStatus.BluebeamRunning:
//                 TrayIcon.ShowToast("Close Bluebeam to set up TABS columns.");
//                 // do NOT update cfg — retry next helper startup
//                 break;
//             default:
//                 Log.Error($"Column install failed: {r.Message}");
//                 break;
//         }
//     }

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace TabsPortalHelper
{
    public static class ColumnInstaller
    {
        // ---- Canonical TABS column schema -------------------------------------
        // Index values matter: they are the BSIColumnData slots used by the
        // clipboard payload. DO NOT reorder without updating the clipboard writer.
        private static readonly ColumnSpec[] TabsColumns =
        {
            new ColumnSpec(0, "Note",           "Text",   776),
            new ColumnSpec(1, "Comment Status", "Choice", 184,
                           choiceOptions: new[] { "Not Acceptable", "Could Not Verify", "Area of Concern" },
                           choiceDefault: "Not Acceptable"),
            new ColumnSpec(2, "Element",        "Text",   216),
            new ColumnSpec(3, "Location",       "Text",   248),
        };

        private static readonly string[] SupportedRevuVersions = { "21", "2024", "2025" };

        // ---- Public API -------------------------------------------------------

        public enum InstallStatus
        {
            NotNeeded,        // columns already present in all found profiles
            Installed,        // columns were added to at least one profile
            BluebeamRunning,  // Revu.exe is running; cannot safely edit .bpx
            NoProfileFound,   // no supported Revu version directories exist
            ConflictDetected, // existing column blocks our canonical index
            Failed            // I/O or parse error
        }

        public sealed class InstallResult
        {
            public InstallStatus Status { get; set; }
            public string Message { get; set; }
            public List<string> TouchedFiles { get; } = new List<string>();
        }

        public static InstallResult CheckAndInstall()
        {
            if (IsBluebeamRunning())
            {
                return new InstallResult
                {
                    Status = InstallStatus.BluebeamRunning,
                    Message = "Bluebeam Revu is running. Close it to set up TABS columns."
                };
            }

            var appdata = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var baseDir = Path.Combine(appdata, "Bluebeam Software", "Revu");

            if (!Directory.Exists(baseDir))
            {
                return new InstallResult
                {
                    Status = InstallStatus.NoProfileFound,
                    Message = $"Not found: {baseDir}"
                };
            }

            var bpxPaths = SupportedRevuVersions
                .Select(v => Path.Combine(baseDir, v, "Revu.bpx"))
                .Where(File.Exists)
                .ToList();

            if (bpxPaths.Count == 0)
            {
                return new InstallResult
                {
                    Status = InstallStatus.NoProfileFound,
                    Message = $"No Revu.bpx found in {string.Join(", ", SupportedRevuVersions)}"
                };
            }

            var result = new InstallResult { Status = InstallStatus.NotNeeded };
            foreach (var bpx in bpxPaths)
            {
                try
                {
                    var changed = InstallIntoProfile(bpx, out var conflictMsg);
                    if (conflictMsg != null)
                    {
                        result.Status = InstallStatus.ConflictDetected;
                        result.Message = $"{bpx}: {conflictMsg}";
                        return result;
                    }
                    if (changed)
                    {
                        result.Status = InstallStatus.Installed;
                        result.TouchedFiles.Add(bpx);
                    }
                }
                catch (Exception ex)
                {
                    result.Status = InstallStatus.Failed;
                    result.Message = $"{bpx}: {ex.Message}";
                    return result;
                }
            }
            return result;
        }

        // ---- Implementation ---------------------------------------------------

        private static bool IsBluebeamRunning()
        {
            try
            {
                return Process.GetProcesses().Any(p =>
                    SafeName(p).IndexOf("Revu", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    SafeName(p).IndexOf("Bluebeam", StringComparison.OrdinalIgnoreCase) >= 0);
            }
            catch
            {
                return false; // don't block install on a permission error
            }
        }

        private static string SafeName(Process p)
        {
            try { return p.ProcessName ?? ""; } catch { return ""; }
        }

        private static bool InstallIntoProfile(string bpxPath, out string conflictMessage)
        {
            conflictMessage = null;

            var doc = new XmlDocument { PreserveWhitespace = false };
            doc.Load(bpxPath);

            var markupList = doc.SelectSingleNode("//Record[@Key='MarkupList']");
            if (markupList == null)
                throw new InvalidDataException("MarkupList record not found in profile");

            var customColumns = markupList.SelectSingleNode("CustomColumns")
                                ?? markupList.AppendChild(doc.CreateElement("CustomColumns"));

            var columnsBlock = markupList.SelectSingleNode("Columns")
                               ?? markupList.AppendChild(doc.CreateElement("Columns"));

            var displayOrder = markupList.SelectSingleNode("DisplayOrder")
                               ?? markupList.AppendChild(doc.CreateElement("DisplayOrder"));

            // Map existing <BSIColumnItem> by name (case-insensitive) and by Index.
            var existingByName = new Dictionary<string, XmlNode>(StringComparer.OrdinalIgnoreCase);
            var existingByIndex = new Dictionary<int, XmlNode>();
            foreach (XmlNode n in customColumns.SelectNodes("BSIColumnItem"))
            {
                var name = n.SelectSingleNode("n")?.InnerText ?? "";
                if (!string.IsNullOrEmpty(name)) existingByName[name] = n;

                if (int.TryParse(n.Attributes?["Index"]?.Value, out var idx))
                    existingByIndex[idx] = n;
            }

            // Short-circuit: all canonical columns already exist.
            if (TabsColumns.All(c => existingByName.ContainsKey(c.Name)))
                return false;

            bool changed = false;
            foreach (var spec in TabsColumns)
            {
                if (existingByName.ContainsKey(spec.Name))
                    continue;

                // Refuse to clobber a different column that already occupies our slot.
                if (existingByIndex.TryGetValue(spec.Index, out var occupant))
                {
                    var occName = occupant.SelectSingleNode("n")?.InnerText ?? "(unnamed)";
                    conflictMessage = $"Index {spec.Index} is occupied by existing column '{occName}' — refusing to overwrite to add '{spec.Name}'";
                    return false;
                }

                // Append the <BSIColumnItem>.
                customColumns.AppendChild(BuildColumnItem(doc, spec));

                // Ensure the matching <Column Key="UserDefinedN"> entry exists and is visible.
                EnsureUserDefinedColumn(doc, columnsBlock, spec);

                // Ensure the column appears in <DisplayOrder>.
                EnsureDisplayOrderEntry(doc, displayOrder, $"UserDefined{spec.Index}");

                changed = true;
            }

            if (changed)
            {
                // Write atomically: temp file, then move. Preserves file on failure.
                var tmp = bpxPath + ".tabstmp";
                var settings = new XmlWriterSettings
                {
                    Indent = true,
                    IndentChars = "  ",
                    Encoding = new UTF8Encoding(false), // no BOM — Bluebeam doesn't emit one
                    OmitXmlDeclaration = false
                };
                using (var writer = XmlWriter.Create(tmp, settings))
                {
                    doc.Save(writer);
                }
                // Backup once (first time we touch the file) then replace.
                var backup = bpxPath + ".tabsbackup";
                if (!File.Exists(backup))
                    File.Copy(bpxPath, backup, false);
                File.Delete(bpxPath);
                File.Move(tmp, bpxPath);
            }
            return changed;
        }

        private static XmlElement BuildColumnItem(XmlDocument doc, ColumnSpec spec)
        {
            var item = doc.CreateElement("BSIColumnItem");
            item.SetAttribute("Index", spec.Index.ToString());
            item.SetAttribute("Subtype", spec.Subtype);

            AppendText(doc, item, "n", spec.Name);
            AppendText(doc, item, "DisplayOrder", spec.Index.ToString());
            AppendText(doc, item, "Deleted", "False");

            if (spec.Subtype == "Text")
            {
                AppendText(doc, item, "Multiline", "False");
            }
            else if (spec.Subtype == "Choice")
            {
                AppendText(doc, item, "DefaultValue", spec.ChoiceDefault ?? "");
                AppendText(doc, item, "Format", "Normal");
                AppendText(doc, item, "Precision", "2");

                var items = doc.CreateElement("Items");
                for (int i = 0; i < spec.ChoiceOptions.Length; i++)
                {
                    var opt = doc.CreateElement("Item");
                    opt.SetAttribute("Index", i.ToString());
                    AppendText(doc, opt, "Value", spec.ChoiceOptions[i]);
                    if (string.Equals(spec.ChoiceOptions[i], spec.ChoiceDefault, StringComparison.Ordinal))
                        AppendText(doc, opt, "Default", "True");
                    items.AppendChild(opt);
                }
                item.AppendChild(items);

                AppendText(doc, item, "AllowCustom", "False");
            }
            return item;
        }

        private static void EnsureUserDefinedColumn(XmlDocument doc, XmlNode columnsBlock, ColumnSpec spec)
        {
            var key = $"UserDefined{spec.Index}";
            var existing = columnsBlock.SelectNodes("Column")
                .Cast<XmlNode>()
                .FirstOrDefault(n => n.Attributes?["Key"]?.Value == key);

            if (existing == null)
            {
                var col = doc.CreateElement("Column");
                col.SetAttribute("Key", key);
                AppendText(doc, col, "Width", spec.Width.ToString());
                AppendText(doc, col, "Visible", "True");
                AppendText(doc, col, "Added", "True");
                columnsBlock.AppendChild(col);
            }
            else
            {
                SetOrAppendText(doc, existing, "Visible", "True");
                SetOrAppendText(doc, existing, "Added", "True");
                // Don't touch Width if user customized it.
                if (existing.SelectSingleNode("Width") == null)
                    AppendText(doc, existing, "Width", spec.Width.ToString());
            }
        }

        private static void EnsureDisplayOrderEntry(XmlDocument doc, XmlNode displayOrder, string columnKey)
        {
            var present = displayOrder.SelectNodes("Column")
                .Cast<XmlNode>()
                .Any(n => n.Attributes?["Key"]?.Value == columnKey);
            if (present) return;

            var col = doc.CreateElement("Column");
            col.SetAttribute("Key", columnKey);
            displayOrder.AppendChild(col);
        }

        private static void AppendText(XmlDocument doc, XmlNode parent, string elementName, string value)
        {
            var el = doc.CreateElement(elementName);
            el.InnerText = value;
            parent.AppendChild(el);
        }

        private static void SetOrAppendText(XmlDocument doc, XmlNode parent, string elementName, string value)
        {
            var existing = parent.SelectSingleNode(elementName);
            if (existing != null) existing.InnerText = value;
            else AppendText(doc, parent, elementName, value);
        }

        // ---- Inner types ------------------------------------------------------

        private sealed class ColumnSpec
        {
            public int Index { get; }
            public string Name { get; }
            public string Subtype { get; }   // "Text" or "Choice"
            public int Width { get; }
            public string[] ChoiceOptions { get; }
            public string ChoiceDefault { get; }

            public ColumnSpec(int index, string name, string subtype, int width,
                              string[] choiceOptions = null, string choiceDefault = null)
            {
                Index = index;
                Name = name;
                Subtype = subtype;
                Width = width;
                ChoiceOptions = choiceOptions ?? Array.Empty<string>();
                ChoiceDefault = choiceDefault;
            }
        }
    }
}
