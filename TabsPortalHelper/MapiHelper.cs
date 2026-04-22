using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

namespace TabsPortalHelper
{
    // ════════════════════════════════════════════════════════════════════════
    // Windows Simple MAPI — opens default mail client with compose window
    // pre-filled with recipients, subject, body, and file attachments.
    //
    // Works with: Outlook, Thunderbird, eM Client, Windows Mail, etc.
    // No authentication required — uses whatever the user's default app is.
    // MAPI_DIALOG flag ensures the user sees the compose window before sending.
    // ════════════════════════════════════════════════════════════════════════

    static class MapiHelper
    {
        const int MAPI_DIALOG            = 0x0008;
        const int MAPI_LOGON_UI          = 0x0001;
        const int MAPI_TO                = 1;
        const int MAPI_CC                = 2;
        const int MAPI_BCC               = 3;

        const int SUCCESS_SUCCESS              = 0;
        const int MAPI_E_USER_ABORT            = 1;
        const int MAPI_E_FAILURE               = 2;
        const int MAPI_E_INSUFFICIENT_MEMORY   = 5;

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        struct MapiMessage
        {
            public int    Reserved;
            public string Subject;
            public string NoteText;
            public string MessageType;
            public string DateReceived;
            public string ConversationID;
            public int    Flags;
            public IntPtr Originator;
            public int    RecipCount;
            public IntPtr Recips;
            public int    FileCount;
            public IntPtr Files;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        struct MapiRecipDesc
        {
            public int    Reserved;
            public int    RecipClass;
            public string Name;
            public string Address;
            public int    EIDSize;
            public IntPtr EntryID;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        struct MapiFileDesc
        {
            public int    Reserved;
            public int    Flags;
            public int    Position;
            public string PathName;
            public string FileName;
            public IntPtr FileType;
        }

        [DllImport("MAPI32.DLL", CharSet = CharSet.Ansi)]
        static extern int MAPISendMail(
            IntPtr lhSession,
            IntPtr ulUIParam,
            ref MapiMessage message,
            int flFlags,
            int ulReserved);

        public class ComposeRequest
        {
            public List<string> To        { get; set; } = new();
            public List<string> Cc        { get; set; } = new();
            public List<string> Bcc       { get; set; } = new();
            public string       Subject   { get; set; } = "";
            public string       Body      { get; set; } = "";
            public List<string> FilePaths { get; set; } = new();
        }

        public class ComposeResult
        {
            public bool    Success       { get; set; }
            public string? Error         { get; set; }
            public bool    UserCancelled { get; set; }
        }

        public static ComposeResult Compose(ComposeRequest request)
        {
            ComposeResult? result = null;
            Exception? threadEx = null;

            var thread = new Thread(() =>
            {
                try { result = ComposeMapi(request); }
                catch (Exception ex) { threadEx = ex; }
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

            if (threadEx != null)
                return new ComposeResult { Success = false, Error = threadEx.Message };

            return result ?? new ComposeResult { Success = false, Error = "Unknown error" };
        }

        static ComposeResult ComposeMapi(ComposeRequest request)
        {
            var allRecipients = new List<(string email, int recipClass)>();
            foreach (var email in request.To)  allRecipients.Add((email, MAPI_TO));
            foreach (var email in request.Cc)  allRecipients.Add((email, MAPI_CC));
            foreach (var email in request.Bcc) allRecipients.Add((email, MAPI_BCC));

            var recipDescs = new MapiRecipDesc[allRecipients.Count];
            for (int i = 0; i < allRecipients.Count; i++)
            {
                var (email, cls) = allRecipients[i];
                recipDescs[i] = new MapiRecipDesc
                {
                    RecipClass = cls,
                    Name       = email,
                    Address    = $"SMTP:{email}",
                };
            }

            var fileDescs = new MapiFileDesc[request.FilePaths.Count];
            for (int i = 0; i < request.FilePaths.Count; i++)
            {
                fileDescs[i] = new MapiFileDesc
                {
                    Position = -1,
                    PathName = request.FilePaths[i],
                    FileName = Path.GetFileName(request.FilePaths[i]),
                };
            }

            IntPtr recipPtr = IntPtr.Zero;
            IntPtr filePtr  = IntPtr.Zero;

            try
            {
                int recipSize = Marshal.SizeOf(typeof(MapiRecipDesc));
                int fileSize  = Marshal.SizeOf(typeof(MapiFileDesc));

                if (recipDescs.Length > 0)
                {
                    recipPtr = Marshal.AllocHGlobal(recipSize * recipDescs.Length);
                    for (int i = 0; i < recipDescs.Length; i++)
                        Marshal.StructureToPtr(recipDescs[i], recipPtr + i * recipSize, false);
                }

                if (fileDescs.Length > 0)
                {
                    filePtr = Marshal.AllocHGlobal(fileSize * fileDescs.Length);
                    for (int i = 0; i < fileDescs.Length; i++)
                        Marshal.StructureToPtr(fileDescs[i], filePtr + i * fileSize, false);
                }

                // Pass body straight through — line ending normalisation
                // is handled on the Flutter side before sending.
                var message = new MapiMessage
                {
                    Subject    = request.Subject,
                    NoteText   = request.Body,
                    RecipCount = recipDescs.Length,
                    Recips     = recipPtr,
                    FileCount  = fileDescs.Length,
                    Files      = filePtr,
                };

                int flags = MAPI_DIALOG | MAPI_LOGON_UI;
                int ret   = MAPISendMail(IntPtr.Zero, IntPtr.Zero, ref message, flags, 0);

                return ret switch
                {
                    SUCCESS_SUCCESS            => new ComposeResult { Success = true },
                    MAPI_E_USER_ABORT          => new ComposeResult { Success = true, UserCancelled = true },
                    MAPI_E_FAILURE             => new ComposeResult { Success = false, Error = "MAPI failure. Make sure a default mail app is configured." },
                    MAPI_E_INSUFFICIENT_MEMORY => new ComposeResult { Success = false, Error = "Insufficient memory." },
                    _                          => new ComposeResult { Success = false, Error = $"MAPI error code {ret}." }
                };
            }
            finally
            {
                if (recipPtr != IntPtr.Zero) Marshal.FreeHGlobal(recipPtr);
                if (filePtr  != IntPtr.Zero) Marshal.FreeHGlobal(filePtr);
            }
        }
    }
}
