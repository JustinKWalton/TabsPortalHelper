using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace TabsPortalHelper
{
    class HttpServer
    {
        private readonly int _port;
        private HttpListener? _listener;
        private CancellationTokenSource? _cts;
        const string Version = "2.1.0";

        public HttpServer(int port)
        {
            _port = port;
        }

        public void Start()
        {
            _cts = new CancellationTokenSource();
            _listener = new HttpListener();
            _listener.Prefixes.Add($"http://localhost:{_port}/");
            _listener.Start();

            Task.Run(() => ListenLoop(_cts.Token));
        }

        public void Stop()
        {
            _cts?.Cancel();
            _listener?.Stop();
        }

        async Task ListenLoop(CancellationToken ct)
        {
            while (!ct.IsCancellationRequested)
            {
                try
                {
                    var context = await _listener!.GetContextAsync();
                    _ = Task.Run(() => HandleRequest(context), ct);
                }
                catch (HttpListenerException) { break; }
                catch (ObjectDisposedException) { break; }
                catch { /* keep running */ }
            }
        }

        void HandleRequest(HttpListenerContext ctx)
        {
            try
            {
                ctx.Response.Headers["Access-Control-Allow-Origin"] = "*";
                ctx.Response.Headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS";
                ctx.Response.Headers["Access-Control-Allow-Headers"] = "Content-Type";

                if (ctx.Request.HttpMethod == "OPTIONS")
                {
                    ctx.Response.StatusCode = 204;
                    ctx.Response.Close();
                    return;
                }

                var path = ctx.Request.Url?.AbsolutePath.TrimEnd('/').ToLowerInvariant() ?? "/";
                var query = ctx.Request.QueryString;

                switch (path)
                {
                    case "/ping":                     HandlePing(ctx);                break;
                    case "/open":                     HandleOpen(ctx, query);         break;
                    case "/file":                     HandleFile(ctx, query);         break;
                    case "/files":                    HandleFiles(ctx, query);        break;
                    case "/compose":                  HandleCompose(ctx);             break;
                    case "/clipboard/bluebeam-markup": HandleClipboardMarkup(ctx);    break;
                    default:
                        WriteJson(ctx, 404, new { error = "Unknown endpoint", path });
                        break;
                }
            }
            catch (Exception ex)
            {
                try { WriteJson(ctx, 500, new { error = ex.Message }); } catch { }
            }
        }

        void HandlePing(HttpListenerContext ctx)
        {
            WriteJson(ctx, 200, new
            {
                status = "ok",
                version = Version,
                driveRoot = DriveHelper.FindDriveRoot(),
                bluebeam = BluebeamHelper.FindBluebeam()
            });
        }

        void HandleOpen(HttpListenerContext ctx, System.Collections.Specialized.NameValueCollection query)
        {
            var fileId = query["fileId"];
            if (string.IsNullOrWhiteSpace(fileId))
            {
                WriteJson(ctx, 400, new { error = "fileId parameter required" });
                return;
            }

            var filePath = DriveHelper.FindLocalPathByFileId(fileId);
            if (filePath == null)
            {
                WriteJson(ctx, 404, new
                {
                    error = "File not found locally. Make sure it is available offline in Google Drive.",
                    fileId
                });
                return;
            }

            if (!File.Exists(filePath))
            {
                WriteJson(ctx, 404, new
                {
                    error = "File not found on disk. It may have been moved or deleted.",
                    filePath
                });
                return;
            }

            var launched = BluebeamHelper.OpenFile(filePath);
            WriteJson(ctx, 200, new { success = true, filePath, openedInBluebeam = launched });
        }

        void HandleFile(HttpListenerContext ctx, System.Collections.Specialized.NameValueCollection query)
        {
            var fileId = query["fileId"];
            if (string.IsNullOrWhiteSpace(fileId))
            {
                WriteJson(ctx, 400, new { error = "fileId parameter required" });
                return;
            }

            var filePath = DriveHelper.FindLocalPathByFileId(fileId);
            if (filePath == null || !File.Exists(filePath))
            {
                WriteJson(ctx, 404, new { error = "File not found locally.", fileId });
                return;
            }

            var bytes = File.ReadAllBytes(filePath);
            var fileName = Path.GetFileName(filePath);
            var contentType = GetContentType(filePath);

            ctx.Response.StatusCode = 200;
            ctx.Response.ContentType = contentType;
            ctx.Response.Headers["Content-Disposition"] = $"attachment; filename=\"{fileName}\"";
            ctx.Response.ContentLength64 = bytes.Length;
            ctx.Response.OutputStream.Write(bytes, 0, bytes.Length);
            ctx.Response.Close();
        }

        void HandleFiles(HttpListenerContext ctx, System.Collections.Specialized.NameValueCollection query)
        {
            var fileIdsParam = query["fileIds"];
            if (string.IsNullOrWhiteSpace(fileIdsParam))
            {
                WriteJson(ctx, 400, new { error = "fileIds parameter required (comma-separated)" });
                return;
            }

            var fileIds = fileIdsParam.Split(',', StringSplitOptions.RemoveEmptyEntries);
            var results = new List<object>();

            foreach (var fileId in fileIds)
            {
                var id = fileId.Trim();
                var filePath = DriveHelper.FindLocalPathByFileId(id);

                if (filePath == null || !File.Exists(filePath))
                {
                    results.Add(new { fileId = id, success = false, error = "File not found locally" });
                    continue;
                }

                var bytes = File.ReadAllBytes(filePath);
                results.Add(new
                {
                    fileId = id,
                    success = true,
                    fileName = Path.GetFileName(filePath),
                    contentType = GetContentType(filePath),
                    base64Data = Convert.ToBase64String(bytes),
                    filePath
                });
            }

            WriteJson(ctx, 200, new { files = results });
        }

        void HandleCompose(HttpListenerContext ctx)
        {
            try
            {
                string bodyJson;
                using (var reader = new System.IO.StreamReader(ctx.Request.InputStream))
                    bodyJson = reader.ReadToEnd();

                var opts = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                var req = JsonSerializer.Deserialize<ComposeRequestDto>(bodyJson, opts);

                if (req == null)
                {
                    WriteJson(ctx, 400, new { error = "Invalid request body" });
                    return;
                }

                var resolvedPaths = new List<string>();
                var notFound = new List<string>();

                foreach (var fileId in req.FileIds ?? new List<string>())
                {
                    var localPath = DriveHelper.FindLocalPathByFileId(fileId);
                    if (localPath != null && File.Exists(localPath))
                        resolvedPaths.Add(localPath);
                    else
                        notFound.Add(fileId);
                }

                var mapiReq = new MapiHelper.ComposeRequest
                {
                    To        = req.To  ?? new List<string>(),
                    Cc        = req.Cc  ?? new List<string>(),
                    Bcc       = req.Bcc ?? new List<string>(),
                    Subject   = req.Subject ?? "",
                    Body      = req.Body    ?? "",
                    FilePaths = resolvedPaths,
                };

                var result = MapiHelper.Compose(mapiReq);

                WriteJson(ctx, result.Success ? 200 : 500, new
                {
                    success        = result.Success,
                    userCancelled  = result.UserCancelled,
                    attachedCount  = resolvedPaths.Count,
                    skippedFileIds = notFound,
                    error          = result.Error,
                });
            }
            catch (Exception ex)
            {
                WriteJson(ctx, 500, new { error = ex.Message });
            }
        }

        // ── Bluebeam clipboard markup ───────────────────────────────────────

        void HandleClipboardMarkup(HttpListenerContext ctx)
        {
            if (ctx.Request.HttpMethod != "POST")
            {
                WriteJson(ctx, 405, new { error = "POST required" });
                return;
            }

            try
            {
                string bodyJson;
                using (var reader = new StreamReader(ctx.Request.InputStream))
                    bodyJson = reader.ReadToEnd();

                var opts = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                var dto = JsonSerializer.Deserialize<ClipboardMarkupDto>(bodyJson, opts);
                if (dto == null)
                {
                    WriteJson(ctx, 400, new { error = "Invalid request body" });
                    return;
                }

                var result = ClipboardHelper.PutMarkupOnClipboard(new ClipboardHelper.MarkupRequest
                {
                    Subject      = dto.Subject,
                    Location     = dto.Location,
                    ItemCategory = dto.ItemCategory,
                    Status       = dto.Status,
                    Note         = dto.Note,       // NEW — user's free-form note → BSI[0]
                    Contents     = dto.Contents,
                    RcHtml       = dto.RcHtml,
                });

                WriteJson(ctx, result.Success ? 200 : 500, new
                {
                    success          = result.Success,
                    bytesWritten     = result.BytesWritten,
                    dataStringLength = result.DataStringLength,
                    error            = result.Error,
                });
            }
            catch (JsonException jex)
            {
                WriteJson(ctx, 400, new { error = "Malformed JSON: " + jex.Message });
            }
            catch (Exception ex)
            {
                WriteJson(ctx, 500, new { error = ex.Message });
            }
        }

        class ClipboardMarkupDto
        {
            public string? Subject      { get; set; }
            public string? Location     { get; set; }
            public string? ItemCategory { get; set; }
            public string? Status       { get; set; }
            public string? Note         { get; set; }   // NEW — user's free-form note (goes to BSI[0])
            public string? Contents     { get; set; }
            public string? RcHtml       { get; set; }
        }

        class ComposeRequestDto
        {
            public List<string>? To      { get; set; }
            public List<string>? Cc      { get; set; }
            public List<string>? Bcc     { get; set; }
            public string?       Subject { get; set; }
            public string?       Body    { get; set; }
            public List<string>? FileIds { get; set; }
        }

        static void WriteJson(HttpListenerContext ctx, int statusCode, object payload)
        {
            var json = JsonSerializer.Serialize(payload);
            var bytes = Encoding.UTF8.GetBytes(json);
            ctx.Response.StatusCode = statusCode;
            ctx.Response.ContentType = "application/json; charset=utf-8";
            ctx.Response.ContentLength64 = bytes.Length;
            ctx.Response.OutputStream.Write(bytes, 0, bytes.Length);
            ctx.Response.Close();
        }

        static string GetContentType(string filePath) =>
            Path.GetExtension(filePath).ToLowerInvariant() switch
            {
                ".pdf"  => "application/pdf",
                ".png"  => "image/png",
                ".jpg"  => "image/jpeg",
                ".jpeg" => "image/jpeg",
                ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                ".txt"  => "text/plain",
                _       => "application/octet-stream"
            };
    }
}
