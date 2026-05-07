// CropperEndpoint.cs
//
// POST /process-plan endpoint. Accepts the marked-up plan PDF inline as
// base64, runs the bundled cropper against it, and uploads manifest +
// per-markup PDFs to Supabase Storage at:
//   tdlr-reports/_cropped/{user_id}/{plan_pdf_hash}/
//
// Why base64 instead of multipart or local path:
//   - FlutterFlow's web file picker hands the app FFUploadedFile bytes in
//     browser memory; there is no local disk path to give the helper.
//   - Multipart parsing in raw HttpListener requires hand-rolling boundaries
//     or pulling in a NuGet package; base64-in-JSON is two lines on each end.
//   - Bandwidth doesn't matter on localhost; a 100 MB plan + base64 overhead
//     uploads in ~1 sec on loopback.
//
// Wire into HttpServer.cs:
//   case "/process-plan":
//       CropperEndpoint.Handle(ctx).GetAwaiter().GetResult();
//       break;
//
// Request (POST, JSON):
//   {
//     "plan_pdf_base64": "JVBERi0xLjQK...",
//     "supabase_url":    "https://vidyyayewzthhisewvea.supabase.co",
//     "access_token":    "<user JWT>",
//     "user_id":         "<uuid>"
//   }
//
// Preflight (OPTIONS): handled with Access-Control-Allow-Private-Network
// because Chrome treats the FF web app -> localhost hop as a public->private
// request.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace TabsPortalHelper
{
    public static class CropperEndpoint
    {
        private const string Bucket = "tdlr-reports";
        private const string CroppedFolderRoot = "_cropped";

        private static readonly HttpClient _http = new HttpClient
        {
            Timeout = TimeSpan.FromMinutes(5),
        };

        public static async Task Handle(HttpListenerContext ctx)
        {
            try
            {
                // ---- CORS preflight ------------------------------------------------
                // FF web app at tabsportal.com -> localhost:52874 is cross-origin AND
                // public->private, so Chrome enforces both standard CORS and Private
                // Network Access checks. Both need their preflight handled here.
                if (ctx.Request.HttpMethod == "OPTIONS")
                {
                    SetCorsHeaders(ctx);
                    ctx.Response.StatusCode = 204;
                    ctx.Response.Close();
                    return;
                }

                if (ctx.Request.HttpMethod != "POST")
                {
                    await WriteJson(ctx, 405, new { error = "method not allowed" });
                    return;
                }

                ProcessPlanRequest req;
                try
                {
                    using var reader = new StreamReader(ctx.Request.InputStream, Encoding.UTF8);
                    var body = await reader.ReadToEndAsync();
                    req = JsonSerializer.Deserialize<ProcessPlanRequest>(body)
                        ?? throw new Exception("empty body");
                }
                catch (Exception ex)
                {
                    await WriteJson(ctx, 400, new { error = $"bad request: {ex.Message}" });
                    return;
                }

                if (string.IsNullOrWhiteSpace(req.plan_pdf_base64) ||
                    string.IsNullOrWhiteSpace(req.supabase_url) ||
                    string.IsNullOrWhiteSpace(req.access_token) ||
                    string.IsNullOrWhiteSpace(req.user_id))
                {
                    await WriteJson(ctx, 400, new { error = "missing required fields: plan_pdf_base64, supabase_url, access_token, user_id" });
                    return;
                }

                // Make sure the bundled node + cropper are on disk.
                if (!CropperBundle.ExtractIfNeeded())
                {
                    await WriteJson(ctx, 500, new
                    {
                        error = "cropper bundle extraction failed",
                        detail = CropperBundle.LastError?.Message,
                    });
                    return;
                }

                var sw = Stopwatch.StartNew();

                // Decode the base64 plan PDF into a temp file. We pay one extra disk
                // write here, but it lets cropper.js operate as a normal CLI on a
                // file path - no need to teach it stdin streaming.
                byte[] planBytes;
                try
                {
                    planBytes = Convert.FromBase64String(req.plan_pdf_base64);
                }
                catch (FormatException)
                {
                    await WriteJson(ctx, 400, new { error = "plan_pdf_base64 is not valid base64" });
                    return;
                }

                if (planBytes.Length < 100)
                {
                    await WriteJson(ctx, 400, new { error = $"plan_pdf_base64 decoded to {planBytes.Length} bytes; that's not a PDF" });
                    return;
                }

                var workDir = Path.Combine(
                    Path.GetTempPath(),
                    $"tabs-cropper-{Guid.NewGuid():N}");
                Directory.CreateDirectory(workDir);

                var tempPlanPath = Path.Combine(workDir, "plan.pdf");
                var tempOutDir = Path.Combine(workDir, "out");
                Directory.CreateDirectory(tempOutDir);

                try
                {
                    await File.WriteAllBytesAsync(tempPlanPath, planBytes);

                    // Hash matches cropper.js (sha256 hex, first 16 chars) and the
                    // edge function's planPdfHash() - so end-to-end keys line up.
                    var hash = ComputePlanHash(planBytes);

                    var cropResult = await RunCropper(tempPlanPath, tempOutDir);
                    if (cropResult.ExitCode != 0)
                    {
                        await WriteJson(ctx, 500, new
                        {
                            error = "cropper exited non-zero",
                            exit_code = cropResult.ExitCode,
                            stderr = cropResult.Stderr,
                        });
                        return;
                    }

                    var manifestPath = Path.Combine(tempOutDir, "manifest.json");
                    if (!File.Exists(manifestPath))
                    {
                        await WriteJson(ctx, 500, new
                        {
                            error = "cropper produced no manifest",
                            stderr = cropResult.Stderr,
                        });
                        return;
                    }

                    var manifestText = await File.ReadAllTextAsync(manifestPath);
                    var manifest = JsonSerializer.Deserialize<CropManifest>(manifestText)
                        ?? throw new Exception("manifest parse failed");

                    if (!string.Equals(manifest.plan_pdf_hash, hash, StringComparison.Ordinal))
                    {
                        await WriteJson(ctx, 500, new
                        {
                            error = "manifest hash mismatch (helper vs cropper)",
                            expected = hash,
                            got = manifest.plan_pdf_hash,
                        });
                        return;
                    }

                    var folderPath = $"{CroppedFolderRoot}/{req.user_id}/{hash}";
                    var uploadTasks = new List<Task>();

                    uploadTasks.Add(UploadFile(
                        req.supabase_url!, req.access_token!,
                        $"{folderPath}/manifest.json",
                        await File.ReadAllBytesAsync(manifestPath),
                        "application/json"
                    ));

                    foreach (var m in manifest.markups)
                    {
                        var localPath = Path.Combine(tempOutDir, m.crop_pdf_path);
                        if (!File.Exists(localPath)) continue;
                        var bytes = await File.ReadAllBytesAsync(localPath);
                        uploadTasks.Add(UploadFile(
                            req.supabase_url!, req.access_token!,
                            $"{folderPath}/{m.crop_pdf_path}",
                            bytes,
                            "application/pdf"
                        ));
                    }

                    await Task.WhenAll(uploadTasks);
                    int uploaded = uploadTasks.Count;

                    sw.Stop();
                    await WriteJson(ctx, 200, new
                    {
                        success = true,
                        plan_pdf_hash = hash,
                        manifest_path = $"{folderPath}/manifest.json",
                        markup_count = manifest.markups.Count,
                        uploaded,
                        elapsed_ms = sw.ElapsedMilliseconds,
                    });
                }
                finally
                {
                    try { Directory.Delete(workDir, recursive: true); }
                    catch { /* best-effort */ }
                }
            }
            catch (Exception ex)
            {
                try
                {
                    await WriteJson(ctx, 500, new { error = ex.Message, stack = ex.StackTrace });
                }
                catch { /* response may already be sent */ }
            }
        }

        // ---- helpers ---------------------------------------------------------

        private static async Task<CropperResult> RunCropper(string planPath, string outDir)
        {
            var nodeExe = CropperBundle.NodeExePath;
            var cropperJs = CropperBundle.CropperJsPath;

            if (!File.Exists(nodeExe))
            {
                return new CropperResult
                {
                    ExitCode = -1,
                    Stderr = $"node.exe not found at {nodeExe}.",
                };
            }
            if (!File.Exists(cropperJs))
            {
                return new CropperResult
                {
                    ExitCode = -1,
                    Stderr = $"cropper.js not found at {cropperJs}.",
                };
            }

            var psi = new ProcessStartInfo
            {
                FileName = nodeExe,
                ArgumentList = { cropperJs, planPath, outDir },
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                WorkingDirectory = Path.GetDirectoryName(cropperJs) ?? CropperBundle.InstallRoot,
            };

            using var proc = new Process { StartInfo = psi };
            var stdoutBuf = new StringBuilder();
            var stderrBuf = new StringBuilder();
            proc.OutputDataReceived += (_, e) => { if (e.Data != null) stdoutBuf.AppendLine(e.Data); };
            proc.ErrorDataReceived += (_, e) => { if (e.Data != null) stderrBuf.AppendLine(e.Data); };

            proc.Start();
            proc.BeginOutputReadLine();
            proc.BeginErrorReadLine();
            await proc.WaitForExitAsync();

            return new CropperResult
            {
                ExitCode = proc.ExitCode,
                Stdout = stdoutBuf.ToString(),
                Stderr = stderrBuf.ToString(),
            };
        }

        /// <summary>
        /// Upload a single file to Supabase Storage via REST. The access_token is
        /// the user's JWT (from the FF Supabase client), so RLS rules apply normally.
        /// </summary>
        private static async Task UploadFile(
            string supabaseUrl, string accessToken,
            string storagePath, byte[] bytes, string contentType)
        {
            var uri = $"{supabaseUrl.TrimEnd('/')}/storage/v1/object/{Bucket}/{storagePath}";
            using var req = new HttpRequestMessage(HttpMethod.Post, uri);
            req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            req.Headers.Add("x-upsert", "true");
            req.Content = new ByteArrayContent(bytes);
            req.Content.Headers.ContentType = new MediaTypeHeaderValue(contentType);

            using var resp = await _http.SendAsync(req);
            if (!resp.IsSuccessStatusCode)
            {
                var body = await resp.Content.ReadAsStringAsync();
                throw new Exception(
                    $"upload failed for {storagePath}: {resp.StatusCode} {body}");
            }
        }

        private static string ComputePlanHash(byte[] bytes)
        {
            using var sha = SHA256.Create();
            var hash = sha.ComputeHash(bytes);
            // First 16 hex chars (8 bytes). Matches cropper.js and edge function.
            var sb = new StringBuilder(16);
            for (int i = 0; i < 8; i++) sb.Append(hash[i].ToString("x2"));
            return sb.ToString();
        }

        private static void SetCorsHeaders(HttpListenerContext ctx)
        {
            ctx.Response.AddHeader("Access-Control-Allow-Origin", "*");
            ctx.Response.AddHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
            ctx.Response.AddHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
            // Chrome's Private Network Access: the FF PWA on a public origin
            // making requests to localhost is treated as public->private and
            // requires this header on the preflight response.
            ctx.Response.AddHeader("Access-Control-Allow-Private-Network", "true");
            ctx.Response.AddHeader("Access-Control-Max-Age", "86400");
        }

        private static async Task WriteJson(HttpListenerContext ctx, int status, object body)
        {
            ctx.Response.StatusCode = status;
            ctx.Response.ContentType = "application/json";
            SetCorsHeaders(ctx);
            var bytes = JsonSerializer.SerializeToUtf8Bytes(body, new JsonSerializerOptions
            {
                PropertyNamingPolicy = null,
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            });
            ctx.Response.ContentLength64 = bytes.Length;
            await ctx.Response.OutputStream.WriteAsync(bytes);
            ctx.Response.Close();
        }

        // ---- DTOs ------------------------------------------------------------

        private class ProcessPlanRequest
        {
            public string? plan_pdf_base64 { get; set; }
            public string? supabase_url { get; set; }
            public string? access_token { get; set; }
            public string? user_id { get; set; }
        }

        private class CropManifest
        {
            public int version { get; set; }
            public string plan_pdf_hash { get; set; } = "";
            public string generated_at { get; set; } = "";
            public List<ManifestEntry> markups { get; set; } = new();
        }

        private class ManifestEntry
        {
            public string nm { get; set; } = "";
            public int page_index { get; set; }
            public string crop_pdf_path { get; set; } = "";
            public List<double> crop_rect { get; set; } = new();
        }

        private class CropperResult
        {
            public int ExitCode { get; set; }
            public string Stdout { get; set; } = "";
            public string Stderr { get; set; } = "";
        }
    }
}
