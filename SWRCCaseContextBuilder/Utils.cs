// -----------------------------------------------------------------------
// SWRC Case Context Builder
// Author: John Cleary, Siptu Workers Rights Centre
// License: MIT
// -----------------------------------------------------------------------
using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace SWRCCaseContextBuilder
{
    /// <summary>
    /// General-purpose helper utilities used across the application, including
    /// file hashing, checking for a system Tesseract installation, and ensuring
    /// the bundled Tesseract .NET OCR engine has the language data it needs.
    /// </summary>
    public static class Utils
    {
        /// <summary>
        /// Computes the SHA-256 hash of the file at the given path and returns it
        /// as a lowercase hexadecimal string. Used to detect byte-for-byte duplicate
        /// files and to build stable, collision-resistant generated file names.
        /// </summary>
        /// <param name="path">Full path to the file to hash.</param>
        /// <returns>The SHA-256 hash of the file contents, as a hex string.</returns>
        public static string Sha256Hex(string path)
        {
            using var sha = SHA256.Create();
            using var fs = File.OpenRead(path);
            var hash = sha.ComputeHash(fs);
            var sb = new StringBuilder();
            foreach (var b in hash) sb.Append(b.ToString("x2"));
            return sb.ToString();
        }

        /// <summary>
        /// Checks whether a system-installed <c>tesseract</c> executable is available on the
        /// current PATH, by attempting to invoke <c>tesseract --version</c>. This is informational
        /// only; the application's OCR features use the bundled Tesseract .NET engine and do not
        /// require a system installation.
        /// </summary>
        /// <param name="version">Receives the reported Tesseract version string, if found.</param>
        /// <param name="details">Receives a human-readable description of the outcome.</param>
        /// <returns><c>true</c> if a system Tesseract installation was detected; otherwise <c>false</c>.</returns>
        public static bool CheckTesseract(out string version, out string details)
        {
            version = "";
            details = "";
            try
            {
                var p = new Process();
                p.StartInfo.FileName = "tesseract";
                p.StartInfo.Arguments = "--version";
                p.StartInfo.RedirectStandardOutput = true;
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.CreateNoWindow = true;
                p.Start();
                var outp = p.StandardOutput.ReadLine();
                p.WaitForExit(2000);
                if (!string.IsNullOrEmpty(outp))
                {
                    version = outp.Trim();
                    details = "Tesseract available.";
                    return true;
                }
                details = "tesseract present but version not returned.";
                return false;
            }
            catch (Exception ex)
            {
                details = ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Folder where OCR language data files are stored (downloaded on demand).
        /// Separate from any system tesseract install so the app works standalone
        /// using the Tesseract .NET engine (no external tesseract.exe required).
        /// </summary>
        public static string TessDataFolder =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SWRCCaseContextBuilder", "tessdata");

        /// <summary>
        /// Download URL for the English Tesseract "fast" trained data file, used to
        /// provision OCR language support on first run.
        /// </summary>
        private const string EngTrainedDataUrl = "https://github.com/tesseract-ocr/tessdata_fast/raw/main/eng.traineddata";

        /// <summary>
        /// Ensures the English OCR language file is available locally, downloading it once if needed.
        /// Returns (success, details) so callers can surface clear diagnostics instead of silently
        /// falling back to "OCR required".
        /// </summary>
        /// <param name="details">Receives a human-readable description of the outcome (success or failure reason).</param>
        /// <returns><c>true</c> if the language data file is present and ready to use; otherwise <c>false</c>.</returns>
        public static bool EnsureTessDataAvailable(out string details)
        {
            details = "";
            try
            {
                Directory.CreateDirectory(TessDataFolder);
                var engPath = Path.Combine(TessDataFolder, "eng.traineddata");
                if (File.Exists(engPath) && new FileInfo(engPath).Length > 0)
                {
                    details = $"English OCR language data found at {engPath}.";
                    return true;
                }

                details = "Downloading English OCR language data (eng.traineddata)...";
                DownloadFileAsync(EngTrainedDataUrl, engPath).GetAwaiter().GetResult();

                if (File.Exists(engPath) && new FileInfo(engPath).Length > 0)
                {
                    details = $"Downloaded English OCR language data to {engPath}.";
                    return true;
                }

                details = "Failed to download English OCR language data.";
                return false;
            }
            catch (Exception ex)
            {
                details = $"Could not prepare OCR language data: {ex.Message}";
                return false;
            }
        }

        /// <summary>
        /// Downloads the file at <paramref name="url"/> to <paramref name="destinationPath"/>,
        /// writing to a temporary file first and then atomically moving it into place so that
        /// a partially-downloaded file is never mistaken for a complete one.
        /// </summary>
        /// <param name="url">The source URL to download from.</param>
        /// <param name="destinationPath">The final destination path for the downloaded file.</param>
        private static async Task DownloadFileAsync(string url, string destinationPath)
        {
            var tempPath = destinationPath + ".download";
            using (var http = new HttpClient())
            {
                http.Timeout = TimeSpan.FromMinutes(2);
                using var response = await http.GetAsync(url, HttpCompletionOption.ResponseHeadersRead);
                response.EnsureSuccessStatusCode();
                using var fs = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None);
                await response.Content.CopyToAsync(fs);
            }
            File.Move(tempPath, destinationPath, true);
        }
    }
}