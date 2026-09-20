// -----------------------------------------------------------------------
// SWRC Case Context Builder
// Author: John Cleary, Siptu Workers Rights Centre
// License: MIT
// -----------------------------------------------------------------------
using MimeKit;
using MsgReader;
using MsgReader.Outlook;
using PdfPig = UglyToad.PdfPig;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Tesseract;
using DocumentFormat.OpenXml.Packaging;
using ClosedXML.Excel;
using PDFtoImage;
using SkiaSharp;

namespace SWRCCaseContextBuilder
{
    /// <summary>
    /// Provides text/content extraction and OCR routines for the various file types
    /// supported by the case builder: Outlook .eml/.msg emails (including nested
    /// attachments), PDFs (native text or OCR fallback), Word documents, Excel
    /// workbooks, and image files (via OCR).
    /// </summary>
    public static class Extractors
    {
        /// <summary>
        /// Extracts metadata and body text from an .eml email file using MimeKit, and invokes
        /// <paramref name="saveAttachment"/> for each attachment (including nested embedded
        /// messages) so callers can persist and further process them.
        /// </summary>
        /// <param name="path">Full path to the .eml file.</param>
        /// <param name="item">The item to populate with extracted email metadata and body text.</param>
        /// <param name="saveAttachment">Callback invoked with an attachment's raw bytes, display name, and parent item; should return true once saved.</param>
        public static void ExtractEml(string path, Item item, Func<byte[], string, Item, bool> saveAttachment)
        {
            var bytes = File.ReadAllBytes(path);
            var msg = MimeMessage.Load(new MemoryStream(bytes));
            item.Subject = msg.Subject ?? "";
            item.Sender = msg.From.ToString();
            item.Recipients = msg.To.ToString();
            item.Cc = msg.Cc.ToString();
            item.SentDate = msg.Date.ToString();
            item.SentDateValue = msg.Date;
            item.MessageId = msg.MessageId ?? "";

            // body extraction: prefer text/plain
            var text = msg.TextBody ?? msg.HtmlBody ?? "";
            item.Text = text.Length > 0 ? text.Substring(0, Math.Min(text.Length, 300000)) : "";

            var attachments = msg.BodyParts.Where(p => p.IsAttachment || p.ContentType.MediaType == "message").ToList();
            item.AttachmentCount = attachments.Count;
            item.Status = "Email extracted";

            // Build display names for attachments (infer from message subject when possible)
            var displayNames = new List<string>();
            foreach (var part in attachments)
            {
                var name = part.ContentDisposition?.FileName ?? part.ContentType.Name ?? "";
                name = name?.Trim() ?? "";
                if (string.IsNullOrEmpty(name) && part is MessagePart mp)
                {
                    var inner = mp.Message;
                    var subj = inner?.Subject ?? "";
                    name = string.IsNullOrWhiteSpace(subj) ? $"attachment_{displayNames.Count + 1}.eml" : $"{Sanitize(subj)}.eml";
                }
                if (string.IsNullOrEmpty(name)) name = $"attachment_{displayNames.Count + 1}";
                displayNames.Add(name);
            }

            // Duplicate-name logic: if multiple attachments share the same display name and are .eml/.msg, suffix subsequent ones
            var groups = displayNames.GroupBy(n => n.ToLowerInvariant()).ToDictionary(g => g.Key, g => g.Count());
            var seen = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < attachments.Count; i++)
            {
                var part = attachments[i];
                var display = displayNames[i];
                var key = display.ToLowerInvariant();
                var ext = Path.GetExtension(display).ToLowerInvariant();
                var qualifies = ext == ".eml" || ext == ".msg" || part is MessagePart;

                if (groups.ContainsKey(key) && groups[key] > 1 && qualifies)
                {
                    if (!seen.ContainsKey(key)) seen[key] = 1;
                    else seen[key]++;
                    if (seen[key] > 1)
                    {
                        var stem = Path.GetFileNameWithoutExtension(display);
                        display = $"{stem} ({seen[key]}){Path.GetExtension(display)}";
                    }
                }

                // retrieve bytes
                byte[]? data = null;
                if (part is MessagePart mpart)
                {
                    using var ms = new MemoryStream();
                    mpart.Message?.WriteTo(ms);
                    data = ms.ToArray();
                }
                else
                {
                    using var ms = new MemoryStream();
                    if (part is MimePart mp)
                    {
                        mp.Content?.DecodeTo(ms);
                        data = ms.ToArray();
                    }
                }

                // Save attachment; saveAttachment returns true if saved
                saveAttachment(data ?? Array.Empty<byte>(), display, item);
            }
        }

        /// <summary>
        /// Extracts metadata and body text from an Outlook .msg email file using MsgReader,
        /// and invokes <paramref name="saveAttachment"/> for each attachment. Embedded
        /// (non-byte-array) attachments are passed to the optional <paramref name="saveEmbedded"/> callback.
        /// </summary>
        /// <param name="path">Full path to the .msg file.</param>
        /// <param name="item">The item to populate with extracted email metadata and body text.</param>
        /// <param name="saveAttachment">Callback invoked with an attachment's raw bytes, display name, and parent item; should return true once saved.</param>
        /// <param name="saveEmbedded">Optional callback for handling embedded/non-byte attachments (currently unimplemented).</param>
        public static void ExtractMsg(string path, Item item, Func<byte[], string, Item, bool> saveAttachment, Action<object, string, Item>? saveEmbedded = null)
        {
            using var msg = new Storage.Message(path);
            item.Subject = msg.Subject ?? "";
            item.Sender = msg.Sender?.Email ?? msg.Sender?.DisplayName ?? "";
            item.Recipients = msg.GetEmailRecipients(RecipientType.To, false, false) ?? "";
            item.Cc = msg.GetEmailRecipients(RecipientType.Cc, false, false) ?? "";
            item.SentDate = msg.SentOn?.ToString() ?? "";
            item.SentDateValue = msg.SentOn;
            var bodyText = msg.BodyText ?? msg.BodyHtml ?? "";
            item.Text = bodyText.Substring(0, Math.Min(300000, bodyText.Length));
            item.AttachmentCount = msg.Attachments?.Count ?? 0;
            item.Status = "Email extracted";

            var attachmentObjs = msg.Attachments?.OfType<Storage.Attachment>().ToList() ?? new List<Storage.Attachment>();

            var displayNames = attachmentObjs.Select((a, idx) =>
            {
                var nm = a.FileName?.Trim();
                if (string.IsNullOrEmpty(nm)) nm = $"attachment_{idx + 1}";
                return nm;
            }).ToList();

            var groups = displayNames.GroupBy(n => n.ToLowerInvariant()).ToDictionary(g => g.Key, g => g.Count());
            var seen = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < attachmentObjs.Count; i++)
            {
                var att = attachmentObjs[i];
                var display = displayNames[i];
                var key = display.ToLowerInvariant();
                var ext = Path.GetExtension(display).ToLowerInvariant();
                var qualifies = ext == ".eml" || ext == ".msg";

                if (groups.ContainsKey(key) && groups[key] > 1 && qualifies)
                {
                    if (!seen.ContainsKey(key)) seen[key] = 1;
                    else seen[key]++;
                    if (seen[key] > 1)
                    {
                        var stem = Path.GetFileNameWithoutExtension(display);
                        display = $"{stem} ({seen[key]}){Path.GetExtension(display)}";
                    }
                }

                var data = att.Data;

                if (data is byte[] b)
                {
                    saveAttachment(b, display, item);
                }
                else
                {
                    // fallback: try to extract to temp and pass path to saveEmbedded if provided
                    if (saveEmbedded != null) saveEmbedded(att, display, item);
                }
            }
        }

        /// <summary>
        /// Extracts the text layer from a PDF file using PdfPig, page by page.
        /// </summary>
        /// <param name="path">Full path to the PDF file.</param>
        /// <param name="needsOcr">Set to <c>true</c> if no extractable text was found (or the PDF could not be opened), indicating OCR should be attempted instead.</param>
        /// <returns>The extracted text, with a "[PDF page N]" header before each page's content.</returns>
        public static string ExtractPdfText(string path, out bool needsOcr)
        {
            needsOcr = false;
            try
            {
                using var doc = UglyToad.PdfPig.PdfDocument.Open(path);
                var sb = new StringBuilder();
                for (int i = 0; i < doc.NumberOfPages; i++)
                {
                    var page = doc.GetPage(i + 1);
                    sb.AppendLine($"[PDF page {i + 1}]");
                    sb.AppendLine(page.Text ?? "");
                }
                var text = sb.ToString();
                var strippedForCheck = text.Replace("[PDF page", "").Replace("]", "");
                needsOcr = string.IsNullOrWhiteSpace(strippedForCheck);
                return text;
            }
            catch
            {
                needsOcr = true;
                return "";
            }
        }

        /// <summary>
        /// Rasterizes each page of a (likely scanned/image-only) PDF and runs OCR on each page image,
        /// used as a fallback when PdfPig finds no extractable text layer.
        /// </summary>
        /// <param name="path">Full path to the PDF file.</param>
        /// <param name="error">Receives the first OCR error message encountered, if any.</param>
        /// <returns>The concatenated OCR'd text for all pages, with a "[PDF page N - OCR]" header before each page.</returns>
        public static string OcrPdf(string path, out string error)
        {
            error = "";
            var sb = new StringBuilder();
            try
            {
                var pageCount = PDFtoImage.Conversion.GetPageCount(path);
                for (int i = 0; i < pageCount; i++)
                {
                    using var bitmap = PDFtoImage.Conversion.ToImage(path, page: i, options: new PDFtoImage.RenderOptions(Dpi: 300));
                    using var data = bitmap.Encode(SKEncodedImageFormat.Png, 100);
                    var tempPath = Path.Combine(Path.GetTempPath(), $"pdfocr_{Guid.NewGuid():N}.png");
                    try
                    {
                        File.WriteAllBytes(tempPath, data.ToArray());
                        var pageText = OcrImage(tempPath, out var pageError);
                        if (!string.IsNullOrEmpty(pageError) && string.IsNullOrEmpty(error)) error = pageError;
                        sb.AppendLine($"[PDF page {i + 1} - OCR]");
                        sb.AppendLine(pageText ?? "");
                    }
                    finally
                    {
                        try { File.Delete(tempPath); } catch { /* best effort */ }
                    }
                }
            }
            catch (Exception ex)
            {
                error = ex.Message;
            }
            return sb.ToString();
        }

        /// <summary>
        /// Extracts the plain text body content of a Word document (.docx/.docm) via OpenXML.
        /// </summary>
        /// <param name="path">Full path to the Word document.</param>
        /// <returns>The document body's inner text, or an empty string if extraction fails.</returns>
        public static string ExtractDocxText(string path)
        {
            try
            {
                using var doc = WordprocessingDocument.Open(path, false);
                var body = doc.MainDocumentPart?.Document?.Body;
                return body?.InnerText ?? "";
            }
            catch { return ""; }
        }

        /// <summary>
        /// Extracts a text representation of every worksheet in an Excel workbook (.xlsx/.xlsm),
        /// with cell values in each used row joined by " | " separators.
        /// </summary>
        /// <param name="path">Full path to the Excel workbook.</param>
        /// <returns>The extracted text, with a "[Worksheet: name]" header before each sheet's rows, or an empty string if extraction fails.</returns>
        public static string ExtractXlsxText(string path)
        {
            try
            {
                var wb = new ClosedXML.Excel.XLWorkbook(path);
                var sb = new StringBuilder();
                foreach (var ws in wb.Worksheets)
                {
                    sb.AppendLine($"[Worksheet: {ws.Name}]");
                    foreach (var row in ws.RowsUsed())
                    {
                        var vals = row.Cells().Select(c => c.GetValue<string>());
                        sb.AppendLine(string.Join(" | ", vals));
                    }
                }
                return sb.ToString();
            }
            catch { return ""; }
        }

        /// <summary>
        /// Runs OCR on an image file using the Tesseract .NET engine. Requires eng.traineddata to be
        /// present in <see cref="Utils.TessDataFolder"/> (auto-provisioned via <see cref="Utils.EnsureTessDataAvailable"/>).
        /// Returns an error message (rather than throwing/silently returning empty) so callers can
        /// correctly flag "OCR required" only when OCR itself actually failed.
        /// </summary>
        /// <param name="imagePath">Full path to the image file to OCR.</param>
        /// <param name="error">Receives a description of any failure (missing language data, exception, or no text recognized).</param>
        /// <returns>The recognized text, or an empty string if OCR could not be performed or found no text.</returns>
        public static string OcrImage(string imagePath, out string error)
        {
            error = "";
            try
            {
                if (!Utils.EnsureTessDataAvailable(out var tessDetails))
                {
                    error = tessDetails;
                    return "";
                }

                using var ocr = new TesseractEngine(Utils.TessDataFolder, "eng", EngineMode.Default);
                using var img = Pix.LoadFromFile(imagePath);
                using var page = ocr.Process(img);
                var text = page.GetText() ?? "";
                if (string.IsNullOrWhiteSpace(text))
                {
                    error = "OCR completed but no text was recognized.";
                }
                return text;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return "";
            }
        }

        /// <summary>
        /// Removes characters that are invalid in file names from a string, used to build
        /// safe display names for email attachments derived from subject lines.
        /// </summary>
        /// <param name="s">The string to sanitize.</param>
        /// <returns>The sanitized string with invalid file name characters removed.</returns>
        private static string Sanitize(string s) =>
            string.Concat(s.Split(Path.GetInvalidFileNameChars())).Trim();
    }
}