// -----------------------------------------------------------------------
// SWRC Case Context Builder
// Author: John Cleary, Siptu Workers Rights Centre
// License: MIT
// -----------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace SWRCCaseContextBuilder
{
    /// <summary>
    /// Orchestrates the end-to-end case processing pipeline: copies the source folder as a
    /// stable reference, recursively extracts and processes every file (including nested
    /// email attachments to any depth), runs text extraction/OCR as appropriate, detects
    /// duplicate content, and generates the manifest and Word context reports via <see cref="Outputs"/>.
    /// </summary>
    public class Builder
    {
        private readonly string _source;
        private readonly string _output;
        private readonly string _caseRef;
        private readonly Action<ProgressReport> _progress;
        private readonly OutputTypes _outputTypes;
        private readonly bool _removeDuplicates;
        private int _counter = 0;
        private string _extractedRoot = "";
        private readonly List<(string, string)> _errors = new();

        /// <summary>
        /// Creates a new <see cref="Builder"/> for a single processing run.
        /// </summary>
        /// <param name="source">Path to the source folder containing the case files to process.</param>
        /// <param name="output">Path to the (must be empty or non-existent) output folder to populate.</param>
        /// <param name="caseRef">Case reference used to name generated report files.</param>
        /// <param name="progress">Callback invoked with progress updates as processing proceeds.</param>
        /// <param name="outputTypes">Which Word context report(s) to generate.</param>
        /// <param name="removeDuplicates">When true, removes duplicate (by content hash) extracted files, keeping only the first copy.</param>
        public Builder(string source, string output, string caseRef, Action<ProgressReport> progress, OutputTypes outputTypes = OutputTypes.AllText, bool removeDuplicates = false)
        {
            _source = source;
            _output = output;
            _caseRef = caseRef;
            _progress = progress;
            _outputTypes = outputTypes;
            _removeDuplicates = removeDuplicates;
        }

        /// <summary>
        /// Generates the next sequential item identifier, prefixed with "E" for emails or "D" for other documents.
        /// </summary>
        /// <param name="email">Whether the item is an email (.eml/.msg).</param>
        /// <returns>A unique, zero-padded identifier such as "E00001" or "D00002".</returns>
        private string NextId(bool email) => $"{(email ? "E" : "D")}{Interlocked.Increment(ref _counter):00000}";

        /// <summary>
        /// Runs the full processing pipeline: validates folders, copies the source as a stable
        /// reference, extracts and processes every file (recursively for emails), detects and
        /// optionally removes duplicate files, and writes the manifest and report outputs.
        /// </summary>
        /// <returns>A <see cref="RunResult"/> summarising the number of items processed, exceptions encountered, and the output path.</returns>
        /// <exception cref="InvalidOperationException">Thrown when the source folder does not exist or the output folder is not empty.</exception>
        public RunResult Run()
        {
            var items = new List<Item>();

            _progress(new ProgressReport { Percent = 0, Message = "Preparing folder structure..." });

            // Validate
            if (!Directory.Exists(_source)) throw new InvalidOperationException("Source folder does not exist.");
            if (!Directory.Exists(_output)) Directory.CreateDirectory(_output);
            if (Directory.EnumerateFileSystemEntries(_output).Any()) throw new InvalidOperationException("Output folder must be empty.");

            var orig = Path.Combine(_output, "00_Original_Source");
            var extracted = Path.Combine(_output, "01_Extracted_Files");
            var context = Path.Combine(_output, "02_Context");
            var exceptions = Path.Combine(_output, "03_Review_Exceptions");
            Directory.CreateDirectory(orig);
            Directory.CreateDirectory(extracted);
            Directory.CreateDirectory(context);
            Directory.CreateDirectory(exceptions);

            _extractedRoot = extracted;

            _progress(new ProgressReport { Percent = 5, Message = "Copying source (stable reference)..." });

            // Copy tree
            CopyDirectory(_source, Path.Combine(orig, Path.GetFileName(_source)));

            // Walk and process
            var allFiles = Directory.EnumerateFiles(Path.Combine(orig, Path.GetFileName(_source)), "*", SearchOption.AllDirectories).OrderBy(p => p, StringComparer.OrdinalIgnoreCase).ToList();
            var total = allFiles.Count;
            var processed = 0;

            foreach (var file in allFiles)
            {
                var rel = Path.GetRelativePath(orig, file);
                var ext = Path.GetExtension(file).ToLowerInvariant();
                var isEmail = ext == ".eml" || ext == ".msg";
                var id = NextId(isEmail);
                var fileInfo = new FileInfo(file);

                var item = new Item
                {
                    Id = id,
                    ParentId = "",
                    Depth = 0,
                    SourceRelativePath = rel,
                    OriginalName = Path.GetFileName(file),
                    Extension = ext,
                    SizeBytes = fileInfo.Length,
                    Sha256 = Utils.Sha256Hex(file),
                    Kind = isEmail ? "email" : "file",
                    DocumentDate = fileInfo.LastWriteTimeUtc
                };

                // Generated name: includes id, sanitized name, hash prefix
                var stem = Path.GetFileNameWithoutExtension(item.OriginalName);
                var safeStem = _sanitize(stem);
                item.GeneratedName = $"{id}_{safeStem}_{item.Sha256.Substring(0,8)}{ext}";
                var dest = Path.Combine(extracted, item.GeneratedName);
                File.Copy(file, dest, true);
                item.ExtractedRelativePath = Path.GetRelativePath(_output, dest);

                items.Add(item);

                ProcessItem(item, dest, items);

                processed++;
                _progress(new ProgressReport { Percent = 5 + (int)((processed / (double)total) * 90), Message = $"Processing {processed}/{total}: {item.OriginalName}" });
            }

            // classify duplicates by hash
            var groupsByHash = items.GroupBy(i => i.Sha256).Where(g => g.Count() > 1).ToList();
            int groupNumber = 0;
            foreach (var g in groupsByHash)
            {
                groupNumber++;
                var groupId = $"DUP-{groupNumber:0000}";
                var ordered = g.OrderBy(i => i.Id, StringComparer.OrdinalIgnoreCase).ToList();
                var keeper = ordered.First();
                foreach (var it in ordered) it.DuplicateGroup = groupId;

                if (_removeDuplicates)
                {
                    foreach (var dup in ordered.Skip(1))
                    {
                        RemoveDuplicateFile(dup, keeper);
                    }
                }
            }

            Outputs.WriteOutputs(items, _output, _caseRef, "1.0.0", _outputTypes, _removeDuplicates);

            return new RunResult { Items = items.Count, Exceptions = _errors.Count, Output = _output };
        }

        /// <summary>
        /// Shared extraction logic for both top-level files and nested attachments.
        /// Recursively processes emails, extracting nested attachments to arbitrary depth.
        /// Populates <paramref name="item"/>'s <see cref="Item.Text"/>, <see cref="Item.Status"/>
        /// and related fields based on the file's extension (email, PDF, Word, Excel, image, or unsupported).
        /// </summary>
        /// <param name="item">The item metadata to populate.</param>
        /// <param name="dest">The full path to the extracted copy of the file on disk.</param>
        /// <param name="items">The overall list of items, to which any newly discovered nested attachments are added.</param>
        private void ProcessItem(Item item, string dest, List<Item> items)
        {
            var ext = item.Extension;
            try
            {
                if (ext == ".eml")
                {
                    Extractors.ExtractEml(dest, item, (bytes, name, parentItem) =>
                    {
                        var attachmentPath = SaveAttachmentBytes(_extractedRoot, parentItem.Id, name, bytes);
                        var childItem = AddExtractedFile(attachmentPath, items, parentItem.Id, parentItem.Depth + 1);
                        ProcessItem(childItem, attachmentPath, items);
                        return true;
                    });
                }
                else if (ext == ".msg")
                {
                    Extractors.ExtractMsg(dest, item, (bytes, name, parentItem) =>
                    {
                        var attachmentPath = SaveAttachmentBytes(_extractedRoot, parentItem.Id, name, bytes);
                        var childItem = AddExtractedFile(attachmentPath, items, parentItem.Id, parentItem.Depth + 1);
                        ProcessItem(childItem, attachmentPath, items);
                        return true;
                    }, (att, name, parentItem) =>
                    {
                        // handle embedded attachments if needed (not implemented)
                    });
                }
                else if (ext == ".pdf")
                {
                    var text = Extractors.ExtractPdfText(dest, out var needsOcr);
                    if (needsOcr)
                    {
                        var ocrText = Extractors.OcrPdf(dest, out var ocrError);
                        if (!string.IsNullOrWhiteSpace(ocrText))
                        {
                            item.Text = ocrText;
                            item.Status = "Text extracted (OCR)";
                        }
                        else
                        {
                            item.Text = text;
                            item.Status = "OCR required";
                            item.Notes = string.IsNullOrEmpty(item.Notes)
                                ? (string.IsNullOrEmpty(ocrError) ? "No extractable text; OCR produced no text." : $"OCR failed: {ocrError}")
                                : item.Notes;
                        }
                    }
                    else
                    {
                        item.Text = text;
                        item.Status = "Text extracted";
                    }
                }
                else if (ext == ".docx" || ext == ".docm")
                {
                    item.Text = Extractors.ExtractDocxText(dest);
                    item.Status = "Text extracted";
                }
                else if (ext == ".xlsx" || ext == ".xlsm")
                {
                    item.Text = Extractors.ExtractXlsxText(dest);
                    item.Status = "Text extracted";
                }
                else if (new[] { ".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp" }.Contains(ext))
                {
                    var ocrText = Extractors.OcrImage(dest, out var ocrError);
                    if (!string.IsNullOrWhiteSpace(ocrText))
                    {
                        item.Text = ocrText;
                        item.Status = "Text extracted (OCR)";
                    }
                    else
                    {
                        item.Status = "OCR required";
                        item.Notes = string.IsNullOrEmpty(item.Notes)
                            ? (string.IsNullOrEmpty(ocrError) ? "OCR produced no text." : $"OCR failed: {ocrError}")
                            : item.Notes;
                    }
                }
                else
                {
                    item.Status = "Copied only";
                }
            }
            catch (Exception ex)
            {
                item.Status = "Exception";
                item.Notes = ex.Message;
                _errors.Add((item.Id, ex.Message));
            }

            item.TextChars = item.Text?.Length ?? 0;
        }

        /// <summary>
        /// Recursively copies all files and subfolders from <paramref name="src"/> to <paramref name="dst"/>,
        /// used to create the stable "00_Original_Source" reference copy of the case folder.
        /// </summary>
        /// <param name="src">The source directory to copy from.</param>
        /// <param name="dst">The destination directory to copy into (created if it does not exist).</param>
        private void CopyDirectory(string src, string dst)
        {
            Directory.CreateDirectory(dst);
            foreach (var dir in Directory.GetDirectories(src, "*", SearchOption.AllDirectories))
            {
                Directory.CreateDirectory(dir.Replace(src, dst));
            }
            foreach (var file in Directory.GetFiles(src, "*", SearchOption.AllDirectories))
            {
                var dest = file.Replace(src, dst);
                var destDir = Path.GetDirectoryName(dest);
                if (!string.IsNullOrEmpty(destDir)) Directory.CreateDirectory(destDir);
                File.Copy(file, dest, true);
            }
        }

        /// <summary>
        /// Saves an email attachment's raw bytes to disk in a per-parent attachment folder,
        /// sanitizing the file name and avoiding collisions by appending a numeric suffix if needed.
        /// </summary>
        /// <param name="extractedRoot">The root "01_Extracted_Files" folder.</param>
        /// <param name="parentId">The <see cref="Item.Id"/> of the parent email this attachment belongs to.</param>
        /// <param name="filename">The attachment's original/display file name.</param>
        /// <param name="bytes">The raw attachment content.</param>
        /// <returns>The full path to the saved attachment file.</returns>
        private string SaveAttachmentBytes(string extractedRoot, string parentId, string filename, byte[] bytes)
        {
            var folder = Path.Combine(extractedRoot, $"._attachments_{parentId}");
            Directory.CreateDirectory(folder);
            var safe = _sanitize(Path.GetFileNameWithoutExtension(filename));
            var ext = Path.GetExtension(filename);
            var path = Path.Combine(folder, safe + ext);
            // When filename duplicates exist, add numeric suffix to file on disk to avoid collision
            var idx = 1;
            var basePath = path;
            while (File.Exists(path))
            {
                path = Path.Combine(folder, $"{safe} ({++idx}){ext}");
            }
            File.WriteAllBytes(path, bytes);
            return path;
        }

        /// <summary>
        /// Creates and registers a new <see cref="Item"/> for a file already extracted to disk
        /// (e.g. a nested email attachment), computing its metadata and hash.
        /// </summary>
        /// <param name="fullPath">The full path to the extracted file.</param>
        /// <param name="items">The overall items list to add the new item to.</param>
        /// <param name="parentId">The <see cref="Item.Id"/> of the parent item this was extracted from.</param>
        /// <param name="depth">The nesting depth of the new item relative to the top-level source item.</param>
        /// <returns>The newly created and registered <see cref="Item"/>.</returns>
        private Item AddExtractedFile(string fullPath, List<Item> items, string parentId, int depth)
        {
            var fileInfo = new FileInfo(fullPath);
            var ext = fileInfo.Extension.ToLowerInvariant();
            var isEmail = ext == ".eml" || ext == ".msg";
            var id = NextId(isEmail);
            var item = new Item
            {
                Id = id,
                ParentId = parentId,
                Depth = depth,
                SourceRelativePath = $"{Path.GetFileName(_source)} > {Path.GetFileName(fullPath)}",
                OriginalName = Path.GetFileName(fullPath),
                GeneratedName = Path.GetFileName(fullPath),
                Extension = ext,
                SizeBytes = fileInfo.Length,
                Sha256 = Utils.Sha256Hex(fullPath),
                Kind = isEmail ? "email" : "file",
                ExtractedRelativePath = Path.GetRelativePath(_output, fullPath),
                DocumentDate = fileInfo.LastWriteTimeUtc
            };
            items.Add(item);
            return item;
        }

        /// <summary>
        /// Removes the extracted copy of a duplicate item's file from disk and records
        /// the duplicate relationship/notes on the item for reporting purposes.
        /// </summary>
        /// <param name="dup">The duplicate item whose extracted file should be removed.</param>
        /// <param name="keeper">The original item that is being kept (the first copy encountered).</param>
        private void RemoveDuplicateFile(Item dup, Item keeper)
        {
            try
            {
                var fullPath = Path.Combine(_output, dup.ExtractedRelativePath);
                if (File.Exists(fullPath))
                {
                    File.Delete(fullPath);
                }
                dup.DuplicateRemoved = true;
                dup.DuplicateOfId = keeper.Id;
                dup.Status = "Duplicate (removed)";
                dup.Notes = string.IsNullOrEmpty(dup.Notes)
                    ? $"Identical content to {keeper.Id} ({keeper.OriginalName}); extracted copy removed to avoid duplication."
                    : dup.Notes;
            }
            catch (Exception ex)
            {
                dup.Notes = string.IsNullOrEmpty(dup.Notes) ? $"Failed to remove duplicate file: {ex.Message}" : dup.Notes;
            }
        }

        /// <summary>
        /// Sanitizes a string for safe use as (part of) a file name by replacing invalid
        /// file name characters and spaces with underscores.
        /// </summary>
        /// <param name="s">The string to sanitize.</param>
        /// <returns>A sanitized string safe for use in file names.</returns>
        private static string _sanitize(string s)
            => string.Join("_", s.Split(Path.GetInvalidFileNameChars())).Replace(' ', '_').Trim('_');
    }
}