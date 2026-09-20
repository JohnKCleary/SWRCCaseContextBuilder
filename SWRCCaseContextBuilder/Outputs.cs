// -----------------------------------------------------------------------
// SWRC Case Context Builder
// Author: John Cleary, Siptu Workers Rights Centre
// License: MIT
// -----------------------------------------------------------------------
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SWRCCaseContextBuilder
{
    /// <summary>
    /// Selects which Word context report(s) should be generated. Multiple flags can be combined.
    /// </summary>
    [Flags]
    public enum OutputTypes
    {
        /// <summary>No reports selected.</summary>
        None = 0,
        /// <summary>"&lt;CaseRef&gt;.AllText.docx" - as processed.</summary>
        AllText = 1,
        /// <summary>"&lt;CaseRef&gt;.AllTextChronological.docx".</summary>
        AllTextChronological = 2,
        /// <summary>"&lt;CaseRef&gt;.AllTextChronologicalReversed.docx".</summary>
        AllTextChronologicalReversed = 4,
        /// <summary>"&lt;CaseRef&gt;.EmailsOnlyChronological.docx".</summary>
        EmailsOnlyChronological = 8,
        /// <summary>"&lt;CaseRef&gt;.EmailsOnlyChronologicalReversed.docx".</summary>
        EmailsOnlyChronologicalReversed = 16
    }

    /// <summary>
    /// Generates the manifest (CSV/Excel) and Word context report outputs for a completed
    /// processing run, based on the requested <see cref="OutputTypes"/>.
    /// </summary>
    public static class Outputs
    {
        /// <summary>Light blue shading colour (hex) used for label cells in metadata tables.</summary>
        private const string ShadeColor = "D9E2F3"; // light blue shading for label cells

        /// <summary>
        /// Writes the CSV and Excel manifests summarising all items, and generates the
        /// requested Word context report(s) into the output folder's "02_Context" subfolder.
        /// </summary>
        /// <param name="items">The full list of processed items.</param>
        /// <param name="outputFolder">The root output folder for the run.</param>
        /// <param name="caseRef">Case reference used to name generated report files.</param>
        /// <param name="version">Builder version string included in report headers.</param>
        /// <param name="outputTypes">Which Word context report(s) to generate.</param>
        /// <param name="removeDuplicates">Whether duplicate removal was enabled, controlling whether companion "Distinct" reports are also produced.</param>
        public static void WriteOutputs(List<Item> items, string outputFolder, string caseRef, string version, OutputTypes outputTypes = OutputTypes.AllText, bool removeDuplicates = false)
        {
            var contextFolder = Path.Combine(outputFolder, "02_Context");
            Directory.CreateDirectory(contextFolder);

            // CSV
            var csv = Path.Combine(contextFolder, "Case_Manifest.csv");
            var fields = new[] { "Id","ParentId","Depth","SourceRelativePath","OriginalName","GeneratedName","Extension","SizeBytes","Sha256","Kind","Subject","Sender","Recipients","Cc","SentDate","MessageId","AttachmentCount","Status","DuplicateGroup","Notes","ExtractedRelativePath","TextChars" };

            using var sw = new StreamWriter(csv, false, System.Text.Encoding.UTF8);
            sw.WriteLine(string.Join(",", fields));
            foreach (var it in items)
            {
                var vals = fields.Select(f => (typeof(Item).GetProperty(f)?.GetValue(it) ?? "").ToString()!.Replace("\"","\"\""));
                sw.WriteLine($"\"{string.Join("\",\"", vals)}\"");
            }

            // Excel via ClosedXML
            var wbPath = Path.Combine(contextFolder, "Case_Manifest.xlsx");
            var wb = new ClosedXML.Excel.XLWorkbook();
            var ws = wb.Worksheets.Add("Manifest");
            for (int i = 0; i < fields.Length; i++) ws.Cell(1, i + 1).Value = fields[i];
            for (int r = 0; r < items.Count; r++)
            {
                var row = items[r];
                for (int c = 0; c < fields.Length; c++)
                {
                    var v = typeof(Item).GetProperty(fields[c])?.GetValue(row);
                    ws.Cell(r + 2, c + 1).Value = v is null ? XLCellValue.FromObject(string.Empty) : XLCellValue.FromObject(v);
                }
            }
            wb.SaveAs(wbPath);

            // Word context report(s) via DocumentFormat.OpenXml (no license required)
            var baseName = SanitizeFileNamePart(string.IsNullOrWhiteSpace(caseRef) ? "Output" : caseRef);

            if (outputTypes.HasFlag(OutputTypes.AllText))
            {
                GenerateWordReport(items, contextFolder, $"{baseName}.AllText.docx", caseRef, version,
                    title: "Case Context", orderDescription: "As processed", emailsOnly: false, removeDuplicates: removeDuplicates);
            }

            if (outputTypes.HasFlag(OutputTypes.AllTextChronological))
            {
                var ordered = OrderChronologically(items, reversed: false);
                GenerateWordReport(ordered, contextFolder, $"{baseName}.AllTextChronological.docx", caseRef, version,
                    title: "Case Context — Chronological", orderDescription: "Chronological (by email/document date)", emailsOnly: false, removeDuplicates: removeDuplicates);
            }

            if (outputTypes.HasFlag(OutputTypes.AllTextChronologicalReversed))
            {
                var ordered = OrderChronologically(items, reversed: true);
                GenerateWordReport(ordered, contextFolder, $"{baseName}.AllTextChronologicalReversed.docx", caseRef, version,
                    title: "Case Context — Chronological (Reversed)", orderDescription: "Reverse chronological (by email/document date)", emailsOnly: false, removeDuplicates: removeDuplicates);
            }

            if (outputTypes.HasFlag(OutputTypes.EmailsOnlyChronological))
            {
                var ordered = OrderChronologically(items.Where(i => i.Kind == "email").ToList(), reversed: false);
                GenerateWordReport(ordered, contextFolder, $"{baseName}.EmailsOnlyChronological.docx", caseRef, version,
                    title: "Case Context — Emails Only, Chronological", orderDescription: "Chronological (emails only, by date)", emailsOnly: true, removeDuplicates: removeDuplicates);
            }

            if (outputTypes.HasFlag(OutputTypes.EmailsOnlyChronologicalReversed))
            {
                var ordered = OrderChronologically(items.Where(i => i.Kind == "email").ToList(), reversed: true);
                GenerateWordReport(ordered, contextFolder, $"{baseName}.EmailsOnlyChronologicalReversed.docx", caseRef, version,
                    title: "Case Context — Emails Only, Chronological (Reversed)", orderDescription: "Reverse chronological (emails only, by date)", emailsOnly: true, removeDuplicates: removeDuplicates);
            }
        }

        /// <summary>Orders items by their <see cref="Item.EffectiveDate"/>, oldest first (or newest first when reversed).</summary>
        /// <param name="items">The items to order.</param>
        /// <param name="reversed">When true, returns newest first instead of oldest first.</param>
        /// <returns>A new list of items in the requested chronological order.</returns>
        private static List<Item> OrderChronologically(List<Item> items, bool reversed)
        {
            var ordered = items.OrderBy(i => i.EffectiveDate ?? i.SentDateValue ?? i.DocumentDate ?? DateTimeOffset.MinValue).ToList();
            if (reversed) ordered.Reverse();
            return ordered;
        }

        /// <summary>Replaces any characters invalid in file names with underscores.</summary>
        /// <param name="value">The string to sanitize.</param>
        /// <returns>A sanitized string safe for use as (part of) a file name.</returns>
        private static string SanitizeFileNamePart(string value)
        {
            var invalid = Path.GetInvalidFileNameChars();
            return new string(value.Select(c => invalid.Contains(c) ? '_' : c).ToArray()).Trim();
        }

        /// <summary>
        /// Generates the main report plus, when duplicate removal is enabled, a companion
        /// "Distinct" report that omits any items whose content was removed as a duplicate.
        /// </summary>
        /// <param name="items">The items to include in the report.</param>
        /// <param name="contextFolder">The "02_Context" folder to write the report(s) into.</param>
        /// <param name="fileName">The file name for the main report.</param>
        /// <param name="caseRef">Case reference shown in the report header.</param>
        /// <param name="version">Builder version shown in the report header.</param>
        /// <param name="title">The report's main title.</param>
        /// <param name="orderDescription">Description of the record order, shown in the report header.</param>
        /// <param name="emailsOnly">Whether this report includes only email items.</param>
        /// <param name="removeDuplicates">Whether duplicate removal was enabled for the run.</param>
        private static void GenerateWordReport(List<Item> items, string contextFolder, string fileName, string caseRef, string version, string title, string orderDescription, bool emailsOnly, bool removeDuplicates)
        {
            var docPath = Path.Combine(contextFolder, fileName);
            var exceptionsCount = items.Count(i => i.Status == "Exception");
            WriteWordContext(docPath, items, caseRef, version, exceptionsCount, title, orderDescription, emailsOnly, isDistinctOnly: false);

            // When duplicate removal is enabled, also produce a second "Distinct" report containing
            // only unique items - no mention of removed duplicates, no wasted space on them at all.
            if (removeDuplicates && items.Any(i => i.DuplicateRemoved))
            {
                var distinctItems = items.Where(i => !i.DuplicateRemoved).ToList();
                var distinctFileName = Path.GetFileNameWithoutExtension(fileName) + " - Distinct" + Path.GetExtension(fileName);
                var distinctDocPath = Path.Combine(contextFolder, distinctFileName);
                var distinctExceptionsCount = distinctItems.Count(i => i.Status == "Exception");
                WriteWordContext(distinctDocPath, distinctItems, caseRef, version, distinctExceptionsCount, title + " — Distinct Items", orderDescription, emailsOnly, isDistinctOnly: true);
            }
        }

        /// <summary>
        /// Builds and saves a single Word context report document, including the header summary,
        /// an optional attachment/folder tree map, and the full extracted record listing.
        /// </summary>
        /// <param name="docPath">Full path where the .docx file will be saved.</param>
        /// <param name="items">The items to include in the report body.</param>
        /// <param name="caseRef">Case reference shown in the report header.</param>
        /// <param name="version">Builder version shown in the report header.</param>
        /// <param name="exceptionsCount">Number of items with an "Exception" status, shown in the header summary.</param>
        /// <param name="title">The report's main title.</param>
        /// <param name="orderDescription">Description of the record order, shown in the report header.</param>
        /// <param name="emailsOnly">When true, omits the attachment/folder tree map section (not applicable for emails-only reports).</param>
        /// <param name="isDistinctOnly">When true, renders header text/labels appropriate for a "Distinct" (duplicate-free) companion report.</param>
        private static void WriteWordContext(string docPath, List<Item> items, string caseRef, string version, int exceptionsCount, string title, string orderDescription, bool emailsOnly, bool isDistinctOnly)
        {
            using var wordDoc = WordprocessingDocument.Create(docPath, WordprocessingDocumentType.Document);
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());

            body.AppendChild(CreateParagraph(title, bold: true, fontSize: 32));

            if (isDistinctOnly)
            {
                body.AppendChild(CreateParagraph($"Case reference: {caseRef}\nGenerated: {DateTime.Now:O}\nBuilder version: {version}\nRecord order: {orderDescription}"));
                body.AppendChild(CreateParagraph($"Unique items: {items.Count}\nExceptions requiring review: {exceptionsCount}\nThis report contains only distinct (non-duplicate) items; duplicate copies have been fully omitted.\nThe source is a copied, locally synchronised OneDrive folder."));
            }
            else
            {
                var duplicatesRemovedCount = items.Count(i => i.DuplicateRemoved);
                body.AppendChild(CreateParagraph($"Case reference: {caseRef}\nGenerated: {DateTime.Now:O}\nBuilder version: {version}\nRecord order: {orderDescription}"));
                body.AppendChild(CreateParagraph($"Items processed: {items.Count}\nExceptions requiring review: {exceptionsCount}\nDuplicate files removed: {duplicatesRemovedCount}\nThe source is a copied, locally synchronised OneDrive folder."));
            }

            if (!emailsOnly)
            {
                body.AppendChild(CreateParagraph("Attachment and folder map", bold: true, fontSize: 24, spacingBefore: 200));
                body.AppendChild(CreateParagraph(
                    "Top-level items are the original emails/documents found in the source folder. Indented items are attachments extracted from the email immediately above them.",
                    italic: true, color: "595959", spacingBefore: 60));

                // build tree recursively (supports arbitrary depth, not just one level)
                // Note: when isDistinctOnly is true, items list already excludes removed duplicates,
                // so any branch rooted at a removed duplicate is naturally omitted here too.
                var childrenByParent = items.GroupBy(i => i.ParentId).ToDictionary(g => g.Key, g => g.ToList());
                if (childrenByParent.TryGetValue("", out var roots))
                {
                    foreach (var root in roots)
                    {
                        AppendTreeNode(body, root, childrenByParent, 0);
                    }
                }

                body.AppendChild(CreatePageBreakParagraph());
            }

            body.AppendChild(CreateParagraph("Extracted records", bold: true, fontSize: 28));

            foreach (var item in items)
            {
                body.AppendChild(CreateParagraph($"{item.Id}: {item.OriginalName}", bold: true, fontSize: 24, spacingBefore: 200));

                var metaRows = new List<(string Label, string Value)>
                {
                    ("Source path", item.SourceRelativePath),
                    ("Parent", string.IsNullOrEmpty(item.ParentId) ? "Top level" : item.ParentId)
                };
                if (item.Depth > 0) metaRows.Add(("Depth", item.Depth.ToString()));
                if (!string.IsNullOrEmpty(item.Sha256)) metaRows.Add(("SHA-256", item.Sha256));
                metaRows.Add(("Status", item.Status));
                if (!isDistinctOnly && item.DuplicateRemoved && !string.IsNullOrEmpty(item.DuplicateOfId))
                {
                    metaRows.Add(("Duplicate of", item.DuplicateOfId));
                }

                if (item.Kind == "email")
                {
                    if (!string.IsNullOrEmpty(item.SentDate)) metaRows.Add(("Date", item.SentDate));
                    if (!string.IsNullOrEmpty(item.Sender)) metaRows.Add(("From", item.Sender));
                    if (!string.IsNullOrEmpty(item.Recipients)) metaRows.Add(("To", item.Recipients));
                    if (!string.IsNullOrEmpty(item.Cc)) metaRows.Add(("Cc", item.Cc));
                    if (!string.IsNullOrEmpty(item.Subject)) metaRows.Add(("Subject", item.Subject));
                    if (item.AttachmentCount > 0) metaRows.Add(("Attachments", item.AttachmentCount.ToString()));
                }
                if (!string.IsNullOrEmpty(item.Notes)) metaRows.Add(("Notes", item.Notes));

                body.AppendChild(CreateMetadataTable(metaRows));
                body.AppendChild(CreateParagraph(item.Text ?? "[No text extracted]", spacingBefore: 120));
                body.AppendChild(CreatePageBreakParagraph());
            }

            mainPart.Document.Save();
        }

        /// <summary>
        /// Recursively appends a node (and all its descendants) to the attachment/folder tree map,
        /// indenting child nodes to reflect their nesting depth.
        /// </summary>
        /// <param name="body">The document body to append paragraphs to.</param>
        /// <param name="node">The item to render as a tree row.</param>
        /// <param name="childrenByParent">Lookup of child items keyed by their parent's <see cref="Item.Id"/>.</param>
        /// <param name="indent">The current nesting depth (0 = top-level).</param>
        private static void AppendTreeNode(Body body, Item node, Dictionary<string, List<Item>> childrenByParent, int indent)
        {
            var isTopLevel = indent == 0;
            var hasChildren = childrenByParent.ContainsKey(node.Id);

            var kindLabel = node.Kind == "email" ? "EMAIL" : node.Extension.TrimStart('.').ToUpperInvariant();
            var roleLabel = isTopLevel ? "Source item" : "Attachment";
            var kindColor = node.Kind == "email" ? "2E5FA3" : "6B6B6B";
            var statusColor = GetStatusColor(node.Status);

            body.AppendChild(CreateTreeRow(node, indent, isTopLevel, kindLabel, roleLabel, kindColor, statusColor));

            if (hasChildren)
            {
                foreach (var c in childrenByParent[node.Id]) AppendTreeNode(body, c, childrenByParent, indent + 1);
            }
        }

        /// <summary>Maps an item's processing status to a colour (hex) used to highlight it in the tree map.</summary>
        /// <param name="status">The item's <see cref="Item.Status"/> value.</param>
        /// <returns>A hex colour string appropriate for the given status.</returns>
        private static string GetStatusColor(string status) => status switch
        {
            "Email extracted" => "2E7D32",
            "Text extracted" => "2E7D32",
            "Text extracted (OCR)" => "2E7D32",
            "OCR required" => "B8860B",
            "Copied only" => "6B6B6B",
            "Duplicate (removed)" => "8E44AD",
            "Exception" => "C0392B",
            _ => "6B6B6B"
        };

        /// <summary>
        /// Renders a single tree entry as a compact, indented paragraph combining:
        /// a connector/indent prefix showing nesting depth, a role badge (Source item / Attachment)
        /// so parents vs attachments are unambiguous, and the item Id, name, a kind badge
        /// (EMAIL/PDF/JPG/etc.) and a colour-coded status badge.
        /// </summary>
        /// <param name="node">The item being rendered.</param>
        /// <param name="indent">The nesting depth, used to compute paragraph indentation.</param>
        /// <param name="isTopLevel">Whether this node is a top-level source item (vs. a nested attachment).</param>
        /// <param name="kindLabel">The badge text describing the item kind (e.g. "EMAIL", "PDF").</param>
        /// <param name="roleLabel">The badge text describing the item's role ("Source item" or "Attachment").</param>
        /// <param name="kindColor">The hex colour used for the kind badge.</param>
        /// <param name="statusColor">The hex colour used for the status badge.</param>
        /// <returns>A formatted <see cref="Paragraph"/> representing the tree row.</returns>
        private static Paragraph CreateTreeRow(Item node, int indent, bool isTopLevel, string kindLabel, string roleLabel, string kindColor, string statusColor)
        {
            var paragraph = new Paragraph();
            var paraProps = new ParagraphProperties(
                new Indentation { Left = (indent * 360).ToString() },
                new SpacingBetweenLines { Before = isTopLevel ? "160" : "40", After = "40" }
            );
            paragraph.ParagraphProperties = paraProps;

            if (!isTopLevel)
            {
                paragraph.AppendChild(CreateRun("\u21B3 ", color: "999999"));
            }

            paragraph.AppendChild(CreateRun($"[{roleLabel}] ", bold: true, color: isTopLevel ? "2E5FA3" : "8A8A8A", fontSize: 16));
            paragraph.AppendChild(CreateRun($"{node.Id} ", bold: true));
            paragraph.AppendChild(CreateRun($"{node.OriginalName}  ", bold: isTopLevel));
            paragraph.AppendChild(CreateRun($"[{kindLabel}] ", bold: true, color: kindColor, fontSize: 16));
            paragraph.AppendChild(CreateRun($"[{node.Status}]", bold: true, color: statusColor, fontSize: 16));

            return paragraph;
        }

        /// <summary>Creates a formatted <see cref="Run"/> of text with optional bold, colour and font size.</summary>
        /// <param name="text">The text content of the run.</param>
        /// <param name="bold">Whether to render the text bold.</param>
        /// <param name="color">Optional hex colour for the text.</param>
        /// <param name="fontSize">Optional font size (in half-points, per OpenXML convention).</param>
        /// <returns>The constructed <see cref="Run"/>.</returns>
        private static Run CreateRun(string text, bool bold = false, string? color = null, int? fontSize = null)
        {
            var run = new Run();
            var runProps = new RunProperties();
            if (bold) runProps.Append(new Bold());
            if (!string.IsNullOrEmpty(color)) runProps.Append(new DocumentFormat.OpenXml.Wordprocessing.Color { Val = color });
            if (fontSize.HasValue) runProps.Append(new FontSize { Val = fontSize.Value.ToString() });
            if (runProps.HasChildren) run.RunProperties = runProps;
            run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            return run;
        }

        /// <summary>
        /// Builds a two-column metadata table (label / value) with shaded label cells and borders,
        /// matching the visual style of the Python-generated report.
        /// </summary>
        /// <param name="rows">The label/value pairs to render as table rows.</param>
        /// <returns>The constructed <see cref="Table"/>.</returns>
        private static Table CreateMetadataTable(List<(string Label, string Value)> rows)
        {
            var table = new Table();

            var tblProps = new TableProperties(
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 6, Color = "BFBFBF" },
                    new BottomBorder { Val = BorderValues.Single, Size = 6, Color = "BFBFBF" },
                    new LeftBorder { Val = BorderValues.Single, Size = 6, Color = "BFBFBF" },
                    new RightBorder { Val = BorderValues.Single, Size = 6, Color = "BFBFBF" },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 6, Color = "BFBFBF" },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 6, Color = "BFBFBF" }
                ),
                new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" }
            );
            table.AppendChild(tblProps);

            var grid = new TableGrid(
                new GridColumn { Width = "2000" },
                new GridColumn { Width = "7000" }
            );
            table.AppendChild(grid);

            foreach (var (label, value) in rows)
            {
                var tr = new TableRow();

                var labelCell = new TableCell();
                labelCell.Append(new TableCellProperties(
                    new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2000" },
                    new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = ShadeColor }
                ));
                labelCell.Append(CreateParagraph(label, bold: true, keepMargins: true));
                tr.Append(labelCell);

                var valueCell = new TableCell();
                valueCell.Append(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "7000" }));
                valueCell.Append(CreateParagraph(value ?? "", keepMargins: true));
                tr.Append(valueCell);

                table.Append(tr);
            }

            return table;
        }

        /// <summary>Creates a paragraph of text, optionally bold/italic/coloured/sized, supporting multi-line text via embedded line breaks.</summary>
        /// <param name="text">The text to render (may contain '\n' for line breaks).</param>
        /// <param name="bold">Whether to render the text bold.</param>
        /// <param name="fontSize">Optional font size (in half-points).</param>
        /// <param name="spacingBefore">Optional spacing (in twentieths of a point) before the paragraph.</param>
        /// <param name="keepMargins">When true, skips applying <paramref name="spacingBefore"/> (used for table cell content).</param>
        /// <param name="italic">Whether to render the text italic.</param>
        /// <param name="color">Optional hex colour for the text.</param>
        /// <returns>The constructed <see cref="Paragraph"/>.</returns>
        private static Paragraph CreateParagraph(string text, bool bold = false, int? fontSize = null, int? spacingBefore = null, bool keepMargins = false, bool italic = false, string? color = null)
        {
            var run = new Run();
            var runProps = new RunProperties();
            if (bold) runProps.Append(new Bold());
            if (italic) runProps.Append(new Italic());
            if (!string.IsNullOrEmpty(color)) runProps.Append(new DocumentFormat.OpenXml.Wordprocessing.Color { Val = color });
            if (fontSize.HasValue) runProps.Append(new FontSize { Val = fontSize.Value.ToString() });
            if (runProps.HasChildren) run.RunProperties = runProps;

            var lines = (text ?? "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                if (i > 0) run.AppendChild(new Break());
                run.AppendChild(new Text(lines[i]) { Space = SpaceProcessingModeValues.Preserve });
            }

            var paragraph = new Paragraph();
            if (spacingBefore.HasValue && !keepMargins)
            {
                paragraph.ParagraphProperties = new ParagraphProperties(new SpacingBetweenLines { Before = spacingBefore.Value.ToString() });
            }
            paragraph.AppendChild(run);
            return paragraph;
        }

        /// <summary>Creates a paragraph containing a single page break.</summary>
        /// <returns>The constructed <see cref="Paragraph"/>.</returns>
        private static Paragraph CreatePageBreakParagraph()
        {
            var run = new Run(new Break { Type = BreakValues.Page });
            var paragraph = new Paragraph();
            paragraph.AppendChild(run);
            return paragraph;
        }
    }
}