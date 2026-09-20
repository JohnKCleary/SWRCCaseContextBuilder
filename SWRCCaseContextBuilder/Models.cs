// -----------------------------------------------------------------------
// SWRC Case Context Builder
// Author: John Cleary, Siptu Workers Rights Centre
// License: MIT
// -----------------------------------------------------------------------
using System;

namespace SWRCCaseContextBuilder
{
    /// <summary>
    /// Represents a single processed item (an original source file, a nested email/attachment,
    /// or an embedded document) tracked throughout the build pipeline. Instances of this class
    /// are used to populate the CSV/Excel manifest and the generated Word context reports.
    /// </summary>
    public class Item
    {
        /// <summary>Unique identifier assigned to this item (e.g. "E00001" for emails, "D00002" for other documents).</summary>
        public string Id { get; set; } = "";

        /// <summary>The <see cref="Id"/> of the parent item this was extracted from (e.g. the email an attachment came from), or empty for top-level items.</summary>
        public string ParentId { get; set; } = "";

        /// <summary>Nesting depth relative to the top-level source item (0 = top-level, 1 = direct attachment, etc.).</summary>
        public int Depth { get; set; }

        /// <summary>Path of the original file relative to the copied source root, or a descriptive path for nested/extracted items.</summary>
        public string SourceRelativePath { get; set; } = "";

        /// <summary>The original file name as found in the source folder or email/attachment.</summary>
        public string OriginalName { get; set; } = "";

        /// <summary>The sanitized, uniquely generated file name used for the extracted copy on disk.</summary>
        public string GeneratedName { get; set; } = "";

        /// <summary>The lowercase file extension, including the leading period (e.g. ".pdf").</summary>
        public string Extension { get; set; } = "";

        /// <summary>The size, in bytes, of the original file.</summary>
        public long SizeBytes { get; set; }

        /// <summary>The SHA-256 hash of the file's content, used for duplicate detection.</summary>
        public string Sha256 { get; set; } = "";

        /// <summary>The kind of item: "email" for .eml/.msg files, otherwise "file".</summary>
        public string Kind { get; set; } = "file";

        /// <summary>The email subject line, if this item is an email.</summary>
        public string Subject { get; set; } = "";

        /// <summary>The email sender address/display name, if this item is an email.</summary>
        public string Sender { get; set; } = "";

        /// <summary>The email "To" recipients, if this item is an email.</summary>
        public string Recipients { get; set; } = "";

        /// <summary>The email "Cc" recipients, if this item is an email.</summary>
        public string Cc { get; set; } = "";

        /// <summary>The email sent/received date as a display string, if this item is an email.</summary>
        public string SentDate { get; set; } = "";

        /// <summary>The email's Message-ID header value, if this item is an email.</summary>
        public string MessageId { get; set; } = "";

        /// <summary>The number of attachments found on this email, if this item is an email.</summary>
        public int AttachmentCount { get; set; }

        /// <summary>The current processing status (e.g. "Text extracted", "OCR required", "Exception").</summary>
        public string Status { get; set; } = "Pending";

        /// <summary>Identifier of the duplicate group this item belongs to, if its content hash matches other items.</summary>
        public string DuplicateGroup { get; set; } = "";

        /// <summary>Free-text notes, typically populated with error messages or duplicate/removal explanations.</summary>
        public string Notes { get; set; } = "";

        /// <summary>Path of the extracted copy of this item, relative to the output folder.</summary>
        public string ExtractedRelativePath { get; set; } = "";

        /// <summary>The number of characters of extracted/OCR text captured for this item.</summary>
        public int TextChars { get; set; }

        /// <summary>The extracted (or OCR'd) text content for this item.</summary>
        public string Text { get; set; } = "";

        /// <summary>
        /// True when this item's file content is a byte-for-byte duplicate of another item
        /// that appears earlier in processing order, and its extracted copy has been removed
        /// from disk to avoid storing duplicate content in the output.
        /// </summary>
        public bool DuplicateRemoved { get; set; }

        /// <summary>The <see cref="Id"/> of the item this one is a duplicate of, when <see cref="DuplicateRemoved"/> is true.</summary>
        public string DuplicateOfId { get; set; } = "";

        /// <summary>Parsed sent/received date, for emails, used for chronological sorting.</summary>
        public DateTimeOffset? SentDateValue { get; set; }

        /// <summary>
        /// Parsed date used for chronological sorting. For emails this is the sent/received date;
        /// for other documents it falls back to the file's last-modified timestamp.
        /// </summary>
        public DateTimeOffset? DocumentDate { get; set; }

        /// <summary>The effective date to use for chronological ordering: <see cref="SentDateValue"/> if present, otherwise <see cref="DocumentDate"/>.</summary>
        public DateTimeOffset? EffectiveDate => SentDateValue ?? DocumentDate;
    }

    /// <summary>
    /// A single progress update reported by <see cref="Builder.Run"/> during processing,
    /// used to update the UI's progress bar and status text.
    /// </summary>
    public class ProgressReport
    {
        /// <summary>Completion percentage (0-100), or -1 to indicate an indeterminate/unknown progress state.</summary>
        public int Percent { get; set; } = -1;

        /// <summary>A short human-readable description of the current processing step.</summary>
        public string Message { get; set; } = "";
    }

    /// <summary>
    /// Summary result returned by <see cref="Builder.Run"/> once processing has completed.
    /// </summary>
    public class RunResult
    {
        /// <summary>The total number of items processed (including nested attachments).</summary>
        public int Items { get; set; }

        /// <summary>The number of items that encountered an exception during processing.</summary>
        public int Exceptions { get; set; }

        /// <summary>The full path to the output folder that was created/populated.</summary>
        public string Output { get; set; } = "";
    }
}