# SWRC Case Context Builder

A Windows desktop (WPF) tool that processes a folder of case files — emails, PDFs, Word/Excel
documents, and images — and builds a searchable, reviewable "context" for a case.

## What it does

- **Copies** the entire source folder to the output as a stable, unmodified reference copy.
- **Extracts** every file (including nested email attachments, to any depth) into a flat,
  uniquely named set of files.
- **Extracts text** from emails (`.eml`/`.msg`), PDFs, Word (`.docx`/`.docm`) and Excel
  (`.xlsx`/`.xlsm`) documents.
- **Runs OCR** (via the bundled Tesseract .NET engine) on scanned PDFs and image files
  (`.jpg`/`.png`/`.tif`/`.bmp`) where no text layer is found.
- **Detects duplicate files** (by SHA-256 content hash) and, if requested, removes duplicate
  copies from the output while noting the original in the report.
- **Generates a CSV and Excel manifest** of every item processed, plus one or more Word
  "context" reports combining all extracted text and metadata — optionally ordered
  chronologically and/or restricted to emails only.

## Typical usage

1. Select the source folder (e.g. a locally synced OneDrive case folder).
2. Choose where the output should go (or let it auto-create a folder in the source).
3. Optionally provide a case reference used to name the output folder and reports.
4. Choose which reports to generate and whether to remove duplicate files.
5. Click **Start** and wait for processing to complete.

The output folder contains:

```
00_Original_Source/     - unmodified copy of the source folder
01_Extracted_Files/     - flat, uniquely named extracted files (incl. nested attachments)
02_Context/             - CSV/Excel manifest + generated Word context report(s)
03_Review_Exceptions/   - reserved for items requiring manual review
```

## Report options

At least one report must be selected before processing can start. Reports are generated as
Word (`.docx`) documents in `02_Context/`, named `<CaseRef>.<ReportType>.docx`. Each report
contains a header summary (item counts, exceptions, duplicate info), an attachment/folder
tree map (except emails-only reports), and the full extracted record listing (metadata table
plus extracted/OCR text for every item).

| Option | File suffix | Contents / ordering |
|---|---|---|
| **All text (as processed)** | `.AllText.docx` | Every item (files, emails, and nested attachments), in the order they were found on disk (source folder walk order). Useful as a straightforward, complete record matching the original folder structure. |
| **All text (chronological)** | `.AllTextChronological.docx` | Every item, ordered oldest → newest by *effective date* (an email's sent/received date, or a document's last-modified date if it has no email date). Useful for building a timeline of events across all file types. |
| **All text (chronological, reversed)** | `.AllTextChronologicalReversed.docx` | Same as above but newest → oldest. Useful for quickly reviewing the most recent items first. |
| **Emails only (chronological)** | `.EmailsOnlyChronological.docx` | Only email items (`.eml`/`.msg`), oldest → newest by sent/received date. Non-email attachments and other documents are excluded, and the attachment/folder tree map is omitted (since it's not relevant to an emails-only view). Useful for reviewing a correspondence thread in isolation. |
| **Emails only (chronological, reversed)** | `.EmailsOnlyChronologicalReversed.docx` | Same as above but newest → oldest. |

You can select any combination of these five reports; each one is generated independently, so
selecting multiple has no effect on the others — it simply produces additional `.docx` files.

## Duplicate removal

The **"Remove duplicate files from output (keep first copy)"** checkbox controls how
byte-for-byte identical files (matched by SHA-256 content hash, across the *entire* case,
including nested email attachments) are handled:

- **Unchecked (default):** All items — including duplicates — are extracted and kept in
  `01_Extracted_Files/`. Every report includes all items; duplicate items are simply flagged
  in their metadata table with a shared `DuplicateGroup` identifier (visible in the CSV/Excel
  manifest) so you can identify them, but nothing is removed or hidden.
- **Checked:** For each group of identical files, only the *first* copy encountered
  (in file-walk order) is kept in `01_Extracted_Files/`; subsequent duplicate copies are
  deleted from the extracted output to save space and reduce review effort. Each removed
  duplicate's entry is retained in the manifest and in the standard reports, but its status
  is changed to `Duplicate (removed)` and a note records which item it duplicates
  (e.g. *"Identical content to E00003 (invoice.pdf); extracted copy removed to avoid
  duplication."*).

  **Additionally**, for every selected report type, a companion **"Distinct"** report is also
  generated (e.g. `<CaseRef>.AllText - Distinct.docx`) that *completely omits* any removed
  duplicate items — no mention of them at all — giving you a clean, deduplicated view of only
  the unique content in the case. This means checking the box can roughly double the number
  of `.docx` files produced (one standard report + one "Distinct" report per selected report
  type), though "Distinct" reports are only generated if at least one duplicate was actually
  found and removed.

**Important consequences of enabling duplicate removal:**

- The **original source copy** in `00_Original_Source/` is never affected — duplicates are
  only removed from the *extracted* copy in `01_Extracted_Files/`, so nothing is ever lost;
  you can always recover a removed duplicate's exact file from the original source folder
  copy or from the kept "original" item referenced in its notes.
- Duplicate detection is based purely on file **content** (SHA-256 hash), not file name —
  so two attachments with different names but identical content are still treated as
  duplicates, while two files with the same name but different content are not.
- The "keeper" of each duplicate group is always the first copy encountered in file-walk
  order (typically top-level files before nested attachments, processed alphabetically) —
  this is not configurable.

## To Test the app:
optionally download the zipped folder of Test_Input 
which contains examples of
**nested emails (3 levels)
emails with duplicate names
email attachments -pdf image, jpg - requiring OCR**

Compare output for the different output reports, emails only, chronological, remove duplicates etc
Practice on your own data

## Requirements

- Windows 10/11 (x64)
- No prerequisites required for the installed app — it is self-contained and bundles the
  .NET runtime, Tesseract, Leptonica, and SkiaSharp. The English OCR language data
  (`eng.traineddata`) is downloaded automatically on first run if not already available.

## Building from source

The solution consists of:

- **`SWRCCaseContextBuilder`** — the WPF application (.NET 8, `net8.0-windows`).
- **`SWRCCaseContextBuilder.Installer`** — a WiX v5 project that packages a self-contained,
  single-file publish of the app into an MSI installer.

To build and publish the app:

```powershell
dotnet publish SWRCCaseContextBuilder\SWRCCaseContextBuilder.csproj -c Release -r win-x64 `
  --self-contained -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true `
  -o SWRCCaseContextBuilder\bin\publish
```

To build the installer:

```powershell
dotnet build SWRCCaseContextBuilder.Installer\SWRCCaseContextBuilder.Installer.wixproj -c Release -p:Platform=x64
```

The resulting MSI is written to `SWRCCaseContextBuilder.Installer\bin\x64\Release\`.


## Key technologies

- WPF (.NET 8)
- [MimeKit](https://github.com/jstedfab/MimeKit) / [MsgReader](https://github.com/Sicos1977/MSGReader) — email parsing
- [PdfPig](https://github.com/UglyToad/PdfPig) / [PDFtoImage](https://github.com/sungaila/PDFtoImage) — PDF text extraction and rasterization
- [Tesseract](https://github.com/charlesw/tesseract) — OCR
- [DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK) — Word report generation
- [ClosedXML](https://github.com/ClosedXML/ClosedXML) — Excel manifest generation
- [WiX Toolset v5](https://wixtoolset.org/) — MSI installer

## License

Licensed under the [MIT License](LICENSE).

## Author

John Cleary, Siptu Workers Rights Centre
