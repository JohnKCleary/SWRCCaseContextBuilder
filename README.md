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
