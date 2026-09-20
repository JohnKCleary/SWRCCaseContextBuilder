// -----------------------------------------------------------------------
// SWRC Case Context Builder
// Author: John Cleary, Siptu Workers Rights Centre
// License: MIT
// -----------------------------------------------------------------------
using Microsoft.Win32;
using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;

namespace SWRCCaseContextBuilder
{
    /// <summary>
    /// Main application window. Lets the user select a source case folder and output
    /// location, choose which reports to generate, and kicks off a <see cref="Builder"/>
    /// run while reporting progress back to the UI.
    /// </summary>
    public partial class MainWindow : Window
    {
        /// <summary>
        /// Initializes the window and sets the initial output controls state. The OCR
        /// (Tesseract) availability check is kicked off asynchronously in the background
        /// (see <see cref="CheckTesseractAtStartupAsync"/>) so that a slow or unavailable
        /// network connection can never delay or prevent the window from appearing.
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
            UpdateOutputControlsState();
            _ = CheckTesseractAtStartupAsync();
        }

        /// <summary>Handles the "Browse..." click for the source folder, opening a folder picker dialog.</summary>
        private void BtnBrowseSource_Click(object sender, RoutedEventArgs e)
        {
            using var dlg = new FolderBrowserDialog { Description = "Select locally synchronised OneDrive case folder" };
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK) TxtSource.Text = dlg.SelectedPath;
        }

        /// <summary>Handles the "Help" button click, showing a summary of the application's functionality.</summary>
        private void BtnHelp_Click(object sender, RoutedEventArgs e)
        {
            const string helpText =
                "SWRC Case File Helper\n\n" +
                "This application processes a folder of case files (emails, PDFs, Word/Excel documents, images) " +
                "and builds a searchable, reviewable 'context' for a case.\n\n" +
                "What it does:\n" +
                "- Copies the entire source folder to the output as a stable, unmodified reference copy.\n" +
                "- Extracts every file (including nested email attachments, to any depth) into a flat, uniquely named set of files.\n" +
                "- Extracts text from emails (.eml/.msg), PDFs, Word (.docx/.docm) and Excel (.xlsx/.xlsm) documents.\n" +
                "- Runs Optical Character Recognition ((OCR), via Tesseract) on scanned PDFs and image files (.jpg/.png/.tif/.bmp) where no text layer is found.\n" +
                "- If the OCR engine can read text in the scanned pdf or image, it extracts the text and copies it into the output files\n" +
                "- Detects duplicate files (by content hash, so only exactly matching file duplicates) and, if requested, removes duplicate copies from the output while noting the original in the report.\n" +
                "- Generates a CSV and Excel manifest of every item processed, plus one or more Word 'context' reports " +
                "combining all extracted text and metadata, optionally ordered chronologically and/or restricted to emails only.\n\n" +
                "Typical usage:\n" +
                "1. Select the source folder (e.g. a locally synced OneDrive case folder).\n" +
                "2. Choose where the output should go (or let it auto-create a folder in the source).\n" +
                "3. Optionally provide a case reference used to name the output folder and reports.\n" +
                "4. Choose which reports to generate and whether to remove duplicate files.\n" +
                "5. Click Start and wait for processing to complete.";

            System.Windows.MessageBox.Show(helpText, "About SWRC Case File Helper", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        /// <summary>Handles the "Browse..." click for the output folder, opening a folder picker dialog.</summary>
        private void BtnBrowseOutput_Click(object sender, RoutedEventArgs e)
        {
            using var dlg = new FolderBrowserDialog { Description = "Select an EMPTY output folder" };
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK) TxtOutput.Text = dlg.SelectedPath;
        }

        /// <summary>Handles the "Create output folder in source" checkbox toggling, updating the output controls' enabled state.</summary>
        private void ChkCreateOutputInSource_CheckedChanged(object sender, RoutedEventArgs e)
        {
            UpdateOutputControlsState();
        }

        /// <summary>Enables/disables the manual output folder controls based on whether auto-created output is selected.</summary>
        private void UpdateOutputControlsState()
        {
            var autoOutput = ChkCreateOutputInSource.IsChecked == true;
            TxtOutput.IsEnabled = !autoOutput;
            BtnBrowseOutput.IsEnabled = !autoOutput;
            if (autoOutput)
            {
                TxtOutput.Text = "";
            }
        }

        /// <summary>
        /// Handles the "Start" button click: validates inputs, determines the output path,
        /// builds the selected <see cref="OutputTypes"/> flags, and runs the <see cref="Builder"/>
        /// asynchronously, updating the progress bar/status text and optionally opening the
        /// output folder on completion.
        /// </summary>
        private async void BtnStart_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var source = TxtSource.Text;
                var caseRef = string.IsNullOrWhiteSpace(TxtCaseRef.Text) ? Path.GetFileName(source ?? "") : TxtCaseRef.Text;

                if (string.IsNullOrWhiteSpace(source))
                {
                    System.Windows.MessageBox.Show("Select a source folder.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                string output;
                if (ChkCreateOutputInSource.IsChecked == true)
                {
                    var folderNamePart = string.IsNullOrWhiteSpace(TxtCaseRef.Text) ? Path.GetFileName(source) : TxtCaseRef.Text;
                    var safeFolderNamePart = string.Join("_", folderNamePart.Split(Path.GetInvalidFileNameChars()));
                    output = Path.Combine(source, $"{safeFolderNamePart}_Output");
                }
                else
                {
                    output = TxtOutput.Text;
                    if (string.IsNullOrWhiteSpace(output))
                    {
                        System.Windows.MessageBox.Show("Select an output folder, or check 'Create output folder in source'.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                }

                var outputTypes = OutputTypes.None;
                if (ChkOutputAllText.IsChecked == true) outputTypes |= OutputTypes.AllText;
                if (ChkOutputAllTextChronological.IsChecked == true) outputTypes |= OutputTypes.AllTextChronological;
                if (ChkOutputAllTextChronologicalReversed.IsChecked == true) outputTypes |= OutputTypes.AllTextChronologicalReversed;
                if (ChkOutputEmailsOnlyChronological.IsChecked == true) outputTypes |= OutputTypes.EmailsOnlyChronological;
                if (ChkOutputEmailsOnlyChronologicalReversed.IsChecked == true) outputTypes |= OutputTypes.EmailsOnlyChronologicalReversed;

                if (outputTypes == OutputTypes.None)
                {
                    System.Windows.MessageBox.Show("Select at least one output report to generate.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                TxtStatus.Text = "Preparing...";
                ProgressBar.IsIndeterminate = true;

                var removeDuplicates = ChkRemoveDuplicates.IsChecked == true;

                var builder = new Builder(source, output, caseRef,
                    progress => Dispatcher.Invoke(() => { if (progress.Percent >= 0) { ProgressBar.IsIndeterminate = false; ProgressBar.Value = progress.Percent; } TxtStatus.Text = progress.Message; }),
                    outputTypes, removeDuplicates);

                var result = await Task.Run(() => builder.Run());

                TxtStatus.Text = $"Completed: {result.Items} items, {result.Exceptions} exceptions.";
                ProgressBar.Value = 100;

                if (ChkOpenOnFinish.IsChecked == true)
                {
                    try { System.Diagnostics.Process.Start("explorer", result.Output); }
                    catch { /* best effort */ }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.ToString(), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                ProgressBar.IsIndeterminate = false;
            }
        }

        /// <summary>Handles the "Exit" button click, closing the application window.</summary>
        private void BtnExit_Click(object sender, RoutedEventArgs e) => Close();

        /// <summary>
        /// Asynchronously checks whether the OCR (Tesseract) engine's language data is available
        /// (downloading it in the background if needed) and updates the info text block accordingly.
        /// Runs off the UI thread and is fully exception-safe so that a slow/unreachable network
        /// (e.g. no internet access, blocked outbound HTTPS) can never hang or crash the UI, and the
        /// main window always appears immediately regardless of OCR data availability.
        /// </summary>
        private async Task CheckTesseractAtStartupAsync()
        {
            try
            {
                var (ok, details) = await Task.Run(() =>
                {
                    var success = Utils.EnsureTessDataAvailable(out var d);
                    return (success, d);
                });

                Dispatcher.Invoke(() =>
                {
                    TxtInfo.Text += "\n\nOCR: " + (ok ? $"Ready. {details}" : $"Not available. {details}");
                });
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                {
                    TxtInfo.Text += $"\n\nOCR: Not available. Startup check failed: {ex.Message}";
                });
            }
        }

        /// <summary>No-op handler for the "Remove duplicate files" checkbox (state is read directly at Start time).</summary>
        private void ChkRemoveDuplicates_Checked(object sender, RoutedEventArgs e)
        {
        
        }
    }
}