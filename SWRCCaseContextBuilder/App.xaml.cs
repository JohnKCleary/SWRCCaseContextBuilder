// -----------------------------------------------------------------------
// SWRC Case Context Builder
// Author: John Cleary, Siptu Workers Rights Centre
// License: MIT
// -----------------------------------------------------------------------
using System;
using System.Configuration;
using System.Data;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace SWRCCaseContextBuilder
{
    /// <summary>
    /// Interaction logic for App.xaml. The WPF application entry point.
    /// Registers global exception handlers so that any unexpected startup or
    /// runtime failure is logged and surfaced to the user instead of causing
    /// the process to exit silently with no visible window (which can otherwise
    /// happen for self-contained/single-file deployments).
    /// </summary>
    public partial class App : System.Windows.Application
    {
        /// <summary>Full path to the crash log file written on unhandled exceptions.</summary>
        private static readonly string CrashLogPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "SWRCCaseContextBuilder", "crash.log");

        /// <summary>
        /// Wires up handlers for unhandled exceptions on the UI thread, background threads,
        /// and unobserved task exceptions before the rest of the application starts up.
        /// </summary>
        public App()
        {
            DispatcherUnhandledException += OnDispatcherUnhandledException;
            AppDomain.CurrentDomain.UnhandledException += OnAppDomainUnhandledException;
            TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;
        }

        /// <summary>Handles exceptions thrown on the WPF UI thread, logging them and showing an error dialog instead of letting the app silently terminate.</summary>
        private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            LogAndReport(e.Exception, "UI thread");
            e.Handled = true;
        }

        /// <summary>Handles exceptions thrown on non-UI threads (e.g. during startup or background work) that would otherwise crash the process silently.</summary>
        private void OnAppDomainUnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            LogAndReport(e.ExceptionObject as Exception, "Background thread");
        }

        /// <summary>Handles exceptions from unobserved (fire-and-forget) tasks, logging them so they don't go unnoticed.</summary>
        private void OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
        {
            LogAndReport(e.Exception, "Unobserved task");
            e.SetObserved();
        }

        /// <summary>
        /// Writes exception details to a crash log file (best effort) and shows a message box
        /// with the error, so failures are always visible to the user rather than resulting
        /// in the application exiting with no UI and no indication of what happened.
        /// </summary>
        /// <param name="ex">The exception to log/report.</param>
        /// <param name="source">A short description of where the exception originated.</param>
        private static void LogAndReport(Exception? ex, string source)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(CrashLogPath)!);
                File.AppendAllText(CrashLogPath, $"[{DateTime.Now:O}] ({source})\n{ex}\n\n");
            }
            catch { /* best effort logging only */ }

            try
            {
                System.Windows.MessageBox.Show(
                    $"An unexpected error occurred ({source}):\n\n{ex?.Message}\n\nDetails have been logged to:\n{CrashLogPath}",
                    "SWRC Case Context Builder - Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            catch { /* if even this fails, there's nothing more we can do */ }
        }
    }

}
