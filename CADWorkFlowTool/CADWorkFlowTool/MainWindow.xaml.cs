using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Forms = System.Windows.Forms;
using MicrosoftWin32 = Microsoft.Win32;

namespace CADWorkFlowTool
{
    public partial class MainWindow : System.Windows.Window
    {
        private int _currentStep = 1;

        public MainWindow()
        {
            InitializeComponent();

            // Wire up buttons
            BtnRunPreprocessor.Click += BtnRunPreprocessor_Click;
            BtnRunExtraction.Click += BtnRunExtraction_Click;
            BtnRunComparison.Click += BtnRunComparison_Click;
            BtnSkipPreprocessor.Click += BtnSkipPreprocessor_Click;
            BtnSkipExtraction.Click += BtnSkipExtraction_Click;
            BtnBrowsePreInput.Click += BtnBrowsePreInput_Click;
            BtnBrowsePreOutput.Click += BtnBrowsePreOutput_Click;
            BtnBrowseExOutput.Click += BtnBrowseExOutput_Click;
            BtnBrowseCmpBaseline.Click += BtnBrowseCmpBaseline_Click;
            BtnBrowseCmpOutput.Click += BtnBrowseCmpOutput_Click;
            BtnSaveLog.Click += BtnSaveLog_Click;
            BtnPreviousToStep1.Click += BtnPreviousToStep1_Click;
            BtnPreviousToStep2.Click += BtnPreviousToStep2_Click;

            // New: reference extraction button
            BtnRunRefExtraction.Click += BtnRunRefExtraction_Click;

            UpdateWizardState(1);
            LogBox.AppendText("[INFO] Welcome! Please provide paths for the Preprocessor.\n");
        }

        private void UpdateWizardState(int step)
        {
            _currentStep = step;

            Step1_Preprocessor.Visibility = (step == 1) ? Visibility.Visible : Visibility.Collapsed;
            Step2_Extraction.Visibility = (step == 2) ? Visibility.Visible : Visibility.Collapsed;
            Step3_Comparison.Visibility = (step == 3) ? Visibility.Visible : Visibility.Collapsed;

            if (step == 1) StepperText.Text = "Step 1 of 3: Preprocessor";
            else if (step == 2) StepperText.Text = "Step 2 of 3: Data Extraction";
            else if (step == 3) StepperText.Text = "Step 3 of 3: XML Comparison";

            if (step < 3)
            {
                BtnRunComparison.Content = "Run Comparison & Finish";
                BtnRunComparison.Background = System.Windows.Media.Brushes.DodgerBlue;
                BtnRunComparison.IsEnabled = true;
            }
        }

        // Helper to pick folder and set textbox
        private void SelectFolder(System.Windows.Controls.TextBox targetTextBox)
        {
            using (var dialog = new Forms.FolderBrowserDialog())
            {
                var result = dialog.ShowDialog();
                if (result == Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedPath))
                {
                    targetTextBox.Text = dialog.SelectedPath;
                }
            }
        }

        // Browse handlers
        private void BtnBrowsePreInput_Click(object sender, RoutedEventArgs e) => SelectFolder(TxtPreprocessorInput);
        private void BtnBrowsePreOutput_Click(object sender, RoutedEventArgs e) => SelectFolder(TxtPreprocessorOutput);
        private void BtnBrowseExOutput_Click(object sender, RoutedEventArgs e) => SelectFolder(TxtExtractionOutput);
        private void BtnBrowseCmpBaseline_Click(object sender, RoutedEventArgs e) => SelectFolder(TxtComparisonBaseline);
        private void BtnBrowseCmpOutput_Click(object sender, RoutedEventArgs e) => SelectFolder(TxtComparisonOutput);

        // Save log
        private void BtnSaveLog_Click(object sender, RoutedEventArgs e)
        {
            var saveDialog = new MicrosoftWin32.SaveFileDialog
            {
                FileName = $"cad_workflow_log_{DateTime.Now:yyyyMMdd_HHmmss}.txt",
                Filter = "Log Files (*.log)|*.log|Text Files (*.txt)|*.txt|All Files (*.*)|*.*",
                DefaultExt = ".log"
            };

            if (saveDialog.ShowDialog(this) == true)
            {
                try
                {
                    File.WriteAllText(saveDialog.FileName, LogBox.Text);
                    System.Windows.MessageBox.Show(this, "Log file saved successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show(this, $"Failed to save log file. Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        // New: Reference extraction flow (shows modal dialog then runs tool)
        private async void BtnRunRefExtraction_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg = new ReferenceExtractionDialog { Owner = this };
                if (dlg.ShowDialog() != true)
                {
                    LogBox.AppendText("[INFO] Reference extraction cancelled by user.\n");
                    return;
                }

                string inPath = dlg.InputPath;
                string outPath = dlg.OutputPath;

                if (string.IsNullOrEmpty(inPath) || string.IsNullOrEmpty(outPath))
                {
                    System.Windows.MessageBox.Show(this, "Please provide both the Evaluation Reference Model folder (input) and the Reference XML Files folder (output).", "Missing Information", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                BtnRunRefExtraction.IsEnabled = false;
                LogBox.AppendText($"[INFO] Starting Reference XML extraction. Input: {inPath}, Output: {outPath}\n");

                string args = $"--input \"{inPath}\" --output \"{outPath}\"";
                bool success = await RunTool("Tools\\xml_data_extraction.exe", args, LogBox);

                if (success)
                {
                    LogBox.AppendText("[INFO] Reference extraction completed successfully.\n");
                    TxtComparisonBaseline.Text = outPath;
                }
                else
                {
                    System.Windows.MessageBox.Show(this, "Reference XML extraction failed. Check the log for details.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                LogBox.AppendText($"\n[FATAL CRASH] An unexpected error occurred during Reference Extraction: {ex}\n");
                System.Windows.MessageBox.Show(this, $"A critical error occurred: {ex.Message}\n\nSee the log for full details.", "Workflow Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                BtnRunRefExtraction.IsEnabled = true;
            }
        }

        // Run external tool with streaming log
        private async Task<bool> RunTool(string exePath, string arguments, System.Windows.Controls.TextBox logBox)
        {
            string fullExePath = Path.Combine(AppContext.BaseDirectory, exePath);
            logBox.AppendText($"[INFO] Starting {exePath}...\n");
            logBox.AppendText($"[CMD] {fullExePath} {arguments}\n");
            logBox.ScrollToEnd();

            var startInfo = new ProcessStartInfo
            {
                FileName = fullExePath,
                Arguments = arguments,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            };

            var stopwatch = Stopwatch.StartNew();

            using (var process = new Process { StartInfo = startInfo, EnableRaisingEvents = true })
            {
                var outputCloseEvent = new TaskCompletionSource<bool>();
                var errorCloseEvent = new TaskCompletionSource<bool>();

                process.OutputDataReceived += (s, e) =>
                {
                    if (e.Data == null) outputCloseEvent.TrySetResult(true);
                    else Dispatcher.Invoke(() => { logBox.AppendText($"[LOG] {e.Data}\n"); logBox.ScrollToEnd(); });
                };
                process.ErrorDataReceived += (s, e) =>
                {
                    if (e.Data == null) errorCloseEvent.TrySetResult(true);
                    else Dispatcher.Invoke(() => { logBox.AppendText($"[ERROR] {e.Data}\n"); logBox.ScrollToEnd(); });
                };

                bool isStarted;
                try
                {
                    isStarted = process.Start();
                }
                catch (Exception ex)
                {
                    stopwatch.Stop();
                    logBox.AppendText($"[FATAL] Could not start process {exePath}. Error: {ex.Message}\n");
                    logBox.ScrollToEnd();
                    return false;
                }

                if (!isStarted)
                {
                    stopwatch.Stop();
                    logBox.AppendText($"[FATAL] Failed to start process {exePath} (Start returned false).\n");
                    logBox.ScrollToEnd();
                    return false;
                }

                process.BeginOutputReadLine();
                process.BeginErrorReadLine();

                await process.WaitForExitAsync();
                await Task.WhenAny(Task.WhenAll(outputCloseEvent.Task, errorCloseEvent.Task), Task.Delay(TimeSpan.FromSeconds(5)));

                stopwatch.Stop();
                var elapsed = stopwatch.Elapsed;

                if (process.ExitCode == 0)
                {
                    logBox.AppendText($"[INFO] {exePath} finished successfully in {elapsed.TotalSeconds:F2} seconds.\n");
                    logBox.ScrollToEnd();
                    return true;
                }
                else
                {
                    logBox.AppendText($"[FATAL] {exePath} failed after {elapsed.TotalSeconds:F2} seconds with exit code {process.ExitCode}.\n");
                    logBox.ScrollToEnd();
                    return false;
                }
            }
        }

        // Run external tool and collect its stdout lines
        private async Task<(bool Success, List<string> Output)> RunToolAndGetOutput(string exePath, string arguments, System.Windows.Controls.TextBox logBox)
        {
            string fullExePath = Path.Combine(AppContext.BaseDirectory, exePath);
            logBox.AppendText($"[INFO] Running {exePath} to get part list...\n");
            logBox.AppendText($"[CMD] {fullExePath} {arguments}\n");
            logBox.ScrollToEnd();

            var outputLines = new List<string>();
            var errorLines = new List<string>();

            var startInfo = new ProcessStartInfo
            {
                FileName = fullExePath,
                Arguments = arguments,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            };

            var stopwatch = Stopwatch.StartNew();

            using (var process = new Process { StartInfo = startInfo, EnableRaisingEvents = true })
            {
                var outputCloseEvent = new TaskCompletionSource<bool>();
                var errorCloseEvent = new TaskCompletionSource<bool>();

                process.OutputDataReceived += (s, e) =>
                {
                    if (e.Data == null) outputCloseEvent.TrySetResult(true);
                    else outputLines.Add(e.Data);
                };
                process.ErrorDataReceived += (s, e) =>
                {
                    if (e.Data == null) errorCloseEvent.TrySetResult(true);
                    else
                    {
                        errorLines.Add(e.Data);
                        Dispatcher.Invoke(() => { logBox.AppendText($"[ERROR] {e.Data}\n"); logBox.ScrollToEnd(); });
                    }
                };

                bool isStarted;
                try
                {
                    isStarted = process.Start();
                }
                catch (Exception ex)
                {
                    stopwatch.Stop();
                    logBox.AppendText($"[FATAL] Could not start process {exePath} to get output. Error: {ex.Message}\n");
                    logBox.ScrollToEnd();
                    return (false, null);
                }

                if (!isStarted)
                {
                    stopwatch.Stop();
                    logBox.AppendText($"[FATAL] Failed to start process {exePath} to get output (Start returned false).\n");
                    logBox.ScrollToEnd();
                    return (false, null);
                }

                process.BeginOutputReadLine();
                process.BeginErrorReadLine();

                await process.WaitForExitAsync();
                await Task.WhenAny(Task.WhenAll(outputCloseEvent.Task, errorCloseEvent.Task), Task.Delay(TimeSpan.FromSeconds(5)));

                stopwatch.Stop();
                var elapsed = stopwatch.Elapsed;

                if (process.ExitCode == 0)
                {
                    logBox.AppendText($"[INFO] Successfully retrieved part list in {elapsed.TotalSeconds:F2} seconds.\n");
                    logBox.ScrollToEnd();
                    return (true, outputLines.Where(line => !string.IsNullOrWhiteSpace(line)).ToList());
                }
                else
                {
                    logBox.AppendText($"[FATAL] Failed to get part list after {elapsed.TotalSeconds:F2} seconds. Exit code {process.ExitCode}.\n");
                    logBox.ScrollToEnd();
                    return (false, null);
                }
            }
        }

        // Preprocessor
        private async void BtnRunPreprocessor_Click(object sender, RoutedEventArgs e)
        {
            BtnRunPreprocessor.IsEnabled = false;
            BtnSkipPreprocessor.IsEnabled = false;
            try
            {
                string inPath = TxtPreprocessorInput.Text;
                string outPath = TxtPreprocessorOutput.Text;
                if (string.IsNullOrEmpty(inPath) || string.IsNullOrEmpty(outPath))
                {
                    System.Windows.MessageBox.Show(this, "Please provide both Input and Output paths to run the preprocessor.", "Missing Information", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                string args = $"--input \"{inPath}\" --output \"{outPath}\"";
                bool success = await RunTool("Tools\\pre_processing_tool.exe", args, LogBox);
                if (success)
                {
                    TxtExtractionInput.Text = outPath;
                    UpdateWizardState(2);
                }
                else
                {
                    System.Windows.MessageBox.Show(this, "Preprocessor failed. Check the log for details.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                LogBox.AppendText($"\n[FATAL CRASH] An unexpected error occurred during Preprocessing: {ex}\n");
                System.Windows.MessageBox.Show(this, $"A critical error occurred: {ex.Message}\n\nSee the log for full details.", "Workflow Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (_currentStep == 1)
                {
                    BtnRunPreprocessor.IsEnabled = true;
                    BtnSkipPreprocessor.IsEnabled = true;
                }
            }
        }

        private void BtnSkipPreprocessor_Click(object sender, RoutedEventArgs e)
        {
            string outPath = TxtPreprocessorOutput.Text;
            if (string.IsNullOrEmpty(outPath))
            {
                System.Windows.MessageBox.Show(this, "Please provide the 'Output Folder (Processed Files)' path.\nThis folder should contain the already arranged files needed for Step 2.", "Missing Information for Skip", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            TxtExtractionInput.Text = outPath;
            LogBox.AppendText("[INFO] Preprocessor step skipped. Using specified folder for Data Extraction.\n");
            UpdateWizardState(2);
        }

        // Extraction
        private async void BtnRunExtraction_Click(object sender, RoutedEventArgs e)
        {
            BtnRunExtraction.IsEnabled = false;
            BtnPreviousToStep1.IsEnabled = false;
            BtnSkipExtraction.IsEnabled = false;
            try
            {
                string inPath = TxtExtractionInput.Text;
                string outPath = TxtExtractionOutput.Text;
                if (string.IsNullOrEmpty(outPath))
                {
                    System.Windows.MessageBox.Show(this, "Please provide an Output path.", "Missing Information", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                string args = $"--input \"{inPath}\" --output \"{outPath}\"";
                bool success = await RunTool("Tools\\xml_data_extraction.exe", args, LogBox);
                if (success)
                {
                    TxtComparisonInput.Text = outPath;
                    UpdateWizardState(3);
                }
                else
                {
                    System.Windows.MessageBox.Show(this, "Data Extraction failed. Check the log for details.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                LogBox.AppendText($"\n[FATAL CRASH] An unexpected error occurred during Extraction: {ex}\n");
                System.Windows.MessageBox.Show(this, $"A critical error occurred: {ex.Message}\n\nSee the log for full details.", "Workflow Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (_currentStep == 2)
                {
                    BtnRunExtraction.IsEnabled = true;
                    BtnPreviousToStep1.IsEnabled = true;
                    BtnSkipExtraction.IsEnabled = true;
                }
            }
        }

        private void BtnSkipExtraction_Click(object sender, RoutedEventArgs e)
        {
            string outPath = TxtExtractionOutput.Text;
            if (string.IsNullOrEmpty(outPath))
            {
                System.Windows.MessageBox.Show(this, "Please provide the 'Output Folder (Extracted XMLs)' path.\nThis folder should contain the XML files needed for Step 3.", "Missing Information for Skip", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            TxtComparisonInput.Text = outPath;
            LogBox.AppendText("[INFO] Data Extraction step skipped. Using specified folder for XML Comparison.\n");
            UpdateWizardState(3);
        }

        // Comparison
        private async void BtnRunComparison_Click(object sender, RoutedEventArgs e)
        {
            BtnRunComparison.IsEnabled = false;
            BtnPreviousToStep2.IsEnabled = false;
            try
            {
                string inPath = TxtComparisonInput.Text;
                string baselinePath = TxtComparisonBaseline.Text;
                string outPath = TxtComparisonOutput.Text;

                    if (string.IsNullOrEmpty(baselinePath) || string.IsNullOrEmpty(outPath) || string.IsNullOrEmpty(inPath))
                {
                    System.Windows.MessageBox.Show(this, "Please provide all Input, Baseline, and Output paths.", "Missing Information", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                string listPartsArgs = $"--baseline \"{baselinePath}\" --list-parts";
                var (listSuccess, partNames) = await RunToolAndGetOutput("Tools\\xml_data_comparision.exe", listPartsArgs, LogBox);

                if (!listSuccess || partNames == null || partNames.Count == 0)
                {
                    System.Windows.MessageBox.Show(this, "Failed to get part list from baseline folder. Check log for errors. Are there .xml files in the baseline folder?", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                var dialog = new PointsEntryDialog(partNames) { Owner = this };
                if (dialog.ShowDialog() != true)
                {
                    LogBox.AppendText("[INFO] User cancelled points entry.\n");
                    return;
                }

                Dictionary<string, double> pointsMap = dialog.PointsMap;
                if (pointsMap == null)
                {
                    LogBox.AppendText("[ERROR] Failed to get points from dialog (PointsMap was null after submit).\n");
                    return;
                }

                string pointsMapJson = JsonSerializer.Serialize(pointsMap);
                string escapedPointsMapJson = $"\"{pointsMapJson.Replace("\"", "\\\"")}\"";

                LogBox.AppendText("[INFO] Starting full comparison with user-defined points...\n");

                string fullRunArgs = $"--input \"{inPath}\" " +
                                     $"--baseline \"{baselinePath}\" " +
                                     $"--output \"{outPath}\" " +
                                     $"--points-map {escapedPointsMapJson}";

                bool finalSuccess = await RunTool("Tools\\xml_data_comparision.exe", fullRunArgs, LogBox);

                if (finalSuccess)
                {
                    LogBox.AppendText("\n[SUCCESS] Workflow Complete! Comparison report generated.\n");
                    System.Windows.MessageBox.Show(this, "Workflow Complete!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    BtnRunComparison.Content = "Finished!";
                    BtnRunComparison.Background = System.Windows.Media.Brushes.Green;
                    BtnRunComparison.IsEnabled = false;
                    BtnPreviousToStep2.IsEnabled = false;
                }
                else
                {
                    System.Windows.MessageBox.Show(this, "XML Comparison failed. Check the log for details.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    BtnRunComparison.IsEnabled = true;
                    BtnPreviousToStep2.IsEnabled = true;
                }
            }
            catch (Exception ex)
            {
                LogBox.AppendText($"\n[FATAL CRASH] An unexpected error occurred during Comparison: {ex}\n");
                System.Windows.MessageBox.Show(this, $"A critical error occurred: {ex.Message}\n\nSee the log for full details.", "Workflow Crash", MessageBoxButton.OK, MessageBoxImage.Error);
                BtnRunComparison.IsEnabled = true;
                BtnPreviousToStep2.IsEnabled = true;
            }
            finally
            {
                if (_currentStep == 3 && BtnRunComparison.Content.ToString() != "Finished!")
                {
                    if (!BtnRunComparison.IsEnabled) BtnRunComparison.IsEnabled = true;
                    if (!BtnPreviousToStep2.IsEnabled) BtnPreviousToStep2.IsEnabled = true;
                }
            }
        }

        private void BtnPreviousToStep1_Click(object sender, RoutedEventArgs e) => UpdateWizardState(1);
        private void BtnPreviousToStep2_Click(object sender, RoutedEventArgs e) => UpdateWizardState(2);
    }
}