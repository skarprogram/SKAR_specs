// SPDX-License-Identifier: GPL-3.0-only

using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace SKAR_specs
{
    partial class Program
    {
        private static async Task<ReportCollectionResult> CollectReportDataAsync()
        {
            ReportCollectionResult collectionResult = new ReportCollectionResult();
            ReportData reportData = collectionResult.ReportData;

            try
            {
                completedSections = 0;
                const int totalSections = 9;

                Task<SectionResult> licensesTask = RunReportSectionAsync("LICENSES", async () => {
                    reportData.LicenseData = await CollectLicenseReportDataAsync();
                    officeLicenseName = reportData.LicenseData.OfficeLicenseName;
                }, totalSections);
                Task<SectionResult> summaryTask = RunReportSectionAsync("SUMMARY", async () => {
                    reportData.SummaryData = await CollectSummaryReportDataAsync();
                }, totalSections);
                Task<SectionResult> motherboardTask = RunReportSectionAsync("Motherboard", async () => {
                    reportData.MotherboardData = await CollectMotherboardReportDataAsync();
                }, totalSections);
                Task<SectionResult> memoryTask = RunReportSectionAsync("MEMORY", async () => {
                    reportData.MemoryData = await CollectMemoryReportDataAsync();
                }, totalSections);
                Task<SectionResult> hardDrivesTask = RunReportSectionAsync("HARD DRIVES", async () => {
                    reportData.DiskData = await CollectDiskReportDataAsync();
                }, totalSections);
                Task<SectionResult> displayTask = RunReportSectionAsync("DISPLAY", async () => {
                    reportData.DisplayData = await CollectDisplayReportDataAsync();
                }, totalSections);
                Task<SectionResult> networkTask = RunReportSectionAsync("NETWORK", async () => {
                    reportData.NetworkData = await CollectNetworkReportDataAsync();
                }, totalSections);
                Task<SectionResult> printersTask = RunReportSectionAsync("PRINTERS", async () => {
                    reportData.PrintersData = await CollectPrintersReportDataAsync();
                }, totalSections);
                Task<SectionResult> softwareTask = RunReportSectionAsync("SOFTWARE", async () => {
                    reportData.SoftwareData = await CollectSoftwareReportDataAsync();
                }, totalSections);

                await Task.WhenAll(licensesTask, summaryTask, motherboardTask, memoryTask, hardDrivesTask, displayTask, networkTask, printersTask, softwareTask);

                reportData.SummaryData.Office = GetOfficeSummaryText();
                StoreSectionError(reportData, licensesTask.Result);
                StoreSectionError(reportData, summaryTask.Result);
                StoreSectionError(reportData, motherboardTask.Result);
                StoreSectionError(reportData, memoryTask.Result);
                StoreSectionError(reportData, hardDrivesTask.Result);
                StoreSectionError(reportData, displayTask.Result);
                StoreSectionError(reportData, networkTask.Result);
                StoreSectionError(reportData, printersTask.Result);
                StoreSectionError(reportData, softwareTask.Result);
            }
            catch (Exception ex)
            {
                collectionResult.CollectionException = ex;
            }

            return collectionResult;
        }

        private static async Task RunTimedSectionAsync(string sectionName, Func<Task> sectionFunc, int totalSections)
        {
            DebugLog($"START {sectionName}");
            Stopwatch sectionStopwatch = Stopwatch.StartNew();
            try
            {
                await sectionFunc();
                int done = Interlocked.Increment(ref completedSections);
                DebugLog($"DONE  {sectionName} in {sectionStopwatch.Elapsed.TotalSeconds:F2}s [{done}/{totalSections}]");
            }
            catch (Exception ex)
            {
                int done = Interlocked.Increment(ref completedSections);
                DebugLog($"FAIL  {sectionName} after {sectionStopwatch.Elapsed.TotalSeconds:F2}s [{done}/{totalSections}]: {ex.Message}");
                throw;
            }
        }

        private static async Task<SectionResult> RunReportSectionAsync(string sectionName, Func<Task> sectionFunc, int totalSections)
        {
            try
            {
                await RunTimedSectionAsync(sectionName, sectionFunc, totalSections);
                return new SectionResult
                {
                    SectionName = sectionName,
                    Succeeded = true
                };
            }
            catch (Exception ex)
            {
                return new SectionResult
                {
                    SectionName = sectionName,
                    ErrorMessage = ex.Message,
                    Succeeded = false
                };
            }
        }

        private static void StoreSectionError(ReportData reportData, SectionResult sectionResult)
        {
            if (reportData == null || sectionResult == null || sectionResult.Succeeded || string.IsNullOrWhiteSpace(sectionResult.SectionName))
            {
                return;
            }

            reportData.SectionErrors[sectionResult.SectionName] = sectionResult.ErrorMessage;
        }

        private static ProcessRunResult RunProcessWithTimeout(ProcessStartInfo psi, int timeoutMilliseconds, string operationName)
        {
            ProcessRunResult result = new ProcessRunResult();

            if (psi == null)
            {
                result.Error = "Process start information is missing.";
                return result;
            }

            string validatedExecutablePath;
            if (!TryGetValidatedLocalExecutablePath(psi.FileName, out validatedExecutablePath))
            {
                result.Error = "Executable path must be an existing absolute local .exe path.";
                DebugLog($"BLOCK {operationName}: unsafe or missing executable path '{psi.FileName}'.");
                return result;
            }

            psi.FileName = validatedExecutablePath;
            if (string.IsNullOrWhiteSpace(psi.WorkingDirectory))
            {
                psi.WorkingDirectory = Path.GetDirectoryName(validatedExecutablePath);
            }

            psi.UseShellExecute = false;
            psi.RedirectStandardOutput = true;
            psi.RedirectStandardError = true;
            psi.CreateNoWindow = true;

            using (Process process = new Process { StartInfo = psi })
            {
                Stopwatch stopwatch = Stopwatch.StartNew();
                Task<string> outputTask = null;
                Task<string> errorTask = null;

                try
                {
                    if (!process.Start())
                    {
                        result.Error = "Process failed to start.";
                        return result;
                    }

                    outputTask = process.StandardOutput.ReadToEndAsync();
                    errorTask = process.StandardError.ReadToEndAsync();

                    bool exited = timeoutMilliseconds <= 0
                        ? WaitForProcessExit(process)
                        : process.WaitForExit(timeoutMilliseconds);

                    if (!exited)
                    {
                        result.TimedOut = true;
                        result.Error = $"{operationName} timed out after {timeoutMilliseconds}ms.";
                        try { process.Kill(); } catch { }
                        try { process.WaitForExit(1000); } catch { }
                        TryCopyCompletedProcessOutput(outputTask, errorTask, result);
                        DebugLog($"TIMEOUT {operationName} after {stopwatch.Elapsed.TotalSeconds:F2}s");
                    }
                    else
                    {
                        double processExitSeconds = stopwatch.Elapsed.TotalSeconds;
                        result.ExitCode = process.ExitCode;
                        bool streamsDrained = Task.WaitAll(new Task[] { outputTask, errorTask }, ProcessStreamDrainTimeoutMs);
                        TryCopyCompletedProcessOutput(outputTask, errorTask, result);
                        if (debugEnabled)
                        {
                            string drainStatus = streamsDrained ? "drained" : "not fully drained";
                            DebugLog($"DONE  {operationName} process exit in {processExitSeconds:F2}s, streams {drainStatus} in {stopwatch.Elapsed.TotalSeconds:F2}s");
                        }
                    }
                }
                catch (Exception ex)
                {
                    result.Error = ex.Message;
                    TryCopyCompletedProcessOutput(outputTask, errorTask, result);
                }
            }

            return result;
        }

        private static string GetSystemExecutablePath(string relativePath)
        {
            if (string.IsNullOrWhiteSpace(relativePath) || Path.IsPathRooted(relativePath))
            {
                throw new ArgumentException("A relative system executable path is required.", nameof(relativePath));
            }

            string systemDirectory = Path.GetFullPath(Environment.SystemDirectory)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            string executablePath = Path.GetFullPath(Path.Combine(systemDirectory, relativePath));
            string requiredPrefix = systemDirectory + Path.DirectorySeparatorChar;

            if (!executablePath.StartsWith(requiredPrefix, StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException("The executable path must remain inside the Windows system directory.", nameof(relativePath));
            }

            return executablePath;
        }

        private static bool TryGetValidatedLocalExecutablePath(string executablePath, out string validatedPath)
        {
            validatedPath = null;
            if (string.IsNullOrWhiteSpace(executablePath))
            {
                return false;
            }

            string trimmedPath = executablePath.Trim().Trim('"');
            if (!Path.IsPathRooted(trimmedPath) ||
                trimmedPath.StartsWith(@"\\", StringComparison.Ordinal) ||
                trimmedPath.StartsWith("//", StringComparison.Ordinal))
            {
                return false;
            }

            string pathRoot = Path.GetPathRoot(trimmedPath);
            if (string.IsNullOrWhiteSpace(pathRoot) || pathRoot.Length < 3 || pathRoot[1] != ':')
            {
                return false;
            }

            try
            {
                string fullPath = Path.GetFullPath(trimmedPath);
                if (!string.Equals(Path.GetExtension(fullPath), ".exe", StringComparison.OrdinalIgnoreCase) ||
                    !File.Exists(fullPath))
                {
                    return false;
                }

                validatedPath = fullPath;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void TryCopyCompletedProcessOutput(Task<string> outputTask, Task<string> errorTask, ProcessRunResult result)
        {
            try
            {
                if (outputTask != null && outputTask.IsCompleted && !outputTask.IsFaulted && !outputTask.IsCanceled)
                {
                    result.Output = outputTask.Result;
                }
            }
            catch
            {
                // Keep original process error.
            }

            try
            {
                if (errorTask != null && errorTask.IsCompleted && !errorTask.IsFaulted && !errorTask.IsCanceled)
                {
                    AppendProcessError(result, errorTask.Result);
                }
            }
            catch
            {
                // Keep original process error.
            }
        }

        private static void AppendProcessError(ProcessRunResult result, string processError)
        {
            if (!string.IsNullOrWhiteSpace(processError))
            {
                result.Error = string.IsNullOrWhiteSpace(result.Error)
                    ? processError
                    : result.Error + "\n" + processError;
            }
        }

        private static bool WaitForProcessExit(Process process)
        {
            process.WaitForExit();
            return true;
        }
    }
}
