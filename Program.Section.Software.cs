// SPDX-License-Identifier: GPL-3.0-only

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Net;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Win32;
namespace SKAR_specs
{
    partial class Program
    {
        private static async Task<SoftwareReportData> CollectSoftwareReportDataAsync()
        {
            return await Task.Run(() => {
                SoftwareReportData reportData = new SoftwareReportData();
                List<string> traditionalSoftware = new List<string>();
                List<string> uwpApps = new List<string>();

                try
                {
                    var tasks = new List<Task>();

                    tasks.Add(Task.Run(() => {
                        CollectTraditionalSoftwareFromRegistry(
                            Registry.LocalMachine,
                            @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                            traditionalSoftware);
                    }));

                    tasks.Add(Task.Run(() => {
                        if (Environment.Is64BitOperatingSystem)
                        {
                            CollectTraditionalSoftwareFromRegistry(
                                Registry.LocalMachine,
                                @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
                                traditionalSoftware);
                        }
                    }));

                    tasks.Add(Task.Run(() => {
                        CollectTraditionalSoftwareFromRegistry(
                            Registry.CurrentUser,
                            @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                            traditionalSoftware);
                    }));

                    tasks.Add(Task.Run(() => {
                        CollectUwpAppLines(uwpApps);
                    }));

                    Task.WaitAll(tasks.ToArray());

                    foreach (string softwareName in traditionalSoftware.Distinct().OrderBy(s => s))
                    {
                        reportData.TraditionalSoftware.Add(softwareName);
                    }

                    foreach (string appInfo in uwpApps.Distinct().OrderBy(s => s))
                    {
                        reportData.UwpApps.Add(ParseUwpAppInfo(appInfo));
                    }
                }
                catch (Exception ex)
                {
                    reportData.ErrorMessage = ex.Message;
                }

                return reportData;
            });
        }

        private static void CollectTraditionalSoftwareFromRegistry(RegistryKey rootKey, string uninstallPath, List<string> traditionalSoftware)
        {
            try
            {
                using (RegistryKey uninstallKey = rootKey.OpenSubKey(uninstallPath))
                {
                    if (uninstallKey == null)
                    {
                        return;
                    }

                    foreach (string subKeyName in uninstallKey.GetSubKeyNames())
                    {
                        try
                        {
                            using (RegistryKey subKey = uninstallKey.OpenSubKey(subKeyName))
                            {
                                string displayName = subKey?.GetValue("DisplayName") as string;
                                if (!string.IsNullOrEmpty(displayName))
                                {
                                    lock (traditionalSoftware)
                                    {
                                        traditionalSoftware.Add(displayName);
                                    }
                                }
                            }
                        }
                        catch
                        {
                            // Skip problematic entries.
                        }
                    }
                }
            }
            catch
            {
                // If we can't access this registry section, continue.
            }
        }

        private static void CollectUwpAppLines(List<string> uwpApps)
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = GetSystemExecutablePath(@"WindowsPowerShell\v1.0\powershell.exe"),
                    Arguments = "-Command \"Get-AppxPackage | Select-Object Name, PackageFullName | Format-Table -AutoSize\"",
                };

                ProcessRunResult processResult = RunProcessWithTimeout(psi, ExternalCommandTimeoutMs, "Get-AppxPackage");
                string output = processResult.Output;
                if (processResult.TimedOut || string.IsNullOrWhiteSpace(output))
                {
                    return;
                }

                string[] lines = output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                bool headerPassed = false;

                foreach (string line in lines)
                {
                    if (line.Contains("----"))
                    {
                        headerPassed = true;
                        continue;
                    }

                    if (headerPassed && !string.IsNullOrWhiteSpace(line))
                    {
                        string appInfo = line.Trim();
                        if (!string.IsNullOrEmpty(appInfo))
                        {
                            lock (uwpApps)
                            {
                                uwpApps.Add(appInfo);
                            }
                        }
                    }
                }
            }
            catch
            {
                // UWP apps info is optional, so we can continue without it.
            }
        }

        private static UwpAppInfo ParseUwpAppInfo(string appInfo)
        {
            string[] parts = (appInfo ?? string.Empty).Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            return new UwpAppInfo
            {
                Name = parts.Length > 0 ? parts[0] : "Unknown",
                PackageId = parts.Length > 1 ? string.Join(" ", parts.Skip(1)) : "Unknown"
            };
        }

        private static string RenderSoftwareInfoHtml(SoftwareReportData softwareData)
        {
            if (softwareData == null)
            {
                return TableRowText("Installed Software", "Error retrieving software information");
            }

            if (!string.IsNullOrWhiteSpace(softwareData.ErrorMessage))
            {
                return TableRowText("Installed Software", $"Error retrieving software information: {softwareData.ErrorMessage}");
            }

            StringBuilder softwareInfo = new StringBuilder();
            softwareInfo.Append("<h3>Traditional Software</h3>");
            softwareInfo.Append("<table border='1' class='software-table' data-table-type='traditional-software'>");
            softwareInfo.Append("<thead><tr><th>#</th><th>Software Name</th></tr></thead><tbody>");

            if (softwareData.TraditionalSoftware.Count > 0)
            {
                for (int i = 0; i < softwareData.TraditionalSoftware.Count; i++)
                {
                    softwareInfo.Append($"<tr data-software-index=\"{i + 1}\">");
                    softwareInfo.Append($"<td data-field=\"index\">{i + 1}</td>");
                    softwareInfo.Append($"<td data-field=\"name\">{HtmlText(softwareData.TraditionalSoftware[i])}</td>");
                    softwareInfo.Append("</tr>");
                }
            }
            else
            {
                softwareInfo.Append("<tr><td colspan='2'>No traditional software information available</td></tr>");
            }

            softwareInfo.Append("</tbody></table>");
            softwareInfo.Append($"<p class='software-count' data-software-type='traditional' data-count='{softwareData.TraditionalSoftware.Count}'>Total traditional software packages: <span>{softwareData.TraditionalSoftware.Count}</span></p>");

            softwareInfo.Append("<h3>Universal Windows Platform (UWP) Apps</h3>");
            softwareInfo.Append("<table border='1' class='software-table' data-table-type='uwp-apps'>");
            softwareInfo.Append("<thead><tr><th>#</th><th>App Name</th><th>Package ID</th></tr></thead><tbody>");

            if (softwareData.UwpApps.Count > 0)
            {
                for (int i = 0; i < softwareData.UwpApps.Count; i++)
                {
                    UwpAppInfo app = softwareData.UwpApps[i];
                    softwareInfo.Append($"<tr data-app-index=\"{i + 1}\" data-app-name=\"{HtmlAttr(app.Name)}\">");
                    softwareInfo.Append($"<td data-field=\"index\">{i + 1}</td>");
                    softwareInfo.Append($"<td data-field=\"app-name\">{HtmlText(app.Name)}</td>");
                    softwareInfo.Append($"<td data-field=\"package-id\">{HtmlText(app.PackageId)}</td>");
                    softwareInfo.Append("</tr>");
                }
            }
            else
            {
                softwareInfo.Append("<tr><td colspan='3'>No UWP apps found</td></tr>");
            }

            softwareInfo.Append("</tbody></table>");
            softwareInfo.Append($"<p class='software-count' data-software-type='uwp' data-count='{softwareData.UwpApps.Count}'>Total UWP apps: <span>{softwareData.UwpApps.Count}</span></p>");
            softwareInfo.Append($"<p class='software-total' data-total-count='{softwareData.TraditionalSoftware.Count + softwareData.UwpApps.Count}'><strong>Total installed software: <span>{softwareData.TraditionalSoftware.Count + softwareData.UwpApps.Count}</span></strong></p>");

            return TableRowHtml("Installed Software", softwareInfo.ToString());
        }
    }
}
