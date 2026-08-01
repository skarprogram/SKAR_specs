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
        private static async Task<LicenseReportData> CollectLicenseReportDataAsync()
        {
            return await Task.Run(() => {
                LicenseReportData reportData = new LicenseReportData();

                try
                {
                    string[] paths = new string[] {
                        @"C:\Program Files\Microsoft Office\root\Office16\OSPP.VBS",
                        @"C:\Program Files (x86)\Microsoft Office\root\Office16\OSPP.VBS",
                        @"C:\Program Files\Microsoft Office\Office16\OSPP.VBS",
                        @"C:\Program Files (x86)\Microsoft Office\Office16\OSPP.VBS",
                        @"C:\Program Files\Microsoft Office\Office15\OSPP.VBS",
                        @"C:\Program Files (x86)\Microsoft Office\Office15\OSPP.VBS",
                        @"C:\Program Files\Microsoft Office\Office14\OSPP.VBS",
                        @"C:\Program Files (x86)\Microsoft Office\Office14\OSPP.VBS"
                    };

                    foreach (string path in paths)
                    {
                        if (!File.Exists(path))
                        {
                            continue;
                        }

                        reportData.OfficeInstallationFound = true;
                        reportData.OfficeInstallationPath = path;

                        try
                        {
                            ProcessStartInfo psi = new ProcessStartInfo
                            {
                                FileName = GetSystemExecutablePath("cscript.exe"),
                                Arguments = $"//Nologo \"{path}\" /dstatus",
                                WorkingDirectory = Path.GetDirectoryName(path)
                            };

                            ProcessRunResult processResult = RunProcessWithTimeout(psi, ExternalCommandTimeoutMs, "Office license check");
                            string licenseOutput = processResult.TimedOut ? processResult.Error : processResult.Output;
                            string errorOutput = processResult.Error;

                            if (!string.IsNullOrWhiteSpace(errorOutput) && !licenseOutput.Contains(errorOutput))
                            {
                                licenseOutput += "\n" + errorOutput;
                            }

                            reportData.OfficeLicenseOutput = licenseOutput;
                            reportData.OfficeLicenseName = ExtractOfficeLicenseName(licenseOutput);
                        }
                        catch (Exception ex)
                        {
                            reportData.OfficeLicenseOutput = $"Error running OSPP.VBS: {ex.Message}";
                        }

                        break;
                    }
                }
                catch (Exception ex)
                {
                    reportData.ErrorMessage = ex.Message;
                }

                return reportData;
            });
        }

        private static string ExtractOfficeLicenseName(string licenseOutput)
        {
            if (string.IsNullOrWhiteSpace(licenseOutput))
            {
                return "";
            }

            string[] lines = licenseOutput.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string line in lines)
            {
                if (line.Contains("LICENSE NAME:"))
                {
                    int startIndex = line.IndexOf("LICENSE NAME:") + "LICENSE NAME:".Length;
                    return line.Substring(startIndex).Trim();
                }
            }

            return "";
        }

        private static string RenderLicensesInfoHtml(LicenseReportData licenseData)
        {
            if (licenseData == null)
            {
                return TableRowText("License Information", "Error retrieving license information");
            }

            if (!string.IsNullOrWhiteSpace(licenseData.ErrorMessage))
            {
                return TableRowText("License Information", $"Error retrieving license information: {licenseData.ErrorMessage}");
            }

            StringBuilder licenseInfo = new StringBuilder();
            licenseInfo.Append("<h3>Microsoft Office Licenses</h3>");
            licenseInfo.Append("<table border='1' class='licenses-table' data-table-type='office-licenses'>");
            licenseInfo.Append("<thead><tr><th>Product</th><th>License Status</th></tr></thead><tbody>");

            if (licenseData.OfficeInstallationFound && !string.IsNullOrWhiteSpace(licenseData.OfficeLicenseOutput))
            {
                licenseInfo.Append($"<tr><td colspan='2'><strong>Found Office installation at:</strong> {HtmlText(licenseData.OfficeInstallationPath)}</td></tr>");

                string htmlOutput = licenseData.OfficeLicenseOutput
                    .Replace("&", "&amp;")
                    .Replace("<", "&lt;")
                    .Replace(">", "&gt;")
                    .Replace("\n", "<br>")
                    .Replace("\r", "");

                licenseInfo.Append($"<tr><td colspan='2'><pre style='margin: 0; font-family: Consolas, monospace; font-size: 12px; white-space: pre-wrap;'>{htmlOutput}</pre></td></tr>");
            }
            else if (!licenseData.OfficeInstallationFound)
            {
                licenseInfo.Append("<tr><td colspan='2'>No supported Microsoft Office installation found.</td></tr>");
            }
            else
            {
                licenseInfo.Append("<tr><td colspan='2'>Office found but no license information available.</td></tr>");
            }

            licenseInfo.Append("</tbody></table>");
            return TableRowHtml("License Information", licenseInfo.ToString());
        }
    }
}
