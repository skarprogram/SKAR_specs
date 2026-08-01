// SPDX-License-Identifier: GPL-3.0-only

using System;
using System.Net;
using System.Text;

namespace SKAR_specs
{
    partial class Program
    {
        private static string HtmlEncode(string value)
        {
            return WebUtility.HtmlEncode(value ?? string.Empty);
        }

        private static string HtmlText(object value)
        {
            return HtmlEncode(value?.ToString() ?? string.Empty);
        }

        private static string HtmlAttr(object value)
        {
            return HtmlEncode(value?.ToString() ?? string.Empty);
        }

        private static string TableRowText(string key, object value)
        {
            return string.Format(Templates.TableRow, HtmlText(key), HtmlText(value));
        }

        private static string TableRowHtml(string key, string htmlValue)
        {
            return string.Format(Templates.TableRow, HtmlText(key), htmlValue ?? string.Empty);
        }

        private static string TableRowHtmlWithId(string id, string key, string htmlValue)
        {
            string encodedKey = HtmlText(key);
            return $@"<tr id=""{HtmlAttr(id)}"">
<td class=""tg-0lax data-label"" data-key=""{encodedKey}"">{encodedKey}</td>
<td class=""tg-0lax data-value"" data-value=""{encodedKey}"">{htmlValue ?? string.Empty}</td>
</tr>";
        }

        private static string TableRowTextWithId(string id, string key, object value)
        {
            return TableRowHtmlWithId(id, key, HtmlText(value));
        }

        private static string RenderFullHtmlReport(
            ReportData reportData,
            string navHeader,
            StringBuilder errorLog,
            Exception collectionException)
        {
            StringBuilder html = new StringBuilder();

            html.Append(Templates.HtmlHeader);
            html.Append(string.Format(Templates.NavigationSection, navHeader));

            if (collectionException == null)
            {
                AppendReportSection(html, "SUMMARY", "SUMMARY", RenderReportSectionHtml(reportData, "SUMMARY", () =>
                    RenderSummaryInfoHtml(reportData.SummaryData)
                        .Replace(HardDriveSummaryToken, RenderHardDriveSummaryHtml(reportData.DiskData))
                        .Replace(ExternalIpSummaryToken, RenderExternalIpSummaryHtml(reportData.NetworkData))));
                AppendReportSection(html, "Motherboard", "Motherboard", RenderReportSectionHtml(reportData, "Motherboard", () => RenderMotherboardInfoHtml(reportData.MotherboardData)));
                AppendReportSection(html, "MEMORY", "MEMORY", RenderReportSectionHtml(reportData, "MEMORY", () => RenderMemoryInfoHtml(reportData.MemoryData)));
                AppendReportSection(html, "HARDDRIVES", "HARD DRIVES", RenderReportSectionHtml(reportData, "HARD DRIVES", () => RenderHardDrivesInfoHtml(reportData.DiskData)));
                AppendReportSection(html, "DISPLAY", "DISPLAY", RenderReportSectionHtml(reportData, "DISPLAY", () => RenderDisplayInfoHtml(reportData.DisplayData)));
                AppendReportSection(html, "Network", "NETWORK", RenderReportSectionHtml(reportData, "NETWORK", () => RenderNetworkInfoHtml(reportData.NetworkData)));
                AppendReportSection(html, "PRINTERS", "PRINTERS", RenderReportSectionHtml(reportData, "PRINTERS", () => RenderPrintersInfoHtml(reportData.PrintersData)));
                AppendReportSection(html, "SOFTWARE", "SOFTWARE", RenderReportSectionHtml(reportData, "SOFTWARE", () => RenderSoftwareInfoHtml(reportData.SoftwareData)));
                AppendReportSection(html, "LICENSES", "LICENSES", RenderReportSectionHtml(reportData, "LICENSES", () => RenderLicensesInfoHtml(reportData.LicenseData)));
            }
            else
            {
                AppendReportSection(html, "ERROR", "Error During Information Collection",
                    TableRowHtml("Error Details", $"{HtmlText(collectionException.Message)}<br><pre>{HtmlText(collectionException.StackTrace)}</pre>"));
            }

            AppendReportSection(html, "END", "End of the report.", string.Empty);

            if (errorLog != null && errorLog.Length > 0)
            {
                AppendReportSection(html, "GENERATIONLOG", "Report Generation Log",
                    TableRowHtml("Generation Log",
                        $"<div style='color: #D32F2F;'>{HtmlText(errorLog.ToString()).Replace(Environment.NewLine, "<br>")}</div>"));
            }

            if (debugEnabled && DebugLogBuilder.Length > 0)
            {
                string debugOutput = HtmlEncode(DebugLogBuilder.ToString()).Replace(Environment.NewLine, "<br>");
                AppendReportSection(html, "DEBUGLOG", "Debug Timing Log",
                    TableRowHtml("Debug Log", $"<div style='font-family: Consolas, monospace; font-size: 12px;'>{debugOutput}</div>"));
            }

            html.Append(Templates.HtmlFooter);
            return html.ToString();
        }

        private static string RenderReportSectionHtml(ReportData reportData, string sectionName, Func<string> renderFunc)
        {
            string sectionError;
            if (reportData != null &&
                reportData.SectionErrors.TryGetValue(sectionName, out sectionError) &&
                !string.IsNullOrWhiteSpace(sectionError))
            {
                return TableRowText(sectionName, $"Error collecting section: {sectionError}");
            }

            return renderFunc();
        }

        private static void AppendReportSection(StringBuilder html, string id, string title, string bodyHtml)
        {
            html.Append(string.Format(Templates.SectionStart, id, title));
            html.Append(bodyHtml ?? string.Empty);
        }
    }
}
