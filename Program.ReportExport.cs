// SPDX-License-Identifier: GPL-3.0-only

using System;
using System.Linq;
using System.Text;

namespace SKAR_specs
{
    partial class Program
    {
        private class ReportExportRequest
        {
            public ReportCollectionResult CollectionResult { get; set; }
            public string GeneratedDate { get; set; } = "";
            public StringBuilder ErrorLog { get; set; }
        }

        private class ReportExporter
        {
            public ReportOutputFormat Format { get; set; }
            public string DisplayName { get; set; } = "";
            public string FileExtension { get; set; } = "";
            public Func<ReportExportRequest, string> Render { get; set; }
        }

        private static ReportExporter GetReportExporter(string[] args)
        {
            return GetReportExporter(GetReportOutputFormat(args));
        }

        private static ReportOutputFormat GetReportOutputFormat(string[] args)
        {
            if (args != null && args.Any(a =>
                string.Equals(a, "-json", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(a, "--json", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(a, "/json", StringComparison.OrdinalIgnoreCase)))
            {
                return ReportOutputFormat.Json;
            }

            if (args != null && args.Any(a =>
                string.Equals(a, "-txt", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(a, "--txt", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(a, "/txt", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(a, "-text", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(a, "--text", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(a, "/text", StringComparison.OrdinalIgnoreCase)))
            {
                return ReportOutputFormat.Text;
            }

            return ReportOutputFormat.Html;
        }

        private static ReportExporter GetReportExporter(ReportOutputFormat outputFormat)
        {
            switch (outputFormat)
            {
                case ReportOutputFormat.Json:
                    return new ReportExporter
                    {
                        Format = ReportOutputFormat.Json,
                        DisplayName = "JSON",
                        FileExtension = ".json",
                        Render = request => RenderJsonReport(request.CollectionResult, request.GeneratedDate, request.ErrorLog)
                    };
                case ReportOutputFormat.Text:
                    return new ReportExporter
                    {
                        Format = ReportOutputFormat.Text,
                        DisplayName = "Text",
                        FileExtension = ".txt",
                        Render = request => RenderTextReport(request.CollectionResult, request.GeneratedDate, request.ErrorLog)
                    };
                default:
                    return new ReportExporter
                    {
                        Format = ReportOutputFormat.Html,
                        DisplayName = "HTML",
                        FileExtension = ".html",
                        Render = request => RenderFullHtmlReport(
                            request.CollectionResult.ReportData,
                            request.GeneratedDate,
                            request.ErrorLog,
                            request.CollectionResult.CollectionException)
                    };
            }
        }

        private static string RenderReportOutput(
            ReportExporter exporter,
            ReportCollectionResult collectionResult,
            string generatedDate,
            StringBuilder errorLog)
        {
            ReportExportRequest request = new ReportExportRequest
            {
                CollectionResult = collectionResult,
                GeneratedDate = generatedDate,
                ErrorLog = errorLog
            };

            return exporter.Render(request);
        }
    }
}
