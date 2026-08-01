// SPDX-License-Identifier: GPL-3.0-only

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SKAR_specs
{
    partial class Program
    {
        private static string RenderTextReport(
            ReportCollectionResult collectionResult,
            string generatedDate,
            StringBuilder errorLog)
        {
            StringBuilder text = new StringBuilder();
            ReportData reportData = collectionResult.ReportData;

            text.AppendLine("SKAR_specs report");
            text.AppendLine("Generated: " + generatedDate);
            text.AppendLine(new string('=', 72));
            text.AppendLine();

            if (collectionResult.CollectionException != null)
            {
                AppendTextSection(text, "ERROR");
                AppendTextValue(text, "Error during information collection", collectionResult.CollectionException.Message);
                AppendTextValue(text, "Stack trace", collectionResult.CollectionException.StackTrace);
                AppendOptionalTextLog(text, errorLog, "Report Generation Log");
                AppendOptionalDebugTextLog(text);
                return text.ToString();
            }

            AppendTextSection(text, "SUMMARY");
            if (!TryAppendSectionError(text, reportData, "SUMMARY"))
            {
                AppendSummaryText(text, reportData);
            }

            AppendTextSection(text, "MOTHERBOARD");
            if (!TryAppendSectionError(text, reportData, "Motherboard"))
            {
                AppendMotherboardText(text, reportData.MotherboardData);
            }

            AppendTextSection(text, "MEMORY");
            if (!TryAppendSectionError(text, reportData, "MEMORY"))
            {
                AppendMemoryText(text, reportData.MemoryData);
            }

            AppendTextSection(text, "HARD DRIVES");
            if (!TryAppendSectionError(text, reportData, "HARD DRIVES"))
            {
                AppendDisksText(text, reportData.DiskData);
            }

            AppendTextSection(text, "DISPLAY");
            if (!TryAppendSectionError(text, reportData, "DISPLAY"))
            {
                AppendDisplayText(text, reportData.DisplayData);
            }

            AppendTextSection(text, "NETWORK");
            if (!TryAppendSectionError(text, reportData, "NETWORK"))
            {
                AppendNetworkText(text, reportData.NetworkData);
            }

            AppendTextSection(text, "PRINTERS");
            if (!TryAppendSectionError(text, reportData, "PRINTERS"))
            {
                AppendPrintersText(text, reportData.PrintersData);
            }

            AppendTextSection(text, "SOFTWARE");
            if (!TryAppendSectionError(text, reportData, "SOFTWARE"))
            {
                AppendSoftwareText(text, reportData.SoftwareData);
            }

            AppendTextSection(text, "LICENSES");
            if (!TryAppendSectionError(text, reportData, "LICENSES"))
            {
                AppendLicensesText(text, reportData.LicenseData);
            }

            AppendOptionalTextLog(text, errorLog, "REPORT GENERATION LOG");
            AppendOptionalDebugTextLog(text);

            return text.ToString();
        }

        private static bool TryAppendSectionError(StringBuilder text, ReportData reportData, string sectionName)
        {
            string sectionError;
            if (reportData.SectionErrors.TryGetValue(sectionName, out sectionError) && !string.IsNullOrWhiteSpace(sectionError))
            {
                AppendTextValue(text, "Error collecting section", sectionError);
                return true;
            }

            return false;
        }

        private static void AppendSummaryText(StringBuilder text, ReportData reportData)
        {
            SummaryReportData summary = reportData.SummaryData;
            AppendTextValue(text, "Computer name", summary.ComputerName);
            AppendTextValue(text, "Manufacturer", summary.Manufacturer);
            AppendTextValue(text, "Model", summary.Model);
            AppendTextValue(text, "Serial number", summary.SerialNumber);
            AppendTextValue(text, "Processor", summary.Processor);
            if (summary.TotalPhysicalProcessors > 1)
            {
                AppendTextValue(text, "Total physical processors", summary.TotalPhysicalProcessors);
            }
            AppendTextValue(text, "RAM", summary.Ram);
            AppendTextValue(text, "Hard drives", GetHardDriveSummaryText(reportData.DiskData));
            AppendTextValue(text, "GPU", summary.Gpu);
            AppendTextValue(text, "Current user", summary.CurrentUser);
            AppendTextValue(text, "Fully qualified domain name", summary.FullyQualifiedDomainName);
            AppendTextValue(text, "TPM", summary.Tpm);
            if (!string.IsNullOrWhiteSpace(summary.Workgroup))
            {
                AppendTextValue(text, "Workgroup", summary.Workgroup);
            }
            AppendTextValue(text, "Operating system", summary.OperatingSystem);
            AppendTextValue(text, "Office", summary.Office);
            AppendTextValue(text, "Remote Software", RenderRemoteSoftwareSummaryText(summary.RemoteSoftware));
            AppendTextValue(text, "External IP", GetExternalIpSummaryText(reportData.NetworkData));

            AppendTextSubsection(text, "Summary Network Adapters");
            if (!string.IsNullOrWhiteSpace(summary.NetworkAdaptersErrorMessage))
            {
                AppendTextValue(text, "Error", summary.NetworkAdaptersErrorMessage);
            }
            else if (summary.NetworkAdapters.Count == 0)
            {
                text.AppendLine("  No network adapters found");
            }
            else
            {
                foreach (SummaryNetworkAdapterInfo adapter in summary.NetworkAdapters)
                {
                    AppendTextBullet(text, adapter.Name);
                    AppendTextValue(text, "MAC", adapter.MacAddress, 4);
                    AppendTextValue(text, "IPv4", JoinOrDefault(adapter.Ipv4Addresses, ", ", "N/A"), 4);
                    AppendTextValue(text, "IP assignment", adapter.IpAssignment, 4);
                    AppendTextValue(text, "Status", adapter.Status, 4);
                    AppendTextValue(text, "Type", adapter.ConnectionType, 4);
                }
            }
        }

        private static void AppendMotherboardText(StringBuilder text, MotherboardReportData motherboard)
        {
            AppendTextValue(text, "Manufacturer", motherboard.Manufacturer);
            AppendTextValue(text, "Model", motherboard.Model);
            AppendTextValue(text, "BIOS version", motherboard.BiosVersion);
            AppendTextValue(text, "Baseboard error", motherboard.BaseBoardErrorMessage);
            AppendTextValue(text, "BIOS error", motherboard.BiosErrorMessage);
            AppendTextValue(text, "Detail error", motherboard.DetailErrorMessage);

            if (motherboard.Detail != null)
            {
                AppendTextSubsection(text, "Details");
                AppendTextValue(text, "Manufacturer", motherboard.Detail.Manufacturer);
                AppendTextValue(text, "Model", motherboard.Detail.Model);
                AppendTextValue(text, "Product", motherboard.Detail.Product);
                AppendTextValue(text, "Serial number", motherboard.Detail.SerialNumber);
                AppendTextValue(text, "Version", motherboard.Detail.Version);
            }
        }

        private static void AppendMemoryText(StringBuilder text, MemoryReportData memory)
        {
            AppendTextValue(text, "Total system memory", memory.TotalSystemMemory);
            AppendTextValue(text, "Memory array error", memory.MemoryArrayErrorMessage);
            AppendTextValue(text, "Physical memory error", memory.PhysicalMemoryErrorMessage);
            AppendTextValue(text, "Error", memory.ErrorMessage);

            AppendTextSubsection(text, "Memory Arrays");
            if (memory.MemoryArrays.Count == 0)
            {
                text.AppendLine("  No memory array data found");
            }
            foreach (MemoryArrayInfo array in memory.MemoryArrays)
            {
                AppendTextBullet(text, $"Max capacity: {array.MaxCapacityGB:F2} GB; Devices: {array.MemoryDevices}");
                AppendTextValue(text, "Error", array.ErrorMessage, 4);
            }

            AppendTextSubsection(text, "Physical Memory Modules");
            if (memory.PhysicalMemoryModules.Count == 0)
            {
                text.AppendLine("  No physical memory modules found");
            }
            foreach (PhysicalMemoryModuleInfo module in memory.PhysicalMemoryModules)
            {
                AppendTextBullet(text, $"{module.DeviceLocator} / {module.BankLabel}");
                AppendTextValue(text, "Manufacturer", module.Manufacturer, 4);
                AppendTextValue(text, "Capacity", $"{module.CapacityGB:F2} GB", 4);
                AppendTextValue(text, "Clock speed", module.ClockSpeed, 4);
                AppendTextValue(text, "Serial number", module.SerialNumber, 4);
                AppendTextValue(text, "Error", module.ErrorMessage, 4);
            }
        }

        private static void AppendDisksText(StringBuilder text, DiskReportData diskData)
        {
            AppendTextValue(text, "Error", diskData.ErrorMessage);
            if (diskData.Disks.Count == 0)
            {
                text.AppendLine("  No disks found");
                return;
            }

            foreach (DiskDriveInfo disk in diskData.Disks)
            {
                AppendTextBullet(text, $"Disk {disk.Number}: {disk.Model}");
                AppendTextValue(text, "Capacity", disk.Capacity, 4);
                AppendTextValue(text, "Status", disk.Status, 4);
                AppendTextValue(text, "Media type", disk.MediaType, 4);
                AppendTextValue(text, "Bus type", disk.BusType, 4);
                AppendTextValue(text, "Device ID", disk.DeviceId, 4);
                AppendTextValue(text, "Serial number", disk.SerialNumber, 4);
                AppendTextValue(text, "Error", disk.ErrorMessage, 4);

                foreach (DiskPartitionInfo partition in disk.Partitions)
                {
                    AppendTextBullet(text, $"{partition.Name} ({partition.DriveLetter})", 4);
                    AppendTextValue(text, "Size", $"{partition.SizeGB:F2} GB", 6);
                    AppendTextValue(text, "Free", $"{partition.FreeSpaceGB:F2} GB", 6);
                    AppendTextValue(text, "Used", $"{partition.UsedSpaceGB:F2} GB", 6);
                    AppendTextValue(text, "Device ID", partition.DeviceId, 6);
                    AppendTextValue(text, "Error", partition.ErrorMessage, 6);
                }
            }
        }

        private static void AppendDisplayText(StringBuilder text, DisplayReportData display)
        {
            AppendTextValue(text, "Video controller error", display.VideoControllersErrorMessage);
            AppendTextValue(text, "Monitor error", display.MonitorsErrorMessage);
            AppendTextValue(text, "Error", display.ErrorMessage);

            AppendTextSubsection(text, "Video Controllers");
            if (display.VideoControllers.Count == 0)
            {
                text.AppendLine("  No video controllers found");
            }
            foreach (VideoControllerInfo controller in display.VideoControllers)
            {
                AppendTextBullet(text, controller.Name);
                AppendTextValue(text, "Device ID", controller.DeviceId, 4);
                AppendTextValue(text, "Processor", controller.Processor, 4);
                AppendTextValue(text, "Resolution", controller.Resolution, 4);
                AppendTextValue(text, "Memory", controller.Memory, 4);
                AppendTextValue(text, "Driver version", controller.DriverVersion, 4);
                AppendTextValue(text, "Error", controller.ErrorMessage, 4);
            }

            AppendTextSubsection(text, "Monitors");
            if (display.Monitors.Count == 0)
            {
                text.AppendLine("  No monitors found");
            }
            foreach (MonitorInfo monitor in display.Monitors)
            {
                AppendTextBullet(text, string.IsNullOrWhiteSpace(monitor.FriendlyName) ? monitor.Model : monitor.FriendlyName);
                AppendTextValue(text, "Vendor", monitor.Vendor, 4);
                AppendTextValue(text, "Model", monitor.Model, 4);
                AppendTextValue(text, "Serial", monitor.Serial, 4);
                AppendTextValue(text, "Serial source", monitor.SerialSource, 4);
                AppendTextValue(text, "Error", monitor.ErrorMessage, 4);
            }
        }

        private static void AppendNetworkText(StringBuilder text, NetworkReportData network)
        {
            AppendTextValue(text, "Error", network.ErrorMessage);

            AppendTextSubsection(text, "Network Adapters");
            if (!string.IsNullOrWhiteSpace(network.AdaptersErrorMessage))
            {
                AppendTextValue(text, "Error", network.AdaptersErrorMessage);
            }
            else if (network.Adapters.Count == 0)
            {
                text.AppendLine("  No network adapters found");
            }
            else
            {
                foreach (NetworkAdapterReportItem adapter in network.Adapters)
                {
                    AppendTextBullet(text, adapter.Name);
                    AppendTextValue(text, "MAC", adapter.MacAddress, 4);
                    AppendTextValue(text, "Status", adapter.Status, 4);
                    AppendTextValue(text, "Speed", adapter.Speed, 4);
                    AppendTextValue(text, "Type", adapter.ConnectionType, 4);
                    AppendTextValue(text, "IP addresses", JoinOrDefault(adapter.IpAddresses, ", ", "N/A"), 4);
                    AppendTextValue(text, "IP assignment", adapter.IpAssignment, 4);
                }
            }

            AppendTextSubsection(text, "External Connectivity");
            if (network.ExternalIpResults.Count == 0)
            {
                text.AppendLine("  External IP: Unable to retrieve");
            }
            foreach (ExternalIpServiceResult result in network.ExternalIpResults)
            {
                AppendTextBullet(text, result.ServiceName);
                AppendTextValue(text, "IP address", result.Success ? result.IpAddress : result.Message, 4);
                AppendTextValue(text, "Status", result.Status, 4);
            }
            AppendTextValue(text, "Internet connectivity", network.InternetConnectivityCheckFailed ? "Error checking" : (network.InternetConnected ? "Connected" : "Disconnected"));
            AppendTextValue(text, "Connectivity status", network.InternetConnectivityStatus);

            AppendTextSubsection(text, "Network Configuration");
            if (!string.IsNullOrWhiteSpace(network.IpConfigText))
            {
                text.AppendLine(network.IpConfigText.TrimEnd());
            }
            else
            {
                AppendTextValue(text, "Error", network.IpConfigErrorMessage);
            }

            AppendNetworkSharesText(text, network);
            AppendMappedDrivesText(text, network);
        }

        private static void AppendNetworkSharesText(StringBuilder text, NetworkReportData network)
        {
            AppendTextSubsection(text, "Network Shares");
            if (!string.IsNullOrWhiteSpace(network.SharesErrorMessage))
            {
                AppendTextValue(text, "Error", network.SharesErrorMessage);
            }
            else if (network.Shares.Count == 0)
            {
                text.AppendLine("  No network shares found");
            }
            else
            {
                foreach (NetworkShareInfo share in network.Shares)
                {
                    AppendTextBullet(text, share.Name);
                    AppendTextValue(text, "Path", share.Path, 4);
                    AppendTextValue(text, "Description", share.Description, 4);
                    AppendTextValue(text, "Error", share.ErrorMessage, 4);
                }
            }
        }

        private static void AppendMappedDrivesText(StringBuilder text, NetworkReportData network)
        {
            AppendTextSubsection(text, "Mapped Network Drives");
            if (!string.IsNullOrWhiteSpace(network.MappedDrivesErrorMessage))
            {
                AppendTextValue(text, "Error", network.MappedDrivesErrorMessage);
            }
            else if (network.MappedDrives.Count == 0)
            {
                text.AppendLine("  No mapped network drives found");
            }
            else
            {
                foreach (MappedNetworkDriveInfo drive in network.MappedDrives)
                {
                    AppendTextBullet(text, drive.DriveLetter);
                    AppendTextValue(text, "Remote path", drive.RemotePath, 4);
                    AppendTextValue(text, "Status", drive.Status, 4);
                    AppendTextValue(text, "Error", drive.ErrorMessage, 4);
                }
            }
        }

        private static void AppendPrintersText(StringBuilder text, PrintersReportData printers)
        {
            AppendTextValue(text, "Enumeration error", printers.EnumerationErrorMessage);
            AppendTextValue(text, "Error", printers.ErrorMessage);

            if (printers.Printers.Count == 0)
            {
                text.AppendLine("  No printers found");
                return;
            }

            foreach (PrinterReportItem printer in printers.Printers)
            {
                AppendTextBullet(text, printer.Name);
                AppendTextValue(text, "Status", printer.Status, 4);
                AppendTextValue(text, "Port", printer.PortName, 4);
                AppendTextValue(text, "IP address", printer.IpAddress, 4);
                AppendTextValue(text, "Shared", printer.Shared, 4);
                AppendTextValue(text, "Error", printer.ErrorMessage, 4);
            }
        }

        private static void AppendSoftwareText(StringBuilder text, SoftwareReportData software)
        {
            AppendTextValue(text, "Error", software.ErrorMessage);

            AppendTextSubsection(text, "Traditional Software");
            if (software.TraditionalSoftware.Count == 0)
            {
                text.AppendLine("  No traditional software found");
            }
            foreach (string item in software.TraditionalSoftware)
            {
                AppendTextBullet(text, item);
            }

            AppendTextSubsection(text, "UWP Apps");
            if (software.UwpApps.Count == 0)
            {
                text.AppendLine("  No UWP apps found");
            }
            foreach (UwpAppInfo app in software.UwpApps)
            {
                AppendTextBullet(text, app.Name);
                AppendTextValue(text, "Package ID", app.PackageId, 4);
            }
        }

        private static void AppendLicensesText(StringBuilder text, LicenseReportData license)
        {
            AppendTextValue(text, "Office installation found", license.OfficeInstallationFound ? "Yes" : "No");
            AppendTextValue(text, "Office installation path", license.OfficeInstallationPath);
            AppendTextValue(text, "Office license name", license.OfficeLicenseName);
            AppendTextValue(text, "Error", license.ErrorMessage);

            if (!string.IsNullOrWhiteSpace(license.OfficeLicenseOutput))
            {
                AppendTextSubsection(text, "Office License Output");
                text.AppendLine(license.OfficeLicenseOutput.TrimEnd());
            }
        }

        private static string GetHardDriveSummaryText(DiskReportData diskData)
        {
            if (diskData == null || diskData.Disks.Count == 0)
            {
                return "No disks found";
            }

            return string.Join("; ", diskData.Disks.Select(disk => $"{disk.Model} ({disk.Capacity})"));
        }

        private static string GetExternalIpSummaryText(NetworkReportData networkData)
        {
            if (networkData == null)
            {
                return "Unable to retrieve";
            }

            ExternalIpServiceResult result = networkData.ExternalIpResults
                .FirstOrDefault(item => item.Success && !string.IsNullOrWhiteSpace(item.IpAddress));

            return result == null ? "Unable to retrieve" : result.IpAddress;
        }

        private static void AppendOptionalTextLog(StringBuilder text, StringBuilder log, string title)
        {
            if (log == null || log.Length == 0)
            {
                return;
            }

            AppendTextSection(text, title);
            text.AppendLine(log.ToString().TrimEnd());
            text.AppendLine();
        }

        private static void AppendOptionalDebugTextLog(StringBuilder text)
        {
            if (!debugEnabled || DebugLogBuilder.Length == 0)
            {
                return;
            }

            AppendTextSection(text, "DEBUG TIMING LOG");
            text.AppendLine(DebugLogBuilder.ToString().TrimEnd());
            text.AppendLine();
        }

        private static void AppendTextSection(StringBuilder text, string title)
        {
            text.AppendLine();
            text.AppendLine(title);
            text.AppendLine(new string('-', title.Length));
        }

        private static void AppendTextSubsection(StringBuilder text, string title)
        {
            text.AppendLine();
            text.AppendLine(title + ":");
        }

        private static void AppendTextValue(StringBuilder text, string key, object value, int indent = 2)
        {
            string stringValue = value?.ToString() ?? "";
            if (string.IsNullOrWhiteSpace(stringValue))
            {
                return;
            }

            text.Append(' ', indent);
            text.Append(key);
            text.Append(": ");
            text.AppendLine(stringValue);
        }

        private static void AppendTextBullet(StringBuilder text, string value, int indent = 2)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return;
            }

            text.Append(' ', indent);
            text.Append("- ");
            text.AppendLine(value);
        }

        private static string JoinOrDefault(IEnumerable<string> values, string separator, string defaultValue)
        {
            if (values == null)
            {
                return defaultValue;
            }

            List<string> nonEmptyValues = values.Where(value => !string.IsNullOrWhiteSpace(value)).ToList();
            return nonEmptyValues.Count == 0 ? defaultValue : string.Join(separator, nonEmptyValues);
        }
    }
}
