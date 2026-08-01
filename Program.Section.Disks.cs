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
        private static object GetManagementValue(ManagementBaseObject managementObject, string propertyName)
        {
            try
            {
                return managementObject?[propertyName];
            }
            catch
            {
                return null;
            }
        }

        private static string GetManagementString(ManagementBaseObject managementObject, string propertyName)
        {
            object value = GetManagementValue(managementObject, propertyName);
            return value?.ToString() ?? string.Empty;
        }

        private static string NormalizeStorageBusType(string busType)
        {
            if (string.IsNullOrWhiteSpace(busType))
            {
                return "Unknown";
            }

            switch (busType.Trim())
            {
                case "0": return "Unknown";
                case "1": return "SCSI";
                case "2": return "ATAPI";
                case "3": return "ATA";
                case "4": return "IEEE 1394";
                case "5": return "SSA";
                case "6": return "Fibre Channel";
                case "7": return "USB";
                case "8": return "RAID";
                case "9": return "iSCSI";
                case "10": return "SAS";
                case "11": return "SATA";
                case "12": return "SD";
                case "13": return "MMC";
                case "14": return "Virtual";
                case "15": return "File Backed Virtual";
                case "16": return "Storage Spaces";
                case "17": return "NVMe";
                default: return busType.Trim();
            }
        }

        private static DiskLookupInfo GetDiskLookupInfo(bool includePhysicalDisk = true)
        {
            DiskLookupInfo lookupInfo = new DiskLookupInfo();

            if (includePhysicalDisk)
            {
                // Get bus type and serial number information from PowerShell Get-PhysicalDisk.
                // This provider can be slow on some systems, so use it only for detailed disk inventory.
                try
                {
                    ProcessStartInfo psi = new ProcessStartInfo
                    {
                        FileName = GetSystemExecutablePath(@"WindowsPowerShell\v1.0\powershell.exe"),
                        Arguments = "-NoProfile -NonInteractive -Command \"Get-PhysicalDisk | Select-Object DeviceID, BusType, SerialNumber | ForEach-Object { Write-Output \\\"$($_.DeviceID)|$($_.BusType)|$($_.SerialNumber)\\\" }\"",
                    };

                    ProcessRunResult processResult = RunProcessWithTimeout(psi, 0, "Get-PhysicalDisk");
                    string output = processResult.Output;
                    if (!processResult.TimedOut && !string.IsNullOrWhiteSpace(output))
                    {
                        string[] lines = output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (string line in lines)
                        {
                            string[] parts = line.Split(new[] { '|' }, 3);
                            if (parts.Length >= 2)
                            {
                                string deviceId = parts[0].Trim();
                                string busType = NormalizeStorageBusType(parts[1]);
                                string serialNumber = parts.Length == 3 ? parts[2].Trim() : string.Empty;

                                lookupInfo.BusTypeMap[deviceId] = busType;
                                if (!string.IsNullOrWhiteSpace(serialNumber))
                                {
                                    lookupInfo.PhysicalDiskSerialMap[deviceId] = serialNumber;
                                }
                            }
                        }
                    }
                }
                catch
                {
                    // If PowerShell fails, continue without Get-PhysicalDisk information
                }
            }

            try
            {
                using (ManagementClass managementClass = new ManagementClass("Win32_PhysicalMedia"))
                using (ManagementObjectCollection collection = managementClass.GetInstances())
                {
                    foreach (ManagementObject media in collection)
                    {
                        string tag = media["Tag"]?.ToString();
                        string serialNumber = media["SerialNumber"]?.ToString();

                        if (!string.IsNullOrWhiteSpace(tag) && !string.IsNullOrWhiteSpace(serialNumber))
                        {
                            lookupInfo.PhysicalMediaSerialMap[tag.Trim()] = serialNumber.Trim();
                        }
                    }
                }
            }
            catch
            {
                // If WMI physical media lookup fails, continue without fallback serial numbers
            }

            return lookupInfo;
        }

        private static string GetDiskBusType(ManagementObject disk, DiskLookupInfo lookupInfo)
        {
            string diskIndex = GetManagementString(disk, "Index");
            if (!string.IsNullOrWhiteSpace(diskIndex))
            {
                if (lookupInfo.BusTypeMap.ContainsKey(diskIndex))
                {
                    return lookupInfo.BusTypeMap[diskIndex];
                }
            }

            string pnpDeviceId = GetManagementString(disk, "PNPDeviceID");
            string model = GetManagementString(disk, "Model");
            string pnpAndModel = pnpDeviceId + " " + model;

            if (pnpAndModel.IndexOf("NVME", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "NVMe";
            }

            if (pnpDeviceId.StartsWith("USBSTOR", StringComparison.OrdinalIgnoreCase))
            {
                return "USB";
            }

            return "Unknown";
        }

        private static string GetDiskSerialNumber(ManagementObject disk, DiskLookupInfo lookupInfo)
        {
            string serialNumber = "Unknown";
            string diskIndex = GetManagementString(disk, "Index");
            if (!string.IsNullOrWhiteSpace(diskIndex))
            {
                if (lookupInfo.PhysicalDiskSerialMap.ContainsKey(diskIndex) &&
                    !string.IsNullOrWhiteSpace(lookupInfo.PhysicalDiskSerialMap[diskIndex]))
                {
                    serialNumber = lookupInfo.PhysicalDiskSerialMap[diskIndex];
                }
            }

            string diskSerialNumber = GetManagementString(disk, "SerialNumber");
            if (!string.IsNullOrWhiteSpace(diskSerialNumber))
            {
                serialNumber = diskSerialNumber.Trim();
            }
            else
            {
                string deviceId = GetManagementString(disk, "DeviceID");
                if (serialNumber == "Unknown" &&
                    lookupInfo.PhysicalMediaSerialMap.ContainsKey(deviceId) &&
                    !string.IsNullOrWhiteSpace(lookupInfo.PhysicalMediaSerialMap[deviceId]))
                {
                    serialNumber = lookupInfo.PhysicalMediaSerialMap[deviceId];
                }
            }

            return serialNumber;
        }

        private static string FormatCapacityFromBytes(object sizeValue)
        {
            if (sizeValue == null)
            {
                return "Unknown";
            }

            double sizeGB = Convert.ToDouble(sizeValue) / (1024 * 1024 * 1024);
            if (sizeGB >= 1024)
            {
                return $"{sizeGB / 1024:F2} TB";
            }

            return $"{sizeGB:F2} GB";
        }

        private static async Task<DiskReportData> CollectDiskReportDataAsync()
        {
            return await Task.Run(() => {
                DiskReportData reportData = new DiskReportData();
                DiskLookupInfo lookupInfo = GetDiskLookupInfo();

                try
                {
                    using (ManagementClass managementClass = new ManagementClass("Win32_DiskDrive"))
                    using (ManagementObjectCollection collection = managementClass.GetInstances())
                    {
                        int diskNumber = 0;
                        foreach (ManagementObject disk in collection)
                        {
                            int displayDiskNumber = diskNumber;
                            string diskIndexValue = GetManagementString(disk, "Index");
                            int parsedDiskIndex;
                            if (int.TryParse(diskIndexValue, out parsedDiskIndex))
                            {
                                displayDiskNumber = parsedDiskIndex;
                            }

                            DiskDriveInfo diskInfo = new DiskDriveInfo
                            {
                                Number = displayDiskNumber,
                                Model = ValueOrUnknown(GetManagementString(disk, "Model")),
                                Capacity = FormatCapacityFromBytes(GetManagementValue(disk, "Size")),
                                Status = ValueOrUnknown(GetManagementString(disk, "Status")),
                                MediaType = ValueOrUnknown(GetManagementString(disk, "MediaType")),
                                DeviceId = ValueOrUnknown(GetManagementString(disk, "DeviceID")),
                                BusType = GetDiskBusType(disk, lookupInfo),
                                SerialNumber = GetDiskSerialNumber(disk, lookupInfo)
                            };

                            object sizeValue = GetManagementValue(disk, "Size");
                            if (sizeValue != null)
                            {
                                try
                                {
                                    diskInfo.SizeGB = Convert.ToDouble(sizeValue) / (1024 * 1024 * 1024);
                                }
                                catch
                                {
                                    diskInfo.SizeGB = 0;
                                }
                            }

                            diskInfo.Partitions.AddRange(CollectDiskPartitionInfo(disk));
                            reportData.Disks.Add(diskInfo);
                            diskNumber++;
                        }
                    }
                }
                catch (Exception ex)
                {
                    reportData.ErrorMessage = ex.Message;
                    DebugLog($"HARD DRIVES detail: disk collection failed: {ex.Message}");
                }

                return reportData;
            });
        }

        private static List<DiskPartitionInfo> CollectDiskPartitionInfo(ManagementObject disk)
        {
            var partitionInfo = new List<DiskPartitionInfo>();
            if (disk == null)
            {
                return partitionInfo;
            }

            try
            {
                using (ManagementObjectCollection partitions = disk.GetRelated("Win32_DiskPartition"))
                {
                    foreach (ManagementObject partition in partitions)
                    {
                        string partitionDeviceId = GetManagementString(partition, "DeviceID");
                        string partitionName = ValueOrUnknown(GetManagementString(partition, "Name"));
                        double partitionSizeGB = BytesToGB(GetManagementValue(partition, "Size"));
                        bool logicalDiskFound = false;

                        try
                        {
                            using (ManagementObjectCollection logicalDisks = partition.GetRelated("Win32_LogicalDisk"))
                            {
                                foreach (ManagementObject logicalDisk in logicalDisks)
                                {
                                    double freeSpaceGB = BytesToGB(GetManagementValue(logicalDisk, "FreeSpace"));
                                    partitionInfo.Add(new DiskPartitionInfo
                                    {
                                        Name = partitionName,
                                        DeviceId = ValueOrUnknown(partitionDeviceId),
                                        DriveLetter = ValueOrDefault(GetManagementString(logicalDisk, "DeviceID"), "N/A"),
                                        SizeGB = partitionSizeGB,
                                        FreeSpaceGB = freeSpaceGB,
                                        UsedSpaceGB = Math.Max(0, partitionSizeGB - freeSpaceGB)
                                    });
                                    logicalDiskFound = true;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            partitionInfo.Add(new DiskPartitionInfo
                            {
                                Name = partitionName,
                                DeviceId = ValueOrUnknown(partitionDeviceId),
                                SizeGB = partitionSizeGB,
                                ErrorMessage = "Error retrieving logical disk information: " + ex.Message
                            });
                            logicalDiskFound = true;
                        }

                        if (!logicalDiskFound)
                        {
                            partitionInfo.Add(new DiskPartitionInfo
                            {
                                Name = partitionName,
                                DeviceId = ValueOrUnknown(partitionDeviceId),
                                DriveLetter = "N/A",
                                SizeGB = partitionSizeGB,
                                FreeSpaceGB = 0,
                                UsedSpaceGB = partitionSizeGB
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                partitionInfo.Add(new DiskPartitionInfo
                {
                    ErrorMessage = "Error retrieving disk partitions: " + ex.Message
                });
            }

            return partitionInfo;
        }

        private static string RenderHardDriveSummaryHtml(DiskReportData diskData)
        {
            StringBuilder summary = new StringBuilder();

            summary.Append("<table border='1' class='disk-summary-table' data-table-type='disk-summary' style='width:100%; font-size:12px; border-collapse:collapse;'>");
            summary.Append("<thead><tr><th>#</th><th>Model</th><th>Capacity</th><th>Serial Number</th><th>Type</th></tr></thead><tbody>");

            if (diskData == null || !string.IsNullOrWhiteSpace(diskData.ErrorMessage))
            {
                summary.Append("<tr><td colspan='5'>Error retrieving hard drive summary</td></tr>");
            }
            else if (diskData.Disks.Count == 0)
            {
                summary.Append("<tr><td colspan='5'>No disks found</td></tr>");
            }
            else
            {
                foreach (DiskDriveInfo disk in diskData.Disks)
                {
                    summary.Append($"<tr data-disk-summary-index=\"{disk.Number}\">");
                    summary.Append($"<td data-field=\"index\">{disk.Number}</td>");
                    summary.Append($"<td data-field=\"model\">{HtmlText(disk.Model)}</td>");
                    summary.Append($"<td data-field=\"capacity\">{HtmlText(disk.Capacity)}</td>");
                    summary.Append($"<td data-field=\"serial-number\">{HtmlText(disk.SerialNumber)}</td>");
                    summary.Append($"<td data-field=\"bus-type\">{HtmlText(disk.BusType)}</td>");
                    summary.Append("</tr>");
                }
            }

            summary.Append("</tbody></table>");
            return summary.ToString();
        }

        private static string RenderHardDrivesInfoHtml(DiskReportData diskData)
        {
            StringBuilder hardDrivesInfo = new StringBuilder();

            if (diskData == null)
            {
                return TableRowText("Hard drives Get-Disk", "Error retrieving hard drive information");
            }

            if (!string.IsNullOrWhiteSpace(diskData.ErrorMessage))
            {
                hardDrivesInfo.Append("<p>Error retrieving disk drives</p>");
                return TableRowHtml("Hard drives Get-Disk", hardDrivesInfo.ToString());
            }

            if (diskData.Disks.Count == 0)
            {
                hardDrivesInfo.Append("<p>No disks found</p>");
                return TableRowHtml("Hard drives Get-Disk", hardDrivesInfo.ToString());
            }

            foreach (DiskDriveInfo disk in diskData.Disks)
            {
                if (!string.IsNullOrWhiteSpace(disk.ErrorMessage))
                {
                    hardDrivesInfo.Append($"<p>Error retrieving information for Disk {disk.Number}: {HtmlText(disk.ErrorMessage)}</p>");
                    continue;
                }

                hardDrivesInfo.Append($"<h3>Disk {disk.Number}: {HtmlText(disk.Model)}</h3>");
                hardDrivesInfo.Append("<table border='1' class='disk-table' data-table-type='disk-drive' data-disk-id=\"" + HtmlAttr(disk.DeviceId) + "\">");
                hardDrivesInfo.Append("<thead><tr><th>Property</th><th>Value</th></tr></thead><tbody>");
                hardDrivesInfo.Append($"<tr><td data-field=\"model\">Model</td><td>{HtmlText(disk.Model)}</td></tr>");
                hardDrivesInfo.Append($"<tr><td data-field=\"size\">Size</td><td>{disk.SizeGB:F2} GB</td></tr>");
                hardDrivesInfo.Append($"<tr><td data-field=\"status\">Status</td><td>{HtmlText(disk.Status)}</td></tr>");
                hardDrivesInfo.Append($"<tr><td data-field=\"media-type\">Media Type</td><td>{HtmlText(disk.MediaType)}</td></tr>");
                hardDrivesInfo.Append($"<tr><td data-field=\"bus-type\">Bus Type</td><td>{HtmlText(disk.BusType)}</td></tr>");
                hardDrivesInfo.Append($"<tr><td data-field=\"serial-number\">Serial Number</td><td>{HtmlText(disk.SerialNumber)}</td></tr>");
                hardDrivesInfo.Append("</tbody></table>");

                hardDrivesInfo.Append("<h4>Partitions</h4>");
                hardDrivesInfo.Append("<table border='1' width='100%' class='partition-table' data-disk-id=\"" + HtmlAttr(disk.DeviceId) + "\">");
                hardDrivesInfo.Append("<thead><tr><th>Partition</th><th>Drive Letter</th><th>Size (GB)</th><th>Free Space (GB)</th><th>Used Space (GB)</th></tr></thead><tbody>");

                if (disk.Partitions.Count == 0)
                {
                    hardDrivesInfo.Append("<tr><td colspan='5'>No partitions found</td></tr>");
                }
                else
                {
                    foreach (DiskPartitionInfo partition in disk.Partitions)
                    {
                        if (!string.IsNullOrWhiteSpace(partition.ErrorMessage))
                        {
                            hardDrivesInfo.Append("<tr><td colspan='5'>" + HtmlText(partition.ErrorMessage) + "</td></tr>");
                            continue;
                        }

                        hardDrivesInfo.Append($"<tr data-partition-id=\"{HtmlAttr(partition.DeviceId)}\" data-drive-letter=\"{HtmlAttr(partition.DriveLetter)}\">");
                        hardDrivesInfo.Append($"<td data-field=\"name\">{HtmlText(partition.Name)}</td>");
                        hardDrivesInfo.Append($"<td data-field=\"drive-id\">{HtmlText(partition.DriveLetter)}</td>");
                        hardDrivesInfo.Append($"<td data-field=\"size\">{partition.SizeGB:F2}</td>");
                        hardDrivesInfo.Append($"<td data-field=\"free-space\">{partition.FreeSpaceGB:F2}</td>");
                        hardDrivesInfo.Append($"<td data-field=\"used-space\">{partition.UsedSpaceGB:F2}</td>");
                        hardDrivesInfo.Append("</tr>");
                    }
                }

                hardDrivesInfo.Append("</tbody></table><br>");
            }

            return TableRowHtml("Hard drives Get-Disk", hardDrivesInfo.ToString());
        }

        private static double BytesToGB(object sizeValue)
        {
            if (sizeValue == null)
            {
                return 0;
            }

            try
            {
                return Convert.ToDouble(sizeValue) / (1024 * 1024 * 1024);
            }
            catch
            {
                return 0;
            }
        }

        private static string ValueOrUnknown(string value)
        {
            return ValueOrDefault(value, "Unknown");
        }

        private static string ValueOrDefault(string value, string defaultValue)
        {
            return string.IsNullOrWhiteSpace(value) ? defaultValue : value.Trim();
        }
    }
}
