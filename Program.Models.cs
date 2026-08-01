// SPDX-License-Identifier: GPL-3.0-only

using System;
using System.Collections.Generic;

namespace SKAR_specs
{
    partial class Program
    {
        private enum ReportOutputFormat
        {
            Html,
            Json,
            Text
        }

        private class ProcessRunResult
        {
            public string Output { get; set; } = "";
            public string Error { get; set; } = "";
            public int ExitCode { get; set; } = -1;
            public bool TimedOut { get; set; }
            public string CombinedOutput
            {
                get
                {
                    if (string.IsNullOrWhiteSpace(Error))
                    {
                        return Output ?? "";
                    }

                    if (string.IsNullOrWhiteSpace(Output))
                    {
                        return Error ?? "";
                    }

                    return Output + "\n" + Error;
                }
            }
        }

        private class SectionResult
        {
            public string SectionName { get; set; } = "";
            public string ErrorMessage { get; set; } = "";
            public bool Succeeded { get; set; }
        }

        private class ReportCollectionResult
        {
            public ReportData ReportData { get; set; } = new ReportData();
            public Exception CollectionException { get; set; }
        }

        private class PrinterBasicInfo
        {
            public string Name { get; set; } = "Unknown";
            public string Status { get; set; } = "Unknown";
            public string PortName { get; set; } = "Unknown";
            public string Shared { get; set; } = "No";
        }

        private class PrintersReportData
        {
            public List<PrinterReportItem> Printers { get; } = new List<PrinterReportItem>();
            public string EnumerationErrorMessage { get; set; } = "";
            public string ErrorMessage { get; set; } = "";
        }

        private class PrinterReportItem
        {
            public string Name { get; set; } = "Unknown";
            public string Status { get; set; } = "Unknown";
            public string PortName { get; set; } = "Unknown";
            public string IpAddress { get; set; } = "Unknown";
            public string Shared { get; set; } = "No";
            public string ErrorMessage { get; set; } = "";
        }

        private class PrinterLookupInfo
        {
            public Dictionary<string, string> TcpIpPortHosts { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, string> PortDescriptions { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        private class SoftwareReportData
        {
            public List<string> TraditionalSoftware { get; } = new List<string>();
            public List<UwpAppInfo> UwpApps { get; } = new List<UwpAppInfo>();
            public string ErrorMessage { get; set; } = "";
        }

        private class UwpAppInfo
        {
            public string Name { get; set; } = "Unknown";
            public string PackageId { get; set; } = "Unknown";
        }

        private class LicenseReportData
        {
            public bool OfficeInstallationFound { get; set; }
            public string OfficeInstallationPath { get; set; } = "";
            public string OfficeLicenseOutput { get; set; } = "";
            public string OfficeLicenseName { get; set; } = "";
            public string ErrorMessage { get; set; } = "";
        }

        private class MemoryReportData
        {
            public string TotalSystemMemory { get; set; } = "Unknown";
            public List<MemoryArrayInfo> MemoryArrays { get; } = new List<MemoryArrayInfo>();
            public List<PhysicalMemoryModuleInfo> PhysicalMemoryModules { get; } = new List<PhysicalMemoryModuleInfo>();
            public string MemoryArrayErrorMessage { get; set; } = "";
            public string PhysicalMemoryErrorMessage { get; set; } = "";
            public string ErrorMessage { get; set; } = "";
        }

        private class MemoryArrayInfo
        {
            public double MaxCapacityGB { get; set; }
            public string MemoryDevices { get; set; } = "Unknown";
            public string ErrorMessage { get; set; } = "";
        }

        private class PhysicalMemoryModuleInfo
        {
            public string Manufacturer { get; set; } = "Unknown";
            public string BankLabel { get; set; } = "Unknown";
            public string ClockSpeed { get; set; } = "Unknown";
            public string DeviceLocator { get; set; } = "Unknown";
            public double CapacityGB { get; set; }
            public string SerialNumber { get; set; } = "Unknown";
            public string ErrorMessage { get; set; } = "";
        }

        private class MotherboardReportData
        {
            public string Manufacturer { get; set; } = "Not detected";
            public string Model { get; set; } = "Not detected";
            public string BiosVersion { get; set; } = "Not detected";
            public MotherboardDetailInfo Detail { get; set; }
            public string BaseBoardErrorMessage { get; set; } = "";
            public string BiosErrorMessage { get; set; } = "";
            public string DetailErrorMessage { get; set; } = "";
        }

        private class MotherboardDetailInfo
        {
            public string Manufacturer { get; set; } = "Not available";
            public string Model { get; set; } = "Not available";
            public string Product { get; set; } = "Not available";
            public string SerialNumber { get; set; } = "Not available";
            public string Version { get; set; } = "Not available";
        }

        private class DisplayReportData
        {
            public List<VideoControllerInfo> VideoControllers { get; } = new List<VideoControllerInfo>();
            public List<MonitorInfo> Monitors { get; } = new List<MonitorInfo>();
            public List<string> MonitorDebugMessages { get; } = new List<string>();
            public string VideoControllersErrorMessage { get; set; } = "";
            public string MonitorsErrorMessage { get; set; } = "";
            public string ErrorMessage { get; set; } = "";
        }

        private class VideoControllerInfo
        {
            public string DeviceId { get; set; } = "Unknown";
            public string Name { get; set; } = "Unknown";
            public string Processor { get; set; } = "Unknown";
            public string Resolution { get; set; } = "N/A";
            public string Memory { get; set; } = "Unknown";
            public string DriverVersion { get; set; } = "Unknown";
            public string ErrorMessage { get; set; } = "";
        }

        private class MonitorInfo
        {
            public string Vendor { get; set; } = "Unknown";
            public string Model { get; set; } = "Unknown Model";
            public string Serial { get; set; } = "";
            public string SerialSource { get; set; } = "";
            public string FriendlyName { get; set; } = "";
            public string ErrorMessage { get; set; } = "";
        }

        private class NetworkReportData
        {
            public string IpConfigText { get; set; } = "";
            public string IpConfigErrorMessage { get; set; } = "";
            public List<NetworkAdapterReportItem> Adapters { get; } = new List<NetworkAdapterReportItem>();
            public string AdaptersErrorMessage { get; set; } = "";
            public List<ExternalIpServiceResult> ExternalIpResults { get; } = new List<ExternalIpServiceResult>();
            public bool InternetConnectivityChecked { get; set; }
            public bool InternetConnected { get; set; }
            public string InternetConnectivityStatus { get; set; } = "Failed";
            public bool InternetConnectivityCheckFailed { get; set; }
            public List<NetworkShareInfo> Shares { get; } = new List<NetworkShareInfo>();
            public string SharesErrorMessage { get; set; } = "";
            public List<MappedNetworkDriveInfo> MappedDrives { get; } = new List<MappedNetworkDriveInfo>();
            public string MappedDrivesErrorMessage { get; set; } = "";
            public string ErrorMessage { get; set; } = "";
        }

        private class NetworkAdapterReportItem
        {
            public string Name { get; set; } = "";
            public string MacAddress { get; set; } = "";
            public string Status { get; set; } = "";
            public string Speed { get; set; } = "Unknown";
            public string ConnectionType { get; set; } = "";
            public List<string> IpAddresses { get; } = new List<string>();
            public string IpAssignment { get; set; } = "";
        }

        private class NetworkShareInfo
        {
            public string Name { get; set; } = "Unknown";
            public string Path { get; set; } = "Unknown";
            public string Description { get; set; } = "Unknown";
            public string ErrorMessage { get; set; } = "";
        }

        private class MappedNetworkDriveInfo
        {
            public string DriveLetter { get; set; } = "Unknown";
            public string RemotePath { get; set; } = "Unknown";
            public string Status { get; set; } = "Unknown";
            public string ErrorMessage { get; set; } = "";
        }

        private class SummaryReportData
        {
            public string ComputerName { get; set; } = "";
            public string Manufacturer { get; set; } = "";
            public string Model { get; set; } = "";
            public string SerialNumber { get; set; } = "";
            public string Processor { get; set; } = "";
            public int TotalPhysicalProcessors { get; set; }
            public string Ram { get; set; } = "";
            public string Gpu { get; set; } = "";
            public List<SummaryNetworkAdapterInfo> NetworkAdapters { get; } = new List<SummaryNetworkAdapterInfo>();
            public string NetworkAdaptersErrorMessage { get; set; } = "";
            public string CurrentUser { get; set; } = "";
            public string FullyQualifiedDomainName { get; set; } = "";
            public string Tpm { get; set; } = "";
            public string Workgroup { get; set; } = "";
            public string OperatingSystem { get; set; } = "";
            public string Office { get; set; } = "";
            public RemoteSoftwareSummaryInfo RemoteSoftware { get; set; } = new RemoteSoftwareSummaryInfo();
            public string ErrorMessage { get; set; } = "";
        }

        private class RemoteSoftwareSummaryInfo
        {
            public string AnyDeskId { get; set; } = "";
            public string TeamViewerId { get; set; } = "";
        }

        private class SummaryNetworkAdapterInfo
        {
            public string Name { get; set; } = "Unknown";
            public string MacAddress { get; set; } = "N/A";
            public List<string> Ipv4Addresses { get; } = new List<string>();
            public string IpAssignment { get; set; } = "Unknown";
            public string Status { get; set; } = "Unknown";
            public string ConnectionType { get; set; } = "Unknown";
        }

        private class ExternalIpServiceResult
        {
            public string ServiceName { get; set; } = "";
            public string IpAddress { get; set; } = "";
            public string Status { get; set; } = "Failed";
            public string Message { get; set; } = "";
            public bool Success { get; set; }
        }

        private class ReportData
        {
            public Dictionary<string, string> SectionErrors { get; } = new Dictionary<string, string>();
            public DiskReportData DiskData { get; set; } = new DiskReportData();
            public PrintersReportData PrintersData { get; set; } = new PrintersReportData();
            public SoftwareReportData SoftwareData { get; set; } = new SoftwareReportData();
            public LicenseReportData LicenseData { get; set; } = new LicenseReportData();
            public MemoryReportData MemoryData { get; set; } = new MemoryReportData();
            public MotherboardReportData MotherboardData { get; set; } = new MotherboardReportData();
            public DisplayReportData DisplayData { get; set; } = new DisplayReportData();
            public NetworkReportData NetworkData { get; set; } = new NetworkReportData();
            public SummaryReportData SummaryData { get; set; } = new SummaryReportData();
        }

        private class DiskReportData
        {
            public List<DiskDriveInfo> Disks { get; } = new List<DiskDriveInfo>();
            public string ErrorMessage { get; set; } = "";
        }

        private class DiskDriveInfo
        {
            public int Number { get; set; }
            public string Model { get; set; } = "Unknown";
            public string Capacity { get; set; } = "Unknown";
            public double SizeGB { get; set; }
            public string Status { get; set; } = "Unknown";
            public string MediaType { get; set; } = "Unknown";
            public string DeviceId { get; set; } = "Unknown";
            public string BusType { get; set; } = "Unknown";
            public string SerialNumber { get; set; } = "Unknown";
            public string ErrorMessage { get; set; } = "";
            public List<DiskPartitionInfo> Partitions { get; } = new List<DiskPartitionInfo>();
        }

        private class DiskPartitionInfo
        {
            public string Name { get; set; } = "Unknown";
            public string DeviceId { get; set; } = "Unknown";
            public string DriveLetter { get; set; } = "N/A";
            public double SizeGB { get; set; }
            public double FreeSpaceGB { get; set; }
            public double UsedSpaceGB { get; set; }
            public string ErrorMessage { get; set; } = "";
        }

        private class DiskLookupInfo
        {
            public Dictionary<string, string> BusTypeMap { get; } = new Dictionary<string, string>();
            public Dictionary<string, string> PhysicalDiskSerialMap { get; } = new Dictionary<string, string>();
            public Dictionary<string, string> PhysicalMediaSerialMap { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        private class NetworkAddressModeLookup
        {
            public Dictionary<string, string> ByInterfaceId { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, string> ByMacAddress { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, string> ByDescription { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }
    }
}
