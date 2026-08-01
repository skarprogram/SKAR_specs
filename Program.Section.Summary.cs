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
        private static async Task<SummaryReportData> CollectSummaryReportDataAsync()
        {
            return await Task.Run(() => {
                SummaryReportData reportData = new SummaryReportData();

                try
                {
                    var tasks = new List<Task>
                    {
                        RunSummaryDataSubtask("Computer name", () => reportData.ComputerName = Environment.MachineName),
                        RunSummaryDataSubtask("Computer system", () => CollectSummaryComputerSystemData(reportData)),
                        RunSummaryDataSubtask("Serial number", () => reportData.SerialNumber = CollectSummarySerialNumber()),
                        RunSummaryDataSubtask("Processor", () => CollectSummaryProcessorData(reportData)),
                        RunSummaryDataSubtask("RAM", () => reportData.Ram = CollectSummaryRam()),
                        RunSummaryDataSubtask("GPU", () => reportData.Gpu = CollectSummaryGpu()),
                        RunSummaryDataSubtask("Network devices", () => CollectSummaryNetworkData(reportData)),
                        RunSummaryDataSubtask("User and TPM", () => CollectSummaryUserAndTpmData(reportData)),
                        RunSummaryDataSubtask("Domain workgroup", () => reportData.Workgroup = CollectSummaryWorkgroup()),
                        RunSummaryDataSubtask("Operating system", () => reportData.OperatingSystem = CollectSummaryOperatingSystem()),
                        RunSummaryDataSubtask("Remote software", () => reportData.RemoteSoftware = CollectRemoteSoftwareSummaryInfo())
                    };

                    Task.WaitAll(tasks.ToArray());
                }
                catch (Exception ex)
                {
                    reportData.ErrorMessage = ex.Message;
                }

                return reportData;
            });
        }

        private static Task RunSummaryDataSubtask(string taskName, Action taskAction)
        {
            return Task.Run(() => {
                DebugLog($"SUMMARY detail: START {taskName}");
                Stopwatch stopwatch = Stopwatch.StartNew();
                try
                {
                    taskAction();
                    DebugLog($"SUMMARY detail: DONE {taskName} in {stopwatch.Elapsed.TotalSeconds:F2}s");
                }
                catch (Exception ex)
                {
                    DebugLog($"SUMMARY detail: FAIL {taskName} after {stopwatch.Elapsed.TotalSeconds:F2}s: {ex.Message}");
                    throw;
                }
            });
        }

        private static void CollectSummaryComputerSystemData(SummaryReportData reportData)
        {
            try
            {
                using (ManagementClass managementClass = new ManagementClass("Win32_ComputerSystem"))
                using (ManagementObjectCollection collection = managementClass.GetInstances())
                {
                    foreach (ManagementObject obj in collection)
                    {
                        reportData.Manufacturer = obj["Manufacturer"]?.ToString() ?? "Not detected";
                        reportData.Model = obj["Model"]?.ToString() ?? "Not detected";
                        return;
                    }
                }

                reportData.Manufacturer = "Not detected";
                reportData.Model = "Not detected";
            }
            catch
            {
                reportData.Manufacturer = "Error retrieving data";
                reportData.Model = "Error retrieving data";
            }
        }

        private static string CollectSummarySerialNumber()
        {
            try
            {
                string serialNumber = "Serial number not detected";
                using (ManagementClass managementClass = new ManagementClass("Win32_BIOS"))
                using (ManagementObjectCollection collection = managementClass.GetInstances())
                {
                    foreach (ManagementObject obj in collection)
                    {
                        if (obj["SerialNumber"] != null && !string.IsNullOrWhiteSpace(obj["SerialNumber"].ToString()))
                        {
                            serialNumber = obj["SerialNumber"].ToString();
                        }
                    }
                }

                return serialNumber;
            }
            catch
            {
                return "Error retrieving serial number";
            }
        }

        private static void CollectSummaryProcessorData(SummaryReportData reportData)
        {
            try
            {
                bool processorFound = false;
                int processorCount = 0;

                using (ManagementClass managementClass = new ManagementClass("Win32_Processor"))
                using (ManagementObjectCollection collection = managementClass.GetInstances())
                {
                    foreach (ManagementObject obj in collection)
                    {
                        string processorName = obj["Name"]?.ToString() ?? "Unknown processor";
                        processorName = processorName.Replace("(R)", "").Replace("(TM)", "");
                        processorName = System.Text.RegularExpressions.Regex.Replace(processorName, @"\s+", " ").Trim();

                        if (processorCount == 0)
                        {
                            reportData.Processor = processorName;
                        }

                        processorCount++;
                        processorFound = true;
                    }
                }

                if (!processorFound)
                {
                    reportData.Processor = "No processor detected";
                }
                else if (processorCount > 1)
                {
                    reportData.TotalPhysicalProcessors = processorCount;
                }
            }
            catch
            {
                reportData.Processor = "Error retrieving processor information";
            }
        }

        private static string CollectSummaryRam()
        {
            try
            {
                double totalRam = 0;
                string ramType = "Unknown";
                bool ramInfoFound = false;

                using (ManagementClass managementClass = new ManagementClass("Win32_ComputerSystem"))
                using (ManagementObjectCollection collection = managementClass.GetInstances())
                {
                    foreach (ManagementObject obj in collection)
                    {
                        if (obj["TotalPhysicalMemory"] != null)
                        {
                            totalRam = Convert.ToDouble(obj["TotalPhysicalMemory"]) / 1073741824;
                            ramInfoFound = true;
                        }

                        break;
                    }
                }

                try
                {
                    int ramModules = 0;
                    double maxSpeed = 0;

                    using (ManagementClass managementClass = new ManagementClass("Win32_PhysicalMemory"))
                    using (ManagementObjectCollection collection = managementClass.GetInstances())
                    {
                        foreach (ManagementObject obj in collection)
                        {
                            if (obj["MemoryType"] != null)
                            {
                                switch (Convert.ToUInt32(obj["MemoryType"]))
                                {
                                    case 20: ramType = "DDR"; break;
                                    case 21: ramType = "DDR2"; break;
                                    case 22: ramType = "DDR3"; break;
                                    case 24: ramType = "DDR4"; break;
                                    case 26: ramType = "DDR5"; break;
                                    default: ramType = "Unknown"; break;
                                }
                            }

                            if (obj["Speed"] != null)
                            {
                                uint speed = Convert.ToUInt32(obj["Speed"]);
                                if (speed > maxSpeed)
                                {
                                    maxSpeed = speed;
                                }
                            }

                            ramModules++;
                        }
                    }

                    if (ramInfoFound)
                    {
                        StringBuilder ramDetails = new StringBuilder();
                        ramDetails.Append($"{totalRam:F2} GB, RAM Type: {ramType}");

                        if (maxSpeed > 0)
                        {
                            ramDetails.Append($", Speed: {maxSpeed} MHz");
                        }

                        if (ramModules > 0)
                        {
                            ramDetails.Append($", Modules: {ramModules}");
                        }

                        return ramDetails.ToString();
                    }
                }
                catch
                {
                    if (ramInfoFound)
                    {
                        return $"{totalRam:F2} GB, RAM Type: {ramType}";
                    }
                }

                return "RAM information not available";
            }
            catch
            {
                return "Error retrieving RAM information";
            }
        }

        private static string CollectSummaryGpu()
        {
            try
            {
                List<string> gpuNames = new List<string>();
                string[] excludePatterns = new string[]
                {
                    "Microsoft Remote Display Adapter", "Microsoft Hyper-V Video", "Microsoft Basic Display Adapter",
                    "Microsoft Basic Render Driver", "Citrix Display Driver", "VMware SVGA", "VirtualBox Graphics Adapter",
                    "Parsec Virtual Display", "Virtual Display", "TeamViewer", "Splashtop", "AnyDesk", "Remote Desktop",
                    "Indirect Display", "USB Display Adapter", "DisplayLink"
                };

                using (ManagementClass managementClass = new ManagementClass("Win32_VideoController"))
                using (ManagementObjectCollection collection = managementClass.GetInstances())
                {
                    foreach (ManagementObject obj in collection)
                    {
                        string gpuName = obj["Name"]?.ToString() ?? "Unknown GPU";
                        if (!excludePatterns.Any(pattern => gpuName.IndexOf(pattern, StringComparison.OrdinalIgnoreCase) >= 0))
                        {
                            gpuNames.Add(gpuName);
                        }
                    }
                }

                return gpuNames.Count > 0 ? string.Join(", ", gpuNames) : "No physical GPU detected";
            }
            catch
            {
                return "Error retrieving GPU information";
            }
        }

        private static void CollectSummaryNetworkData(SummaryReportData reportData)
        {
            try
            {
                NetworkAddressModeLookup addressModeLookup = BuildNetworkAddressModeLookup();
                try
                {
                    foreach (NetworkInterface nic in NetworkInterface.GetAllNetworkInterfaces())
                    {
                        try
                        {
                            if (!IsSummaryNetworkAdapter(nic))
                            {
                                continue;
                            }

                            string macAddress = "N/A";
                            try
                            {
                                byte[] macBytes = nic.GetPhysicalAddress().GetAddressBytes();
                                if (macBytes != null && macBytes.Length > 0)
                                {
                                    macAddress = string.Join(":", macBytes.Select(b => b.ToString("X2")));
                                }
                            }
                            catch
                            {
                                macAddress = "N/A";
                            }

                            string adapterName = nic.Name ?? "Unknown";
                            SummaryNetworkAdapterInfo adapterInfo = new SummaryNetworkAdapterInfo
                            {
                                Name = adapterName,
                                MacAddress = macAddress,
                                IpAssignment = GetNetworkAddressAssignmentMode(nic, addressModeLookup),
                                Status = nic.OperationalStatus.ToString(),
                                ConnectionType = nic.NetworkInterfaceType.ToString()
                            };

                            try
                            {
                                foreach (UnicastIPAddressInformation ip in nic.GetIPProperties().UnicastAddresses)
                                {
                                    if (ip.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                                    {
                                        adapterInfo.Ipv4Addresses.Add(ip.Address.ToString());
                                    }
                                }
                            }
                            catch
                            {
                                // Leave IPv4 address as N/A if adapter properties cannot be read.
                            }

                            reportData.NetworkAdapters.Add(adapterInfo);
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }
                catch
                {
                    reportData.NetworkAdaptersErrorMessage = "Error accessing network adapters";
                }
            }
            catch (Exception ex)
            {
                reportData.NetworkAdaptersErrorMessage = $"Error retrieving network information: {ex.Message}";
            }
        }

        private static void CollectSummaryUserAndTpmData(SummaryReportData reportData)
        {
            try
            {
                string currentUser = Environment.UserName ?? "Unknown";
                string consoleUser = "";
                bool consoleUserDetected = false;
                try
                {
                    string output = GetConsoleUserName();
                    if (!string.IsNullOrEmpty(output))
                    {
                        string[] parts = output.Split('\\');
                        string username = parts.Length > 1 ? parts[1] : output;
                        consoleUserDetected = true;
                        if (username != currentUser)
                        {
                            consoleUser = username;
                        }
                    }
                }
                catch
                {
                    // If console user detection fails, just show current user.
                }

                string userString = currentUser;
                if (!string.IsNullOrEmpty(consoleUser))
                {
                    userString += " (Console: " + consoleUser + ")";
                }
                else if (consoleUserDetected)
                {
                    userString += " (Console: Same)";
                }

                reportData.CurrentUser = userString;
            }
            catch
            {
                reportData.CurrentUser = "Not available";
            }

            try
            {
                string fqdn = Environment.MachineName;
                try
                {
                    string domainName = IPGlobalProperties.GetIPGlobalProperties().DomainName;
                    if (!string.IsNullOrWhiteSpace(domainName))
                    {
                        fqdn = $"{Environment.MachineName}.{domainName}";
                    }
                }
                catch
                {
                    // Use machine name if domain name can't be retrieved.
                }

                reportData.FullyQualifiedDomainName = fqdn;
            }
            catch
            {
                reportData.FullyQualifiedDomainName = "Not available";
            }

            try
            {
                reportData.Tpm = GetTpmInfo();
            }
            catch (Exception ex)
            {
                reportData.Tpm = $"Error: {ex.Message}";
            }
        }

        private static string CollectSummaryWorkgroup()
        {
            try
            {
                bool partOfDomain = false;
                string workgroup = "WORKGROUP";
                bool domainInfoFound = false;

                using (ManagementClass managementClass = new ManagementClass("Win32_ComputerSystem"))
                using (ManagementObjectCollection collection = managementClass.GetInstances())
                {
                    foreach (ManagementObject obj in collection)
                    {
                        if (obj["PartOfDomain"] != null)
                        {
                            partOfDomain = Convert.ToBoolean(obj["PartOfDomain"]);
                        }

                        if (obj["Workgroup"] != null && !string.IsNullOrWhiteSpace(obj["Workgroup"].ToString()))
                        {
                            workgroup = obj["Workgroup"].ToString();
                        }

                        domainInfoFound = true;
                        break;
                    }
                }

                if (domainInfoFound && !partOfDomain)
                {
                    return workgroup;
                }

                return domainInfoFound ? "" : "Information not available";
            }
            catch
            {
                return "Error retrieving information";
            }
        }

        private static string CollectSummaryOperatingSystem()
        {
            try
            {
                using (ManagementClass managementClass = new ManagementClass("Win32_OperatingSystem"))
                using (ManagementObjectCollection collection = managementClass.GetInstances())
                {
                    foreach (ManagementObject obj in collection)
                    {
                        try
                        {
                            string caption = (obj["Caption"]?.ToString() ?? "Unknown OS").Replace("Microsoft ", "").Trim();
                            string buildNumber = obj["BuildNumber"]?.ToString() ?? "Unknown";
                            string architecture = (obj["OSArchitecture"]?.ToString() ?? "Unknown").Replace("-bit", "");
                            string versionName = GetWindowsVersionName(buildNumber);

                            string condensedOS = caption;
                            if (!string.IsNullOrEmpty(architecture) && architecture != "Unknown")
                            {
                                condensedOS += " " + architecture;
                            }

                            if (!string.IsNullOrEmpty(versionName))
                            {
                                condensedOS += " " + versionName;
                            }

                            if (!string.IsNullOrEmpty(buildNumber) && buildNumber != "Unknown")
                            {
                                condensedOS += " " + buildNumber;
                            }

                            condensedOS += " (" + GetBootMode() + ")";
                            return condensedOS;
                        }
                        catch (Exception ex)
                        {
                            return $"Error retrieving OS info: {ex.Message}";
                        }
                    }
                }

                return "Operating system information not available";
            }
            catch
            {
                return "Error retrieving information";
            }
        }

        private static string GetWindowsVersionName(string buildNumber)
        {
            int build;
            if (!int.TryParse(buildNumber, out build))
            {
                return "";
            }

            if (build >= 22631) return "23H2";
            if (build >= 22621) return "22H2";
            if (build >= 22000) return "21H2";
            if (build >= 19045) return "22H2";
            if (build >= 19044) return "21H2";
            if (build >= 19043) return "21H1";
            if (build >= 19042) return "20H2";
            if (build >= 19041) return "2004";
            if (build >= 18363) return "1909";
            if (build >= 18362) return "1903";
            if (build >= 17763) return "1809";
            if (build >= 17134) return "1803";

            return "";
        }

        private static string RenderSummaryInfoHtml(SummaryReportData summaryData)
        {
            if (summaryData == null)
            {
                return TableRowText("SUMMARY", "Error retrieving summary information");
            }

            if (!string.IsNullOrWhiteSpace(summaryData.ErrorMessage))
            {
                return TableRowText("SUMMARY", $"Error retrieving summary information: {summaryData.ErrorMessage}");
            }

            StringBuilder result = new StringBuilder();

            AppendSummaryTextRow(result, "Computer name:", summaryData.ComputerName);
            AppendSummaryTextRow(result, "Manufacturer:", summaryData.Manufacturer);
            AppendSummaryTextRow(result, "Model:", summaryData.Model);
            AppendSummaryTextRow(result, "Serial Number:", summaryData.SerialNumber);
            AppendSummaryTextRow(result, "Processor:", summaryData.Processor);

            if (summaryData.TotalPhysicalProcessors > 1)
            {
                AppendSummaryTextRow(result, "Total Physical Processors:", summaryData.TotalPhysicalProcessors.ToString());
            }

            AppendSummaryTextRow(result, "RAM:", summaryData.Ram);
            result.Append(TableRowHtml("Hard drives:", HardDriveSummaryToken));
            AppendSummaryTextRow(result, "GPU:", summaryData.Gpu);
            result.Append(TableRowHtml("Network Devices", RenderSummaryNetworkAdaptersHtml(summaryData)));
            AppendSummaryTextRow(result, "Current User:", summaryData.CurrentUser);
            AppendSummaryTextRow(result, "Full Qualified Domain Name", summaryData.FullyQualifiedDomainName);
            AppendSummaryTextRow(result, "TPM", summaryData.Tpm);

            if (!string.IsNullOrWhiteSpace(summaryData.Workgroup))
            {
                AppendSummaryTextRow(result, "Workgroup:", summaryData.Workgroup);
            }

            AppendSummaryTextRow(result, "OS:", summaryData.OperatingSystem);
            AppendSummaryTextRow(result, "Office:", summaryData.Office);
            result.Append(TableRowHtml("Remote Software:", RenderRemoteSoftwareSummaryHtml(summaryData.RemoteSoftware)));

            return result.ToString();
        }

        private static void AppendSummaryTextRow(StringBuilder result, string key, string value)
        {
            result.Append(TableRowText(key, string.IsNullOrWhiteSpace(value) ? "Not available" : value));
        }

        private static RemoteSoftwareSummaryInfo CollectRemoteSoftwareSummaryInfo()
        {
            RemoteSoftwareSummaryInfo remoteSoftware = new RemoteSoftwareSummaryInfo();
            remoteSoftware.AnyDeskId = CollectAnyDeskId();
            remoteSoftware.TeamViewerId = CollectTeamViewerId();
            return remoteSoftware;
        }

        private static string CollectAnyDeskId()
        {
            foreach (string executablePath in GetAnyDeskExecutableCandidates())
            {
                try
                {
                    ProcessStartInfo psi = new ProcessStartInfo
                    {
                        FileName = executablePath,
                        Arguments = "--get-id"
                    };

                    string workingDirectory = Path.GetDirectoryName(executablePath);
                    if (!string.IsNullOrWhiteSpace(workingDirectory))
                    {
                        psi.WorkingDirectory = workingDirectory;
                    }

                    ProcessRunResult processResult = RunProcessWithTimeout(psi, QuickExternalCommandTimeoutMs, "AnyDesk ID lookup");
                    string anyDeskId = ExtractNumericRemoteSoftwareId(processResult.CombinedOutput);
                    if (!string.IsNullOrWhiteSpace(anyDeskId))
                    {
                        return anyDeskId;
                    }
                }
                catch
                {
                    // AnyDesk is optional summary information.
                }
            }

            return "";
        }

        private static IEnumerable<string> GetAnyDeskExecutableCandidates()
        {
            HashSet<string> candidates = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            AddAnyDeskRegistryExecutableCandidates(candidates, RegistryHive.LocalMachine);
            AddAnyDeskRegistryExecutableCandidates(candidates, RegistryHive.CurrentUser);

            AddAnyDeskPathCandidate(candidates, Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), @"AnyDesk\AnyDesk.exe");
            AddAnyDeskPathCandidate(candidates, Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), @"AnyDesk\AnyDesk.exe");
            AddAnyDeskPathCandidate(candidates, Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Programs\AnyDesk\AnyDesk.exe");
            AddAnyDeskPathCandidate(candidates, Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"AnyDesk\AnyDesk.exe");

            return candidates;
        }

        private static void AddAnyDeskRegistryExecutableCandidates(HashSet<string> candidates, RegistryHive hive)
        {
            foreach (RegistryView registryView in GetRegistryViewsForOperatingSystem())
            {
                try
                {
                    using (RegistryKey baseKey = RegistryKey.OpenBaseKey(hive, registryView))
                    using (RegistryKey appPathKey = baseKey.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\AnyDesk.exe"))
                    {
                        AddAnyDeskExecutableCandidate(candidates, appPathKey?.GetValue(null)?.ToString());
                    }
                }
                catch
                {
                    // Registry app paths are a best-effort shortcut.
                }
            }
        }

        private static void AddAnyDeskPathCandidate(HashSet<string> candidates, string basePath, string relativePath)
        {
            if (string.IsNullOrWhiteSpace(basePath) || string.IsNullOrWhiteSpace(relativePath))
            {
                return;
            }

            AddAnyDeskExecutableCandidate(candidates, Path.Combine(basePath, relativePath));
        }

        private static void AddAnyDeskExecutableCandidate(HashSet<string> candidates, string executablePath)
        {
            if (candidates == null || string.IsNullOrWhiteSpace(executablePath))
            {
                return;
            }

            string validatedPath;
            if (TryGetValidatedLocalExecutablePath(executablePath, out validatedPath))
            {
                candidates.Add(validatedPath);
            }
        }

        private static string CollectTeamViewerId()
        {
            foreach (RegistryHive hive in new[] { RegistryHive.LocalMachine, RegistryHive.CurrentUser })
            {
                foreach (RegistryView registryView in GetRegistryViewsForOperatingSystem())
                {
                    foreach (string registryPath in new[] { @"SOFTWARE\TeamViewer" })
                    {
                        try
                        {
                            using (RegistryKey baseKey = RegistryKey.OpenBaseKey(hive, registryView))
                            using (RegistryKey teamViewerKey = baseKey.OpenSubKey(registryPath))
                            {
                                string teamViewerId = FormatRemoteSoftwareRegistryId(teamViewerKey?.GetValue("ClientID"));
                                if (!string.IsNullOrWhiteSpace(teamViewerId))
                                {
                                    return teamViewerId;
                                }
                            }
                        }
                        catch
                        {
                            // TeamViewer may not be installed, or access to a view may be unavailable.
                        }
                    }
                }
            }

            return "";
        }

        private static IEnumerable<RegistryView> GetRegistryViewsForOperatingSystem()
        {
            if (Environment.Is64BitOperatingSystem)
            {
                yield return RegistryView.Registry64;
                yield return RegistryView.Registry32;
            }
            else
            {
                yield return RegistryView.Registry32;
            }
        }

        private static string FormatRemoteSoftwareRegistryId(object value)
        {
            if (value == null)
            {
                return "";
            }

            string stringValue = value as string;
            if (stringValue != null)
            {
                return ExtractNumericRemoteSoftwareId(stringValue);
            }

            try
            {
                if (value is int)
                {
                    return unchecked((uint)(int)value).ToString(System.Globalization.CultureInfo.InvariantCulture);
                }

                if (value is uint)
                {
                    return ((uint)value).ToString(System.Globalization.CultureInfo.InvariantCulture);
                }

                if (value is long)
                {
                    long longValue = (long)value;
                    return longValue < 0 ? "" : longValue.ToString(System.Globalization.CultureInfo.InvariantCulture);
                }

                if (value is ulong)
                {
                    return ((ulong)value).ToString(System.Globalization.CultureInfo.InvariantCulture);
                }

                return Convert.ToUInt64(value).ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
            catch
            {
                return ExtractNumericRemoteSoftwareId(value.ToString());
            }
        }

        private static string ExtractNumericRemoteSoftwareId(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return "";
            }

            System.Text.RegularExpressions.Match match = System.Text.RegularExpressions.Regex.Match(text, @"\b\d(?:[\s-]?\d){5,}\b");
            if (!match.Success)
            {
                return "";
            }

            return System.Text.RegularExpressions.Regex.Replace(match.Value, @"[\s-]", "");
        }

        private static string RenderRemoteSoftwareSummaryText(RemoteSoftwareSummaryInfo remoteSoftware)
        {
            string anyDeskId = remoteSoftware == null || string.IsNullOrWhiteSpace(remoteSoftware.AnyDeskId)
                ? "Not detected"
                : remoteSoftware.AnyDeskId;
            string teamViewerId = remoteSoftware == null || string.IsNullOrWhiteSpace(remoteSoftware.TeamViewerId)
                ? "Not detected"
                : remoteSoftware.TeamViewerId;

            return $"Anydesk - {anyDeskId} TeamViewer - {teamViewerId}";
        }

        private static string RenderRemoteSoftwareSummaryHtml(RemoteSoftwareSummaryInfo remoteSoftware)
        {
            string anyDeskId = remoteSoftware == null ? "" : remoteSoftware.AnyDeskId;
            string teamViewerId = remoteSoftware == null ? "" : remoteSoftware.TeamViewerId;
            string anyDeskDisplay = string.IsNullOrWhiteSpace(anyDeskId) ? "Not detected" : anyDeskId;
            string teamViewerDisplay = string.IsNullOrWhiteSpace(teamViewerId) ? "Not detected" : teamViewerId;

            return $"<span data-field=\"anydesk-id\" data-value=\"{HtmlAttr(anyDeskId)}\">Anydesk - {HtmlText(anyDeskDisplay)}</span> <span data-field=\"teamviewer-id\" data-value=\"{HtmlAttr(teamViewerId)}\">TeamViewer - {HtmlText(teamViewerDisplay)}</span>";
        }

        private static string RenderSummaryNetworkAdaptersHtml(SummaryReportData summaryData)
        {
            StringBuilder networkInfo = new StringBuilder();
            networkInfo.Append("<table border='1' class='network-adapters-table' data-table-type='network-adapters' style='width:100%; font-size:12px; border-collapse:collapse;'>");
            networkInfo.Append("<thead><tr><th>Name</th><th>MAC Address</th><th>IPv4 Address</th><th>IP Assignment</th><th>Status</th><th>Connection Type</th></tr></thead><tbody>");

            if (!string.IsNullOrWhiteSpace(summaryData.NetworkAdaptersErrorMessage))
            {
                networkInfo.Append($"<tr><td colspan='6'>{HtmlText(summaryData.NetworkAdaptersErrorMessage)}</td></tr>");
            }
            else if (summaryData.NetworkAdapters.Count == 0)
            {
                networkInfo.Append("<tr><td colspan='6'>No network adapters found</td></tr>");
            }
            else
            {
                foreach (SummaryNetworkAdapterInfo adapter in summaryData.NetworkAdapters)
                {
                    string ipv4AddressHtml = adapter.Ipv4Addresses.Count > 0
                        ? string.Join("<br>", adapter.Ipv4Addresses.Select(HtmlText))
                        : "N/A";

                    networkInfo.Append($"<tr data-adapter-name=\"{HtmlAttr(adapter.Name)}\"><td data-field=\"name\">{HtmlText(adapter.Name)}</td><td data-field=\"mac\">{HtmlText(adapter.MacAddress)}</td><td data-field=\"ipv4-address\">{ipv4AddressHtml}</td><td data-field=\"ip-assignment\">{HtmlText(adapter.IpAssignment)}</td><td data-field=\"status\">{HtmlText(adapter.Status)}</td><td data-field=\"type\">{HtmlText(adapter.ConnectionType)}</td></tr>");
                }
            }

            networkInfo.Append("</tbody></table>");
            networkInfo.Append(ExternalIpSummaryToken);
            return networkInfo.ToString();
        }
    }
}
