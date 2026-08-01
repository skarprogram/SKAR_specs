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
        static void Main(string[] args)
        {
            MainAsync(args).GetAwaiter().GetResult();
        }

        private static async Task MainAsync(string[] args)
        {
            InitializeDebugMode(args);
            StringBuilder errorLog = new StringBuilder();
            string filePath = "";
            string newFilePath = "";

            try
            {
                DebugLog("Report generation started.");
                ReportExporter exporter = GetReportExporter(args);
                string serialNumber = GetComputerSerialNumber();
                string pcName = Environment.MachineName;
                string userName = Environment.UserName;
                string dateTime = DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss");
                filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"{serialNumber} - {pcName} - {userName} - {dateTime}.tmp");
                newFilePath = Path.ChangeExtension(filePath, exporter.FileExtension);

                // Create the selected report file.
                using (StreamWriter writer = new StreamWriter(filePath))
                {
                    string navHeader = $"{DateTime.Now.ToString("yyyy-MM-dd HH:mm")}";

                    ReportCollectionResult collectionResult = await CollectReportDataAsync();
                    if (collectionResult.CollectionException != null)
                    {
                        errorLog.AppendLine($"Error during information collection: {collectionResult.CollectionException.Message}");
                    }

                    writer.Write(RenderReportOutput(
                        exporter,
                        collectionResult,
                        navHeader,
                        errorLog));
                }

                // Rename file from .tmp to .html
                if (File.Exists(filePath))
                {
                    try
                    {
                    if (File.Exists(newFilePath))
                    {
                        File.Delete(newFilePath);
                    }
                    File.Move(filePath, newFilePath);

                    }
                    catch (Exception ex)
                    {
                        if (exporter.Format == ReportOutputFormat.Html)
                        {
                            // If we can't rename the file, try to append the error to the original HTML file.
                            try
                            {
                                using (StreamWriter writer = new StreamWriter(filePath, true))
                                {
                                    writer.WriteLine("<div style='color: red; background-color: #FFEBEE; padding: 10px; border: 1px solid #D32F2F;'>");
                                    writer.WriteLine($"<h3>Error Renaming File</h3>");
                                    writer.WriteLine($"<p>Could not rename temp file to HTML: {HtmlText(ex.Message)}</p>");
                                    writer.WriteLine("</div>");
                                }
                            }
                            catch
                            {
                                // If we can't even append to the file, just let it go silently.
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Silently fail if we can't even create the file
                // No console output, no error messages
                DebugLog($"Report generation failed before the report file could be completed: {ex}");
                TryWriteStartupFailureLog(ex);
            }
            finally
            {
                if (debugConsoleAttached)
                {
                    try { FreeConsole(); } catch { }
                }
            }
        }

        private static void TryWriteStartupFailureLog(Exception ex)
        {
            try
            {
                string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SKAR_specs_error.log");
                File.AppendAllText(logPath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + ex + Environment.NewLine);
            }
            catch
            {
                // Preserve the app's historical silent-failure behavior.
            }
        }

        private static string GetComputerSerialNumber()
        {
            try
            {
                using (ManagementClass managementClass = new ManagementClass("Win32_BIOS"))
                using (ManagementObjectCollection collection = managementClass.GetInstances())
                {
                    foreach (ManagementObject obj in collection)
                    {
                        if (obj["SerialNumber"] != null && !string.IsNullOrWhiteSpace(obj["SerialNumber"].ToString()))
                        {
                            string serial = obj["SerialNumber"].ToString().Trim();
                            // Remove invalid filename characters
                            foreach (char c in Path.GetInvalidFileNameChars())
                            {
                                serial = serial.Replace(c.ToString(), "");
                            }
                            return serial;
                        }
                    }
                }
            }
            catch
            {
                // Return default if serial number cannot be retrieved
            }
            return "NoSerial";
        }

        private static string GetConsoleUserName()
        {
            try
            {
                using (ManagementClass managementClass = new ManagementClass("Win32_ComputerSystem"))
                using (ManagementObjectCollection collection = managementClass.GetInstances())
                {
                    foreach (ManagementObject obj in collection)
                    {
                        string userName = obj["UserName"]?.ToString();
                        if (!string.IsNullOrWhiteSpace(userName))
                        {
                            return userName.Trim();
                        }

                        break;
                    }
                }
            }
            catch
            {
                // Console user is optional summary information.
            }

            return string.Empty;
        }

        private static string GetBootMode()
        {
            try
            {
                RegistryView registryView = Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32;
                using (RegistryKey baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, registryView))
                using (RegistryKey controlKey = baseKey.OpenSubKey(@"SYSTEM\CurrentControlSet\Control"))
                {
                    object firmwareType = controlKey?.GetValue("PEFirmwareType");
                    if (firmwareType != null)
                    {
                        int value = Convert.ToInt32(firmwareType);
                        if (value == 1)
                        {
                            return "BIOS";
                        }

                        if (value == 2)
                        {
                            return "UEFI";
                        }
                    }
                }
            }
            catch
            {
                // Boot mode is best-effort summary information.
            }

            return "Unknown boot mode";
        }



        private static bool IsSummaryNetworkAdapter(NetworkInterface nic)
        {
            if (nic == null)
            {
                return false;
            }

            switch (nic.NetworkInterfaceType)
            {
                case NetworkInterfaceType.Ethernet:
                case NetworkInterfaceType.FastEthernetFx:
                case NetworkInterfaceType.FastEthernetT:
                case NetworkInterfaceType.GigabitEthernet:
                case NetworkInterfaceType.Wireless80211:
                    return true;
                default:
                    return false;
            }
        }

        private static NetworkAddressModeLookup BuildNetworkAddressModeLookup()
        {
            NetworkAddressModeLookup lookup = new NetworkAddressModeLookup();

            try
            {
                using (ManagementObjectSearcher searcher = new ManagementObjectSearcher(
                    "SELECT SettingID, MACAddress, Description, DHCPEnabled FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = TRUE"))
                using (ManagementObjectCollection configurations = searcher.Get())
                {
                    foreach (ManagementObject configuration in configurations)
                    {
                        try
                        {
                            object dhcpEnabledValue = configuration["DHCPEnabled"];
                            if (!(dhcpEnabledValue is bool))
                            {
                                continue;
                            }

                            string mode = ((bool)dhcpEnabledValue) ? "DHCP" : "STATIC";
                            AddInterfaceIdLookupValue(lookup.ByInterfaceId, Convert.ToString(configuration["SettingID"]), mode);
                            AddLookupValue(lookup.ByMacAddress, NormalizeMacAddress(Convert.ToString(configuration["MACAddress"])), mode);
                            AddLookupValue(lookup.ByDescription, Convert.ToString(configuration["Description"]), mode);
                        }
                        catch
                        {
                            // Skip a single malformed adapter record.
                        }
                    }
                }
            }
            catch
            {
                // WMI is a best-effort read-only fallback; registry and IP metadata can still provide the value.
            }

            return lookup;
        }

        private static void AddLookupValue(Dictionary<string, string> lookup, string key, string value)
        {
            if (lookup == null || string.IsNullOrWhiteSpace(key) || string.IsNullOrWhiteSpace(value))
            {
                return;
            }

            lookup[key.Trim()] = value;
        }

        private static void AddInterfaceIdLookupValue(Dictionary<string, string> lookup, string key, string value)
        {
            if (lookup == null || string.IsNullOrWhiteSpace(key) || string.IsNullOrWhiteSpace(value))
            {
                return;
            }

            string normalizedInterfaceId = NormalizeInterfaceId(key);
            if (!string.IsNullOrWhiteSpace(normalizedInterfaceId))
            {
                lookup[normalizedInterfaceId] = value;
                lookup["{" + normalizedInterfaceId + "}"] = value;
            }
        }

        private static string GetNetworkAddressAssignmentMode(NetworkInterface nic, NetworkAddressModeLookup lookup)
        {
            if (nic == null)
            {
                return "Unknown";
            }

            string registryMode = GetNetworkAddressAssignmentModeFromRegistry(nic.Id);
            if (!string.IsNullOrWhiteSpace(registryMode))
            {
                return registryMode;
            }

            string lookupMode = GetNetworkAddressAssignmentModeFromLookup(nic, lookup);
            if (!string.IsNullOrWhiteSpace(lookupMode))
            {
                return lookupMode;
            }

            string addressOriginMode = GetNetworkAddressAssignmentModeFromAddressOrigin(nic);
            if (!string.IsNullOrWhiteSpace(addressOriginMode))
            {
                return addressOriginMode;
            }

            return "Unknown";
        }

        private static string GetNetworkAddressAssignmentModeFromRegistry(string interfaceId)
        {
            if (string.IsNullOrWhiteSpace(interfaceId))
            {
                return null;
            }

            try
            {
                foreach (string candidateId in GetInterfaceIdCandidates(interfaceId))
                {
                    using (RegistryKey key = Registry.LocalMachine.OpenSubKey(
                        @"SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\" + candidateId))
                    {
                        object enableDhcp = key?.GetValue("EnableDHCP");
                        int enableDhcpValue;
                        if (TryConvertToInt(enableDhcp, out enableDhcpValue))
                        {
                            return enableDhcpValue == 0 ? "STATIC" : "DHCP";
                        }
                    }
                }
            }
            catch
            {
                // Local HKLM reads normally do not require elevation, but keep this non-fatal.
            }

            return null;
        }

        private static string GetNetworkAddressAssignmentModeFromLookup(NetworkInterface nic, NetworkAddressModeLookup lookup)
        {
            if (lookup == null)
            {
                return null;
            }

            string value;
            foreach (string candidateId in GetInterfaceIdCandidates(nic.Id))
            {
                if (lookup.ByInterfaceId.TryGetValue(candidateId, out value))
                {
                    return value;
                }
            }

            string macAddress = null;
            try
            {
                macAddress = NormalizeMacAddress(nic.GetPhysicalAddress()?.ToString());
            }
            catch
            {
                macAddress = null;
            }

            if (!string.IsNullOrWhiteSpace(macAddress) && lookup.ByMacAddress.TryGetValue(macAddress, out value))
            {
                return value;
            }

            if (!string.IsNullOrWhiteSpace(nic.Description) && lookup.ByDescription.TryGetValue(nic.Description, out value))
            {
                return value;
            }

            return null;
        }

        private static string GetNetworkAddressAssignmentModeFromAddressOrigin(NetworkInterface nic)
        {
            try
            {
                bool hasManualAddress = false;
                foreach (UnicastIPAddressInformation ip in nic.GetIPProperties().UnicastAddresses)
                {
                    if (ip.Address.AddressFamily != System.Net.Sockets.AddressFamily.InterNetwork)
                    {
                        continue;
                    }

                    if (ip.PrefixOrigin == PrefixOrigin.Dhcp || ip.SuffixOrigin == SuffixOrigin.OriginDhcp)
                    {
                        return "DHCP";
                    }

                    if (ip.PrefixOrigin == PrefixOrigin.Manual || ip.SuffixOrigin == SuffixOrigin.Manual)
                    {
                        hasManualAddress = true;
                    }
                }

                if (hasManualAddress)
                {
                    return "STATIC";
                }
            }
            catch
            {
                // Address-origin metadata is best effort only.
            }

            return null;
        }

        private static IEnumerable<string> GetInterfaceIdCandidates(string interfaceId)
        {
            if (string.IsNullOrWhiteSpace(interfaceId))
            {
                yield break;
            }

            string trimmed = interfaceId.Trim();
            yield return trimmed;

            string normalized = NormalizeInterfaceId(trimmed);
            if (!string.IsNullOrWhiteSpace(normalized) && !string.Equals(trimmed, normalized, StringComparison.OrdinalIgnoreCase))
            {
                yield return normalized;
            }

            string braced = "{" + normalized + "}";
            if (!string.IsNullOrWhiteSpace(normalized) && !string.Equals(trimmed, braced, StringComparison.OrdinalIgnoreCase))
            {
                yield return braced;
            }
        }

        private static string NormalizeInterfaceId(string interfaceId)
        {
            return string.IsNullOrWhiteSpace(interfaceId) ? null : interfaceId.Trim().Trim('{', '}');
        }

        private static string NormalizeMacAddress(string macAddress)
        {
            if (string.IsNullOrWhiteSpace(macAddress))
            {
                return null;
            }

            StringBuilder normalized = new StringBuilder();
            foreach (char c in macAddress)
            {
                if (Uri.IsHexDigit(c))
                {
                    normalized.Append(char.ToUpperInvariant(c));
                }
            }

            return normalized.Length == 0 ? null : normalized.ToString();
        }

        private static bool TryConvertToInt(object value, out int result)
        {
            result = 0;
            if (value == null)
            {
                return false;
            }

            try
            {
                result = Convert.ToInt32(value);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static Task<List<ExternalIpServiceResult>> GetExternalIpServiceResultsAsync()
        {
            lock (ExternalIpLookupSync)
            {
                if (externalIpLookupTask == null)
                {
                    externalIpLookupTask = Task.Run(async () => {
                        string[] serviceNames = { "ifconfig.me", "ipify.org", "ipinfo.io" };
                        string[] serviceUrls = { "https://ifconfig.me/ip", "https://api.ipify.org", "https://ipinfo.io/ip" };

                        var tasks = new List<Task<ExternalIpServiceResult>>();
                        for (int i = 0; i < serviceNames.Length; i++)
                        {
                            string serviceName = serviceNames[i];
                            string url = serviceUrls[i];
                            tasks.Add(GetExternalIpFromServiceAsync(serviceName, url));
                        }

                        ExternalIpServiceResult[] results = await Task.WhenAll(tasks);
                        return results.ToList();
                    });
                }

                return externalIpLookupTask;
            }
        }

        private static async Task<ExternalIpServiceResult> GetExternalIpFromServiceAsync(string serviceName, string url)
        {
            ExternalIpServiceResult result = new ExternalIpServiceResult
            {
                ServiceName = serviceName,
                Status = "Failed"
            };

            try
            {
                using (WebClient client = new WebClient())
                {
                    client.Headers.Add("User-Agent", "Mozilla/5.0");
                    string publicIp = await client.DownloadStringWithTimeoutAsync(url, NetworkTimeoutMs);
                    publicIp = publicIp.Trim();

                    if (!string.IsNullOrWhiteSpace(publicIp) && IsValidIPAddress(publicIp))
                    {
                        result.IpAddress = publicIp;
                        result.Status = "Success";
                        result.Success = true;
                    }
                    else
                    {
                        result.Message = "Invalid response";
                    }
                }
            }
            catch (Exception ex)
            {
                result.Message = "Connection error: " + ex.Message;
            }

            return result;
        }

        private static string RenderExternalIpSummaryHtml(NetworkReportData networkData)
        {
            string externalIp = "";
            if (networkData != null)
            {
                ExternalIpServiceResult successfulResult = networkData.ExternalIpResults
                    .FirstOrDefault(result => result.Success && !string.IsNullOrWhiteSpace(result.IpAddress));

                if (successfulResult != null)
                {
                    externalIp = successfulResult.IpAddress;
                }
            }

            StringBuilder summary = new StringBuilder();
            summary.Append("<table border='1' class='external-ip-summary-table' data-table-type='external-ip-summary' style='width:100%; font-size:12px; border-collapse:collapse;'>");
            summary.Append("<thead><tr><th>Item</th><th>Value</th></tr></thead><tbody>");

            if (!string.IsNullOrWhiteSpace(externalIp))
            {
                summary.Append($"<tr class='public-ip-info' data-public-ip='{HtmlAttr(externalIp)}'><td data-field='label'>External IP</td><td data-field='value'>{HtmlText(externalIp)}</td></tr>");
            }
            else
            {
                summary.Append("<tr class='public-ip-info' data-ip-error='connection'><td data-field='label'>External IP</td><td data-field='value'>Unable to retrieve</td></tr>");
            }

            summary.Append("</tbody></table>");

            return summary.ToString();
        }

        private static string GetOfficeSummaryText()
        {
            return string.IsNullOrWhiteSpace(officeLicenseName) ? "Not detected" : officeLicenseName;
        }

        private static string GetTpmInfo()
        {
            string specVersion;
            bool? pnpResult = GetTpmFromPnPUtil(out specVersion);

            if (pnpResult == true)
            {
                return $"Present, Version {specVersion}";
            }

            if (pnpResult == false)
            {
                return "Not present";
            }

            if (TryGetTpmFromAcpiRegistry(out specVersion))
            {
                return $"Present, Version {specVersion}";
            }

            return "Not present";
        }

        private static bool? GetTpmFromPnPUtil(out string specVersion)
        {
            specVersion = "Unknown";

            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = GetSystemExecutablePath("pnputil.exe"),
                    Arguments = "/enum-devices /connected /class SecurityDevices",
                };

                ProcessRunResult processResult = RunProcessWithTimeout(psi, QuickExternalCommandTimeoutMs, "TPM pnputil lookup");
                if (processResult.TimedOut)
                {
                    return null;
                }

                string output = processResult.CombinedOutput.Trim();
                if (string.IsNullOrWhiteSpace(output))
                {
                    return null;
                }

                if (!ContainsIgnoreCase(output, "Trusted Platform Module") &&
                    !ContainsIgnoreCase(output, "TPM") &&
                    !ContainsIgnoreCase(output, "MSFT0101"))
                {
                    return false;
                }

                specVersion = DetectTpmVersion(output);
                return true;
            }
            catch
            {
                return null;
            }
        }

        private static bool TryGetTpmFromAcpiRegistry(out string specVersion)
        {
            specVersion = "Unknown";

            // Do not use HKLM\SYSTEM\CurrentControlSet\Services\TPM\WMI for presence.
            // That key means Windows has TPM management support installed, not that hardware exists.
            if (RegistrySubKeyHasChildren(@"SYSTEM\CurrentControlSet\Enum\ACPI\MSFT0101"))
            {
                specVersion = "2.0";
                return true;
            }

            return false;
        }

        private static string DetectTpmVersion(string text)
        {
            if (ContainsIgnoreCase(text, "2.0") || ContainsIgnoreCase(text, "MSFT0101"))
            {
                return "2.0";
            }

            if (ContainsIgnoreCase(text, "1.2"))
            {
                return "1.2";
            }

            return "Unknown";
        }

        private static bool RegistrySubKeyHasChildren(string path)
        {
            try
            {
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(path))
                {
                    return key != null && key.GetSubKeyNames().Length > 0;
                }
            }
            catch
            {
                return false;
            }
        }

        private static bool ContainsIgnoreCase(string text, string value)
        {
            return text != null && value != null && text.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0;
        }


    }
}
