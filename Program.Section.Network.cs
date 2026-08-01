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
        private static async Task<NetworkReportData> CollectNetworkReportDataAsync()
        {
            return await Task.Run(async () => {
                NetworkReportData reportData = new NetworkReportData();

                try
                {
                    Task ipConfigTask = Task.Run(() => CollectNetworkConfiguration(reportData));
                    Task adaptersTask = Task.Run(() => CollectNetworkAdapters(reportData));
                    Task publicIpTask = Task.Run(async () => await CollectInternetConnectivity(reportData));
                    Task sharesTask = Task.Run(() => CollectNetworkShares(reportData));
                    Task mappedDrivesTask = Task.Run(() => CollectMappedNetworkDrives(reportData));

                    await Task.WhenAll(ipConfigTask, adaptersTask, publicIpTask, sharesTask, mappedDrivesTask);
                }
                catch (Exception ex)
                {
                    reportData.ErrorMessage = ex.Message;
                }

                return reportData;
            });
        }

        private static void CollectNetworkConfiguration(NetworkReportData reportData)
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = GetSystemExecutablePath("ipconfig.exe"),
                    Arguments = "/all",
                };

                ProcessRunResult processResult = RunProcessWithTimeout(psi, ExternalCommandTimeoutMs, "ipconfig /all");
                string output = processResult.Output;
                if (!processResult.TimedOut && !string.IsNullOrWhiteSpace(output))
                {
                    reportData.IpConfigText = System.Text.RegularExpressions.Regex.Replace(
                        output,
                        @"([0-9A-Fa-f]{2})-([0-9A-Fa-f]{2})-([0-9A-Fa-f]{2})-([0-9A-Fa-f]{2})-([0-9A-Fa-f]{2})-([0-9A-Fa-f]{2})",
                        "$1:$2:$3:$4:$5:$6");
                }
                else
                {
                    reportData.IpConfigErrorMessage = processResult.TimedOut
                        ? "Network configuration lookup timed out"
                        : "No network configuration data available";
                }
            }
            catch
            {
                reportData.IpConfigErrorMessage = "Error running ipconfig command";
            }
        }

        private static void CollectNetworkAdapters(NetworkReportData reportData)
        {
            try
            {
                NetworkAddressModeLookup addressModeLookup = BuildNetworkAddressModeLookup();
                foreach (NetworkInterface nic in NetworkInterface.GetAllNetworkInterfaces())
                {
                    if (nic.NetworkInterfaceType == NetworkInterfaceType.Loopback)
                    {
                        continue;
                    }

                    NetworkAdapterReportItem adapter = new NetworkAdapterReportItem
                    {
                        Name = nic.Name,
                        MacAddress = nic.GetPhysicalAddress().ToString(),
                        Status = nic.OperationalStatus.ToString(),
                        ConnectionType = nic.NetworkInterfaceType.ToString(),
                        IpAssignment = GetNetworkAddressAssignmentMode(nic, addressModeLookup)
                    };

                    if (nic.Speed > 0)
                    {
                        if (nic.Speed >= 1000000000)
                        {
                            adapter.Speed = $"{nic.Speed / 1000000000.0:F1} Gbps";
                        }
                        else if (nic.Speed >= 1000000)
                        {
                            adapter.Speed = $"{nic.Speed / 1000000.0:F0} Mbps";
                        }
                        else
                        {
                            adapter.Speed = $"{nic.Speed / 1000.0:F0} Kbps";
                        }
                    }

                    foreach (UnicastIPAddressInformation ip in nic.GetIPProperties().UnicastAddresses)
                    {
                        if (ip.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork ||
                            ip.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetworkV6)
                        {
                            adapter.IpAddresses.Add(ip.Address.ToString());
                        }
                    }

                    reportData.Adapters.Add(adapter);
                }
            }
            catch
            {
                reportData.AdaptersErrorMessage = "Error retrieving network adapter information";
            }
        }

        private static async Task CollectInternetConnectivity(NetworkReportData reportData)
        {
            try
            {
                List<ExternalIpServiceResult> externalIpResults = await GetExternalIpServiceResultsAsync();
                reportData.ExternalIpResults.AddRange(externalIpResults);
            }
            catch
            {
                // External IP rows are optional; renderer keeps the current fallback row.
            }

            try
            {
                string[] hosts = { "8.8.8.8", "1.1.1.1", "9.9.9.9" };
                foreach (string host in hosts)
                {
                    try
                    {
                        using (Ping ping = new Ping())
                        {
                            PingReply reply = ping.Send(host, 1000);
                            if (reply.Status == IPStatus.Success)
                            {
                                reportData.InternetConnectivityChecked = true;
                                reportData.InternetConnected = true;
                                reportData.InternetConnectivityStatus = reply.RoundtripTime.ToString() + "ms";
                                return;
                            }
                        }
                    }
                    catch
                    {
                        // Try next host.
                    }
                }

                reportData.InternetConnectivityChecked = true;
                reportData.InternetConnected = false;
                reportData.InternetConnectivityStatus = "Failed";
            }
            catch
            {
                reportData.InternetConnectivityCheckFailed = true;
            }
        }

        private static void CollectNetworkShares(NetworkReportData reportData)
        {
            try
            {
                using (ManagementClass managementClass = new ManagementClass("Win32_Share"))
                using (ManagementObjectCollection collection = managementClass.GetInstances())
                {
                    foreach (ManagementObject obj in collection)
                    {
                        try
                        {
                            reportData.Shares.Add(new NetworkShareInfo
                            {
                                Name = obj["Name"]?.ToString() ?? "Unknown",
                                Path = obj["Path"]?.ToString() ?? "Unknown",
                                Description = obj["Description"]?.ToString() ?? "Unknown"
                            });
                        }
                        catch
                        {
                            reportData.Shares.Add(new NetworkShareInfo { ErrorMessage = "Error retrieving share information" });
                        }
                    }
                }
            }
            catch
            {
                reportData.SharesErrorMessage = "Error retrieving network shares";
            }
        }

        private static void CollectMappedNetworkDrives(NetworkReportData reportData)
        {
            try
            {
                using (ManagementClass managementClass = new ManagementClass("Win32_MappedLogicalDisk"))
                using (ManagementObjectCollection collection = managementClass.GetInstances())
                {
                    foreach (ManagementObject obj in collection)
                    {
                        try
                        {
                            reportData.MappedDrives.Add(new MappedNetworkDriveInfo
                            {
                                DriveLetter = obj["DeviceID"]?.ToString() ?? "Unknown",
                                RemotePath = obj["ProviderName"]?.ToString() ?? "Unknown",
                                Status = obj["Status"]?.ToString() ?? "Unknown"
                            });
                        }
                        catch
                        {
                            reportData.MappedDrives.Add(new MappedNetworkDriveInfo { ErrorMessage = "Error retrieving mapped drive information" });
                        }
                    }
                }
            }
            catch
            {
                reportData.MappedDrivesErrorMessage = "Error retrieving mapped drives";
            }
        }

        private static string RenderNetworkInfoHtml(NetworkReportData networkData)
        {
            if (networkData == null)
            {
                return TableRowText("Network", "Error retrieving network information");
            }

            if (!string.IsNullOrWhiteSpace(networkData.ErrorMessage))
            {
                return TableRowText("Network", $"Error retrieving network information: {networkData.ErrorMessage}");
            }

            StringBuilder result = new StringBuilder();
            result.Append(RenderNetworkAdaptersHtml(networkData));
            result.Append(RenderInternetConnectivityHtml(networkData));
            result.Append(RenderNetworkConfigurationHtml(networkData));
            result.Append(RenderNetworkSharesHtml(networkData));
            result.Append(RenderMappedNetworkDrivesHtml(networkData));
            return result.ToString();
        }

        private static string RenderNetworkConfigurationHtml(NetworkReportData networkData)
        {
            string value = !string.IsNullOrWhiteSpace(networkData.IpConfigText)
                ? HtmlText(networkData.IpConfigText).Replace(Environment.NewLine, "<br>")
                : HtmlText(networkData.IpConfigErrorMessage);

            return string.Equals(networkData.IpConfigErrorMessage, "Error running ipconfig command", StringComparison.Ordinal)
                ? TableRowText("Network Configuration", networkData.IpConfigErrorMessage)
                : TableRowHtml("Network Configuration", value);
        }

        private static string RenderNetworkAdaptersHtml(NetworkReportData networkData)
        {
            if (!string.IsNullOrWhiteSpace(networkData.AdaptersErrorMessage))
            {
                return TableRowText("Network Adapters", networkData.AdaptersErrorMessage);
            }

            StringBuilder adaptersInfo = new StringBuilder();
            adaptersInfo.Append("<h3>Network Adapters</h3>");
            adaptersInfo.Append("<table border='1'><tr><th>Name</th><th>MAC Address</th><th>Status</th><th>Speed</th><th>Connection Type</th><th>IP Address</th><th>IP Assignment</th></tr>");

            if (networkData.Adapters.Count == 0)
            {
                adaptersInfo.Append("<tr><td colspan='7'>No network adapters found</td></tr>");
            }
            else
            {
                foreach (NetworkAdapterReportItem adapter in networkData.Adapters)
                {
                    string ipAddressHtml = adapter.IpAddresses.Count > 0
                        ? string.Join("<br>", adapter.IpAddresses.Select(HtmlText))
                        : "N/A";

                    adaptersInfo.Append($"<tr><td>{HtmlText(adapter.Name)}</td><td>{HtmlText(adapter.MacAddress)}</td><td>{HtmlText(adapter.Status)}</td><td>{HtmlText(adapter.Speed)}</td><td>{HtmlText(adapter.ConnectionType)}</td><td>{ipAddressHtml}</td><td>{HtmlText(adapter.IpAssignment)}</td></tr>");
                }
            }

            adaptersInfo.Append("</table>");
            return TableRowHtml("Network Adapters", adaptersInfo.ToString());
        }

        private static string RenderInternetConnectivityHtml(NetworkReportData networkData)
        {
            StringBuilder publicIpInfo = new StringBuilder();
            publicIpInfo.Append("<h3>External Connectivity</h3>");
            publicIpInfo.Append("<table border='1'><tr><th>Service</th><th>IP Address</th><th>Status</th></tr>");

            if (networkData.ExternalIpResults.Count == 0)
            {
                publicIpInfo.Append("<tr><td>External IP</td><td>Unable to retrieve</td><td>Failed</td></tr>");
            }

            foreach (ExternalIpServiceResult resultItem in networkData.ExternalIpResults)
            {
                string displayValue = resultItem.Success ? resultItem.IpAddress : resultItem.Message;
                if (string.IsNullOrWhiteSpace(displayValue))
                {
                    displayValue = "Failed";
                }

                publicIpInfo.Append("<tr><td>" + HtmlText(resultItem.ServiceName) + "</td><td>" + HtmlText(displayValue) + "</td><td>" + HtmlText(resultItem.Status) + "</td></tr>");
            }

            if (networkData.InternetConnectivityCheckFailed)
            {
                publicIpInfo.Append("<tr><td>Internet Connectivity</td><td>Error checking</td><td>Failed</td></tr>");
            }
            else
            {
                publicIpInfo.Append("<tr><td>Internet Connectivity</td><td>" +
                    (networkData.InternetConnected ? "Connected" : "Disconnected") + "</td><td>" + HtmlText(networkData.InternetConnectivityStatus) + "</td></tr>");
            }

            publicIpInfo.Append("</table>");
            return TableRowHtml("Internet Connectivity", publicIpInfo.ToString());
        }

        private static string RenderNetworkSharesHtml(NetworkReportData networkData)
        {
            if (!string.IsNullOrWhiteSpace(networkData.SharesErrorMessage))
            {
                return TableRowTextWithId("NetworkShares", "Network Shares", networkData.SharesErrorMessage);
            }

            StringBuilder sharesInfo = new StringBuilder();
            sharesInfo.Append("<table border='1'><tr><th>Name</th><th>Path</th><th>Description</th></tr>");

            if (networkData.Shares.Count == 0)
            {
                sharesInfo.Append("<tr><td colspan='3'>No network shares found</td></tr>");
            }
            else
            {
                foreach (NetworkShareInfo share in networkData.Shares)
                {
                    if (!string.IsNullOrWhiteSpace(share.ErrorMessage))
                    {
                        sharesInfo.Append($"<tr><td colspan='3'>{HtmlText(share.ErrorMessage)}</td></tr>");
                        continue;
                    }

                    sharesInfo.Append($"<tr><td>{HtmlText(share.Name)}</td><td>{HtmlText(share.Path)}</td><td>{HtmlText(share.Description)}</td></tr>");
                }
            }

            sharesInfo.Append("</table>");
            return TableRowHtmlWithId("NetworkShares", "Network Shares", sharesInfo.ToString());
        }

        private static string RenderMappedNetworkDrivesHtml(NetworkReportData networkData)
        {
            if (!string.IsNullOrWhiteSpace(networkData.MappedDrivesErrorMessage))
            {
                return TableRowTextWithId("Networkmappeddrives", "Network mapped drives", networkData.MappedDrivesErrorMessage);
            }

            StringBuilder mappedDrivesInfo = new StringBuilder();
            mappedDrivesInfo.Append("<table border='1'><tr><th>Drive Letter</th><th>Remote Path</th><th>Status</th></tr>");

            if (networkData.MappedDrives.Count == 0)
            {
                mappedDrivesInfo.Append("<tr><td colspan='3'>No mapped network drives found</td></tr>");
            }
            else
            {
                foreach (MappedNetworkDriveInfo drive in networkData.MappedDrives)
                {
                    if (!string.IsNullOrWhiteSpace(drive.ErrorMessage))
                    {
                        mappedDrivesInfo.Append($"<tr><td colspan='3'>{HtmlText(drive.ErrorMessage)}</td></tr>");
                        continue;
                    }

                    mappedDrivesInfo.Append($"<tr><td>{HtmlText(drive.DriveLetter)}</td><td>{HtmlText(drive.RemotePath)}</td><td>{HtmlText(drive.Status)}</td></tr>");
                }
            }

            mappedDrivesInfo.Append("</table>");
            return TableRowHtmlWithId("Networkmappeddrives", "Network mapped drives", mappedDrivesInfo.ToString());
        }
    }
}
