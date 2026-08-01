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
        private static async Task<PrintersReportData> CollectPrintersReportDataAsync()
        {
            return await Task.Factory.StartNew(() => {
                Stopwatch sectionStopwatch = Stopwatch.StartNew();
                PrintersReportData reportData = new PrintersReportData();

                try
                {
                    DebugLog("PRINTERS detail: starting Win32_Printer enumeration");

                    List<PrinterBasicInfo> printers = new List<PrinterBasicInfo>();

                    try
                    {
                        Stopwatch printerEnumStopwatch = Stopwatch.StartNew();
                        using (ManagementClass managementClass = new ManagementClass("Win32_Printer"))
                        using (ManagementObjectCollection collection = managementClass.GetInstances())
                        {
                            // Collect all printers first.
                            foreach (ManagementObject obj in collection)
                            {
                                PrinterBasicInfo printerInfo = new PrinterBasicInfo();

                                if (obj["Name"] != null)
                                {
                                    printerInfo.Name = obj["Name"].ToString();
                                }

                                if (obj["Status"] != null)
                                {
                                    printerInfo.Status = obj["Status"].ToString();
                                }

                                if (obj["PortName"] != null)
                                {
                                    printerInfo.PortName = obj["PortName"].ToString();
                                }

                                if (obj["Shared"] != null)
                                {
                                    printerInfo.Shared = Convert.ToBoolean(obj["Shared"]) ? "Yes" : "No";
                                }

                                printers.Add(printerInfo);
                            }
                        }

                        DebugLog($"PRINTERS detail: Win32_Printer enumeration found {printers.Count} printers in {printerEnumStopwatch.Elapsed.TotalSeconds:F2}s");
                    }
                    catch (Exception ex)
                    {
                        DebugLog($"PRINTERS detail: Win32_Printer enumeration failed after {sectionStopwatch.Elapsed.TotalSeconds:F2}s: {ex.Message}");
                        reportData.EnumerationErrorMessage = ex.Message;
                        return reportData;
                    }

                    Stopwatch lookupStopwatch = Stopwatch.StartNew();
                    PrinterLookupInfo lookupInfo = GetPrinterLookupInfo();
                    DebugLog($"PRINTERS detail: printer lookup maps loaded in {lookupStopwatch.Elapsed.TotalSeconds:F2}s");

                    if (printers.Count > 0)
                    {
                        var printerTasks = new List<Task<PrinterReportItem>>();

                        foreach (var printer in printers)
                        {
                            printerTasks.Add(Task.Run(() => {
                                Stopwatch printerStopwatch = Stopwatch.StartNew();
                                PrinterReportItem printerItem = new PrinterReportItem
                                {
                                    Name = printer.Name,
                                    Status = printer.Status,
                                    PortName = printer.PortName,
                                    Shared = printer.Shared
                                };

                                try
                                {
                                    printerItem.IpAddress = GetPrinterIPAddress(printer.PortName, lookupInfo);
                                    DebugLog($"PRINTERS detail: processed printer '{printer.Name}' on port '{printer.PortName}' in {printerStopwatch.Elapsed.TotalSeconds:F2}s");
                                }
                                catch (Exception ex)
                                {
                                    printerItem.ErrorMessage = ex.Message;
                                    DebugLog($"PRINTERS detail: printer '{printer.Name}' failed after {printerStopwatch.Elapsed.TotalSeconds:F2}s: {ex.Message}");
                                }

                                return printerItem;
                            }));
                        }

                        Stopwatch printerProcessStopwatch = Stopwatch.StartNew();
                        Task.WaitAll(printerTasks.ToArray());
                        DebugLog($"PRINTERS detail: processed {printerTasks.Count} printer rows in {printerProcessStopwatch.Elapsed.TotalSeconds:F2}s");

                        foreach (var task in printerTasks)
                        {
                            reportData.Printers.Add(task.Result);
                        }
                    }

                    DebugLog($"PRINTERS detail: data collection completed in {sectionStopwatch.Elapsed.TotalSeconds:F2}s");
                    return reportData;
                }
                catch (Exception ex)
                {
                    reportData.ErrorMessage = ex.Message;
                    return reportData;
                }
            }, CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default);
        }

        private static string RenderPrintersInfoHtml(PrintersReportData printerData)
        {
            if (printerData == null)
            {
                return TableRowText("Printers", "Error retrieving printer information");
            }

            if (!string.IsNullOrWhiteSpace(printerData.ErrorMessage))
            {
                return TableRowText("Printers", $"Error retrieving printer information: {printerData.ErrorMessage}");
            }

            StringBuilder printersInfo = new StringBuilder();
            printersInfo.Append("<table border='1'><tr><th>Name</th><th>Status</th><th>Port</th><th>IP Address</th><th>Shared</th></tr>");

            if (!string.IsNullOrWhiteSpace(printerData.EnumerationErrorMessage))
            {
                printersInfo.Append($"<tr><td colspan='5'>Error accessing printers: {HtmlText(printerData.EnumerationErrorMessage)}</td></tr>");
            }
            else if (printerData.Printers.Count == 0)
            {
                printersInfo.Append("<tr><td colspan='5'>No printers found</td></tr>");
            }
            else
            {
                foreach (PrinterReportItem printer in printerData.Printers)
                {
                    if (!string.IsNullOrWhiteSpace(printer.ErrorMessage))
                    {
                        printersInfo.Append($"<tr><td colspan='5'>Error retrieving printer information: {HtmlText(printer.ErrorMessage)}</td></tr>");
                        continue;
                    }

                    printersInfo.Append($"<tr><td>{HtmlText(printer.Name)}</td><td>{HtmlText(printer.Status)}</td><td>{HtmlText(printer.PortName)}</td><td>{HtmlText(printer.IpAddress)}</td><td>{HtmlText(printer.Shared)}</td></tr>");
                }
            }

            printersInfo.Append("</table>");
            DebugLog("PRINTERS detail: rendered printer HTML from structured data");

            return TableRowHtml("Printers", printersInfo.ToString());
        }

        // Helper method to extract IP address from printer port information
        private static string GetPrinterIPAddress(string portName)
        {
            return GetPrinterIPAddress(portName, GetPrinterLookupInfo());
        }

        private static PrinterLookupInfo GetPrinterLookupInfo()
        {
            PrinterLookupInfo lookupInfo = new PrinterLookupInfo();

            try
            {
                using (ManagementClass mgmtClass = new ManagementClass("Win32_TCPIPPrinterPort"))
                using (ManagementObjectCollection collection = mgmtClass.GetInstances())
                {
                    foreach (ManagementObject obj in collection)
                    {
                        string name = obj["Name"]?.ToString();
                        string hostAddress = obj["HostAddress"]?.ToString();

                        if (!string.IsNullOrWhiteSpace(name) && !string.IsNullOrWhiteSpace(hostAddress))
                        {
                            lookupInfo.TcpIpPortHosts[name] = hostAddress;
                        }
                    }
                }
            }
            catch
            {
                // Printer TCP/IP port lookup is optional.
            }

            try
            {
                using (ManagementClass managementClass = new ManagementClass("Win32_PrinterConfiguration"))
                using (ManagementObjectCollection configCollection = managementClass.GetInstances())
                {
                    foreach (ManagementObject obj in configCollection)
                    {
                        string name = obj["Name"]?.ToString();
                        string description = obj["PortDescription"]?.ToString();

                        if (!string.IsNullOrWhiteSpace(name) && !string.IsNullOrWhiteSpace(description))
                        {
                            lookupInfo.PortDescriptions[name] = description;
                        }
                    }
                }
            }
            catch
            {
                // Printer configuration lookup is optional.
            }

            return lookupInfo;
        }

        private static string GetPrinterIPAddress(string portName, PrinterLookupInfo lookupInfo)
        {
            if (string.IsNullOrEmpty(portName))
            {
                return "Unknown";
            }

            portName = portName.Trim();

            // For standard TCP/IP ports, the port name is usually the IP address itself or hostname
            if (portName.StartsWith("IP_") || portName.StartsWith("TCP/IP"))
            {
                // Try to extract IP from port name (format is often IP_x.x.x.x or similar)
                string[] parts = portName.Split('_');
                if (parts.Length > 1)
                {
                    string potentialIP = parts[1];
                    if (IsValidIPAddress(potentialIP))
                    {
                        return potentialIP;
                    }
                }
            }

            string extractedIP = ExtractIPFromString(portName);
            if (!string.IsNullOrEmpty(extractedIP))
            {
                return extractedIP;
            }

            // For WSD (Web Services for Devices) ports
            if (portName.StartsWith("WSD-") || portName.Contains("WSD"))
            {
                // Try to get IP from WSD port mapping
                try
                {
                    if (lookupInfo.TcpIpPortHosts.ContainsKey(portName))
                    {
                        return lookupInfo.TcpIpPortHosts[portName];
                    }

                    // If the above method doesn't work, try to query the registry for WSD ports
                    // WSD printers often store their info in registry
                    string wsdIPAddress = GetWSDPrinterIPFromRegistry(portName);
                    if (!string.IsNullOrEmpty(wsdIPAddress))
                    {
                        return wsdIPAddress;
                    }

                    // If WSD printer but can't determine IP through normal means
                    // try to get it through parsing the port description which sometimes contains the IP
                    foreach (var item in lookupInfo.PortDescriptions)
                    {
                        if (item.Key.Contains(portName))
                        {
                            string ipFromDescription = ExtractIPFromString(item.Value);
                            if (!string.IsNullOrEmpty(ipFromDescription))
                            {
                                return ipFromDescription;
                            }
                        }
                    }

                    return "WSD (IP unknown)";
                }
                catch
                {
                    return "WSD (Error retrieving IP)";
                }
            }

            // For USB ports
            if (portName.StartsWith("USB") || portName.Contains("USB"))
            {
                return "USB Printer (No IP)";
            }

            if (IsRedirectedPrinterPort(portName))
            {
                return "Redirected printer (No IP)";
            }

            if (IsLocalOrVirtualPrinterPort(portName))
            {
                return "Local/virtual port (No IP)";
            }

            // For network printers with hostname
            if (CouldBePrinterHostName(portName))
            {
                try
                {
                    // Check if it's a hostname - try to resolve it to an IP
                    Task<IPAddress[]> dnsTask = Task.Run(() => Dns.GetHostAddresses(portName));
                    if (dnsTask.Wait(PrinterDnsTimeoutMs) && dnsTask.Result.Length > 0)
                    {
                        return dnsTask.Result[0].ToString();
                    }

                    return "DNS lookup timed out";
                }
                catch
                {
                    // If it's not a resolvable hostname, continue
                }
            }

            // If no IP could be determined
            return "Not available";
        }

        private static bool IsRedirectedPrinterPort(string portName)
        {
            return System.Text.RegularExpressions.Regex.IsMatch(portName ?? string.Empty, @"^TS\d+$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        }

        private static bool IsLocalOrVirtualPrinterPort(string portName)
        {
            if (string.IsNullOrWhiteSpace(portName))
            {
                return true;
            }

            string normalized = portName.Trim();
            return normalized.Equals("PORTPROMPT:", StringComparison.OrdinalIgnoreCase) ||
                   normalized.EndsWith(":", StringComparison.OrdinalIgnoreCase) ||
                   normalized.StartsWith("LPT", StringComparison.OrdinalIgnoreCase) ||
                   normalized.StartsWith("COM", StringComparison.OrdinalIgnoreCase) ||
                   normalized.StartsWith("FILE", StringComparison.OrdinalIgnoreCase) ||
                   normalized.IndexOf("PDF", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   normalized.IndexOf("XPS", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   normalized.IndexOf("OneNote", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   normalized.IndexOf("Evernote", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool CouldBePrinterHostName(string portName)
        {
            if (string.IsNullOrWhiteSpace(portName))
            {
                return false;
            }

            string normalized = portName.Trim();
            return normalized.All(c => char.IsLetterOrDigit(c) || c == '-' || c == '.') &&
                   normalized.Any(char.IsLetter);
        }

        // Helper to validate if a string is a valid IP address
        private static bool IsValidIPAddress(string ipAddress)
        {
            return IPAddress.TryParse(ipAddress, out _);
        }

        // Helper to extract IP from any string using regex
        private static string ExtractIPFromString(string input)
        {
            if (string.IsNullOrEmpty(input))
                return string.Empty;

            // Look for IPv4 pattern in the string
            var match = System.Text.RegularExpressions.Regex.Match(input,
                @"\b(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b");

            if (match.Success)
            {
                return match.Value;
            }

            return string.Empty;
        }

        // Helper method to try to get WSD printer IP from registry
        private static string GetWSDPrinterIPFromRegistry(string portName)
        {
            try
            {
                // WSD printer information is sometimes stored in the registry
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Control\Print\Monitors\WSD Port"))
                {
                    if (key != null)
                    {
                        foreach (string subKeyName in key.GetSubKeyNames())
                        {
                            using (RegistryKey portKey = key.OpenSubKey(subKeyName))
                            {
                                if (portKey != null)
                                {
                                    string currentPort = portKey.GetValue("PortName") as string;
                                    if (currentPort == portName)
                                    {
                                        // Try different possible value names for the IP address
                                        string[] possibleValueNames = { "IPAddress", "PrinterAddress", "Address", "HostAddress" };
                                        foreach (string valueName in possibleValueNames)
                                        {
                                            string ip = portKey.GetValue(valueName) as string;
                                            if (!string.IsNullOrEmpty(ip) && IsValidIPAddress(ip))
                                            {
                                                return ip;
                                            }
                                        }

                                        // If no direct IP address, check PrinterIP and strip any port number
                                        string printerIP = portKey.GetValue("PrinterIP") as string;
                                        if (!string.IsNullOrEmpty(printerIP))
                                        {
                                            // Often stored as IP:port (e.g., 192.168.1.100:631)
                                            int colonIndex = printerIP.IndexOf(':');
                                            if (colonIndex > 0)
                                            {
                                                printerIP = printerIP.Substring(0, colonIndex);
                                            }

                                            if (IsValidIPAddress(printerIP))
                                            {
                                                return printerIP;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch
            {
                // Ignore registry errors
            }

            return string.Empty;
        }
    }
}
