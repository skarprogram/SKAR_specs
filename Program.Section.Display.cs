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
        private static async Task<DisplayReportData> CollectDisplayReportDataAsync()
        {
            return await Task.Run(async () => {
                DisplayReportData reportData = new DisplayReportData();

                try
                {
                    Task videoControllersTask = Task.Run(() => CollectVideoControllerInfo(reportData));
                    Task monitorsTask = Task.Run(() => CollectMonitorInfo(reportData));

                    Stopwatch displayWaitStopwatch = Stopwatch.StartNew();
                    await Task.WhenAll(videoControllersTask, monitorsTask);
                    DebugLog($"DISPLAY detail: combined display tasks completed in {displayWaitStopwatch.Elapsed.TotalSeconds:F2}s");
                }
                catch
                {
                    reportData.ErrorMessage = "Error retrieving display information";
                }

                return reportData;
            });
        }

        private static void CollectVideoControllerInfo(DisplayReportData reportData)
        {
            Stopwatch videoStopwatch = Stopwatch.StartNew();
            DebugLog("DISPLAY detail: START video controllers WMI");

            try
            {
                using (ManagementClass managementClass = new ManagementClass("Win32_VideoController"))
                using (ManagementObjectCollection collection = managementClass.GetInstances())
                {
                    foreach (ManagementObject obj in collection)
                    {
                        try
                        {
                            int horizontalRes = 0;
                            int verticalRes = 0;

                            if (obj["CurrentHorizontalResolution"] != null)
                            {
                                horizontalRes = Convert.ToInt32(obj["CurrentHorizontalResolution"]);
                            }

                            if (obj["CurrentVerticalResolution"] != null)
                            {
                                verticalRes = Convert.ToInt32(obj["CurrentVerticalResolution"]);
                            }

                            string memory = "Unknown";
                            if (obj["AdapterRAM"] != null)
                            {
                                long ram = Convert.ToInt64(obj["AdapterRAM"]);
                                memory = ram > 1073741824
                                    ? $"{ram / 1073741824.0:F2} GB"
                                    : $"{ram / 1048576.0:F0} MB";
                            }

                            reportData.VideoControllers.Add(new VideoControllerInfo
                            {
                                DeviceId = obj["DeviceID"]?.ToString() ?? "Unknown",
                                Name = obj["Caption"]?.ToString() ?? "Unknown",
                                Processor = obj["VideoProcessor"]?.ToString() ?? "Unknown",
                                Resolution = (horizontalRes > 0 && verticalRes > 0) ? $"{horizontalRes}x{verticalRes}" : "N/A",
                                Memory = memory,
                                DriverVersion = obj["DriverVersion"]?.ToString() ?? "Unknown"
                            });
                        }
                        catch
                        {
                            reportData.VideoControllers.Add(new VideoControllerInfo { ErrorMessage = "Error retrieving video controller information" });
                        }
                    }
                }
            }
            catch
            {
                reportData.VideoControllersErrorMessage = "Error accessing video controllers";
            }

            DebugLog($"DISPLAY detail: DONE video controllers WMI in {videoStopwatch.Elapsed.TotalSeconds:F2}s");
        }

        private static void CollectMonitorInfo(DisplayReportData reportData)
        {
            Stopwatch monitorStopwatch = Stopwatch.StartNew();
            DebugLog("DISPLAY detail: START monitor PnP/EDID lookup");

            try
            {
                string encodedScript = Convert.ToBase64String(Encoding.Unicode.GetBytes(GetMonitorPnPEdidScript()));
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = GetSystemExecutablePath(@"WindowsPowerShell\v1.0\powershell.exe"),
                    Arguments = $"-NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand {encodedScript}",
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                Stopwatch monitorProcessStopwatch = Stopwatch.StartNew();
                ProcessRunResult processResult = RunProcessWithTimeout(psi, ExternalCommandTimeoutMs, "Monitor PnP/EDID lookup");
                DebugLog($"DISPLAY detail: monitor PnP/EDID process completed in {monitorProcessStopwatch.Elapsed.TotalSeconds:F2}s");

                string output = processResult.Output;
                string errors = processResult.Error;

                if (!processResult.TimedOut && !string.IsNullOrWhiteSpace(output))
                {
                    ParseMonitorOutput(reportData, output);
                }
                else if (!string.IsNullOrWhiteSpace(errors))
                {
                    reportData.Monitors.Add(new MonitorInfo { ErrorMessage = "PowerShell error: " + errors });
                }
            }
            catch (Exception ex)
            {
                reportData.MonitorsErrorMessage = "Error accessing monitors: " + ex.Message;
            }

            DebugLog($"DISPLAY detail: DONE monitor PnP/EDID lookup in {monitorStopwatch.Elapsed.TotalSeconds:F2}s");
        }

        private static void ParseMonitorOutput(DisplayReportData reportData, string output)
        {
            try
            {
                string[] lines = output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                foreach (string line in lines)
                {
                    if (line.StartsWith("MONITOR|"))
                    {
                        string[] parts = line.Split('|');
                        if (parts.Length >= 4)
                        {
                            string manufacturer = parts[1];
                            string name = parts[2];
                            string serial = parts[3];
                            string serialSource = parts.Length >= 5 ? parts[4] : "EDID";

                            if (name.StartsWith("LEN "))
                            {
                                string[] nameParts = name.Split(' ');
                                if (nameParts.Length > 1)
                                {
                                    name = nameParts[1];
                                }
                            }

                            string make = MapManufacturerToName(manufacturer);

                            if (string.IsNullOrEmpty(name))
                            {
                                name = "Unknown Model";
                            }

                            reportData.Monitors.Add(new MonitorInfo
                            {
                                Vendor = make,
                                Model = name,
                                Serial = serial,
                                SerialSource = serialSource,
                                FriendlyName = $"{make} {name}: {serial}"
                            });
                        }
                    }
                    else if (line.StartsWith("ERROR|"))
                    {
                        reportData.Monitors.Add(new MonitorInfo { ErrorMessage = "PowerShell error: " + line.Substring("ERROR|".Length) });
                    }
                    else if (line.StartsWith("INFO|"))
                    {
                        string message = line.Substring("INFO|".Length);
                        reportData.MonitorDebugMessages.Add(message);
                        DebugLog($"DISPLAY detail: monitor {message}");
                    }
                }
            }
            catch (Exception ex)
            {
                reportData.Monitors.Add(new MonitorInfo { ErrorMessage = "Error parsing monitor data: " + ex.Message });
            }
        }

        private static string RenderDisplayInfoHtml(DisplayReportData displayData)
        {
            if (displayData == null || !string.IsNullOrWhiteSpace(displayData?.ErrorMessage))
            {
                return TableRowText("Display", "Error retrieving display information");
            }

            return TableRowHtml("Display Devices", RenderVideoControllersHtml(displayData) + "<br>" + RenderMonitorsHtml(displayData));
        }

        private static string RenderVideoControllersHtml(DisplayReportData displayData)
        {
            StringBuilder displayInfo = new StringBuilder();
            displayInfo.Append("<h3>Video Controllers</h3><table border='1' class='video-controllers-table' data-table-type='video-controllers'>");
            displayInfo.Append("<thead><tr><th>Name</th><th>Processor</th><th>Resolution</th><th>Memory</th><th>Driver Version</th></tr></thead><tbody>");

            if (!string.IsNullOrWhiteSpace(displayData.VideoControllersErrorMessage))
            {
                displayInfo.Append($"<tr><td colspan='5'>{HtmlText(displayData.VideoControllersErrorMessage)}</td></tr>");
            }
            else if (displayData.VideoControllers.Count == 0)
            {
                displayInfo.Append("<tr><td colspan='5'>No video controllers found</td></tr>");
            }
            else
            {
                foreach (VideoControllerInfo videoController in displayData.VideoControllers)
                {
                    if (!string.IsNullOrWhiteSpace(videoController.ErrorMessage))
                    {
                        displayInfo.Append($"<tr><td colspan='5'>{HtmlText(videoController.ErrorMessage)}</td></tr>");
                        continue;
                    }

                    displayInfo.Append($"<tr data-gpu-id=\"{HtmlAttr(videoController.DeviceId)}\" data-gpu-name=\"{HtmlAttr(videoController.Name)}\">");
                    displayInfo.Append($"<td data-field=\"name\">{HtmlText(videoController.Name)}</td>");
                    displayInfo.Append($"<td data-field=\"processor\">{HtmlText(videoController.Processor)}</td>");
                    displayInfo.Append($"<td data-field=\"resolution\">{HtmlText(videoController.Resolution)}</td>");
                    displayInfo.Append($"<td data-field=\"memory\">{HtmlText(videoController.Memory)}</td>");
                    displayInfo.Append($"<td data-field=\"driver-version\">{HtmlText(videoController.DriverVersion)}</td>");
                    displayInfo.Append("</tr>");
                }
            }

            displayInfo.Append("</tbody></table>");
            return displayInfo.ToString();
        }

        private static string RenderMonitorsHtml(DisplayReportData displayData)
        {
            StringBuilder monitorsInfo = new StringBuilder();
            monitorsInfo.Append("<h3>Monitors</h3><table border='1' class='monitors-table' data-table-type='monitors'>");
            monitorsInfo.Append("<thead><tr><th>Vendor</th><th>Model</th><th>Serial</th><th>Friendly Name</th></tr></thead><tbody>");

            if (!string.IsNullOrWhiteSpace(displayData.MonitorsErrorMessage))
            {
                monitorsInfo.Append($"<tr><td colspan='4'>{HtmlText(displayData.MonitorsErrorMessage)}</td></tr>");
            }
            else if (displayData.Monitors.Count == 0)
            {
                monitorsInfo.Append("<tr><td colspan='4'>No monitors found</td></tr>");
            }
            else
            {
                foreach (MonitorInfo monitor in displayData.Monitors)
                {
                    if (!string.IsNullOrWhiteSpace(monitor.ErrorMessage))
                    {
                        monitorsInfo.Append($"<tr><td colspan='4'>{HtmlText(monitor.ErrorMessage)}</td></tr>");
                        continue;
                    }

                    monitorsInfo.Append($"<tr data-monitor-vendor=\"{HtmlAttr(monitor.Vendor)}\" data-monitor-model=\"{HtmlAttr(monitor.Model)}\" data-monitor-serial-source=\"{HtmlAttr(monitor.SerialSource)}\">");
                    monitorsInfo.Append($"<td data-field=\"vendor\">{HtmlText(monitor.Vendor)}</td>");
                    monitorsInfo.Append($"<td data-field=\"model\">{HtmlText(monitor.Model)}</td>");
                    monitorsInfo.Append($"<td data-field=\"serial\">{HtmlText(monitor.Serial)}</td>");
                    monitorsInfo.Append($"<td data-field=\"friendly\">{HtmlText(monitor.FriendlyName)}</td>");
                    monitorsInfo.Append("</tr>");
                }
            }

            monitorsInfo.Append("</tbody></table>");
            return monitorsInfo.ToString();
        }

        private static string GetMonitorPnPEdidScript()
        {
            return @"
$ProgressPreference = 'SilentlyContinue'
$WarningPreference = 'SilentlyContinue'
$InformationPreference = 'SilentlyContinue'
$monitors = @()
$monitorRows = 0
$edidReads = 0
$edidMissing = 0
$parseErrors = 0
$displayIdSerials = 0
$overrideReads = 0

function Convert-EdidText([byte[]]$bytes) {
    if (-not $bytes) { return '' }

    $textBytes = New-Object System.Collections.Generic.List[byte]
    foreach ($value in $bytes) {
        # EDID strings end at NUL, LF, or CR. Do not remove bytes from the
        # middle of a string because that can accidentally join two values.
        if ($value -eq 0 -or $value -eq 10 -or $value -eq 13) { break }
        $textBytes.Add($value)
    }

    if ($textBytes.Count -eq 0) { return '' }
    return [System.Text.Encoding]::ASCII.GetString($textBytes.ToArray()).Trim()
}

function Convert-WmiText($values) {
    if (-not $values) { return '' }
    $bytes = New-Object System.Collections.Generic.List[byte]
    foreach ($value in $values) {
        if ($value -eq 0) { break }
        if ($value -le 255) { $bytes.Add([byte]$value) }
    }
    if ($bytes.Count -eq 0) { return '' }
    return [System.Text.Encoding]::ASCII.GetString($bytes.ToArray()).Trim()
}

function Get-CompleteEdid([string]$deviceParametersPath) {
    [byte[]]$baseEdid = (Get-ItemProperty -Path $deviceParametersPath -Name EDID -ErrorAction SilentlyContinue).EDID
    $blocks = New-Object System.Collections.Generic.List[byte]

    if ($baseEdid) {
        $blocks.AddRange($baseEdid)
    }

    # Driver INF files can replace individual 128-byte EDID blocks. Apply
    # those replacements over the normal EDID, as Windows does.
    $overridePath = Join-Path $deviceParametersPath 'EDID_OVERRIDE'
    $overrideValues = Get-ItemProperty -Path $overridePath -ErrorAction SilentlyContinue
    if ($overrideValues) {
        $properties = @($overrideValues.PSObject.Properties |
            Where-Object { $_.Name -match '^\d+$' -and $_.Value -is [byte[]] } |
            Sort-Object { [int]$_.Name })

        foreach ($property in $properties) {
            $blockIndex = [int]$property.Name
            [byte[]]$overrideBlock = $property.Value
            if ($overrideBlock.Length -lt 128) { continue }

            $requiredLength = ($blockIndex + 1) * 128
            while ($blocks.Count -lt $requiredLength) { $blocks.Add(0) }
            for ($j = 0; $j -lt 128; $j++) {
                $blocks[$blockIndex * 128 + $j] = $overrideBlock[$j]
            }
            $script:overrideReads++
        }
    }

    return $blocks.ToArray()
}

function Get-DisplayIdSerial([byte[]]$edid) {
    if (-not $edid -or $edid.Length -lt 256) { return '' }

    $availableExtensions = [Math]::Floor($edid.Length / 128) - 1
    $extensionCount = [Math]::Min([int]$edid[126], [int]$availableExtensions)

    for ($block = 1; $block -le $extensionCount; $block++) {
        $blockOffset = $block * 128

        # 0x70 is a DisplayID extension block. Its data blocks begin at byte
        # five and use tag/revision/length followed by a variable payload.
        if ($edid[$blockOffset] -ne 0x70) { continue }

        $payloadLength = [Math]::Min([int]$edid[$blockOffset + 2], 121)
        $position = $blockOffset + 5
        $end = [Math]::Min($position + $payloadLength, $blockOffset + 127)

        while ($position + 3 -le $end) {
            $tag = $edid[$position]
            $length = [int]$edid[$position + 2]
            $dataStart = $position + 3
            $dataEnd = $dataStart + $length
            if ($length -eq 0 -or $dataEnd -gt $end) { break }

            # DisplayID 1.x tag 0x0A is the variable-length Product Serial
            # Number Data Block. Unlike base EDID descriptor 0xFF, it is not
            # restricted to 13 characters.
            if ($tag -eq 0x0A) {
                [byte[]]$serialBytes = $edid[$dataStart..($dataEnd - 1)]
                $value = Convert-EdidText $serialBytes
                if ($value) { return $value }
            }

            $position = $dataEnd
        }
    }

    return ''
}

function Clean-MonitorField([string]$value) {
    if ($null -eq $value) { return '' }
    return ($value -replace '[|\r\n]', ' ').Trim()
}

try {
    $monitors = @(Get-CimInstance -ClassName Win32_PnPEntity -ErrorAction Stop | Where-Object { $_.Service -eq 'monitor' })
    Write-Output ('INFO|Win32_PnPEntity Service=monitor returned ' + $monitors.Count + ' devices')
} catch {
    Write-Output ('ERROR|Win32_PnPEntity monitor query failed: ' + $_.Exception.Message)
    exit 0
}

$wmiMonitorIds = @()
try {
    $wmiMonitorIds = @(Get-CimInstance -Namespace 'root\wmi' -ClassName WmiMonitorID -ErrorAction Stop)
} catch {
    Write-Output ('INFO|WmiMonitorID fallback unavailable: ' + $_.Exception.Message)
}

foreach ($monitor in $monitors) {
    $deviceID = $monitor.DeviceID

    try {
        $edidPath = 'HKLM:\SYSTEM\CurrentControlSet\Enum\' + $deviceID + '\Device Parameters'
        [byte[]]$edid = Get-CompleteEdid $edidPath

        if ($edid -and $edid.Length -ge 128 -and
            $edid[0] -eq 0 -and $edid[1] -eq 0xFF -and $edid[7] -eq 0) {
            $edidReads++

            # Parse manufacturer ID (bytes 8-9) - 3 letter code compressed into 2 bytes
            $mfgID = [BitConverter]::ToUInt16($edid, 8)
            $mfgID = (($mfgID -band 0xFF) -shl 8) -bor (($mfgID -shr 8) -band 0xFF)
            $char1 = [char](64 + (($mfgID -shr 10) -band 0x1F))
            $char2 = [char](64 + (($mfgID -shr 5) -band 0x1F))
            $char3 = [char](64 + ($mfgID -band 0x1F))
            $manufacturer = [string]$char1 + $char2 + $char3

            $serialNum = Get-DisplayIdSerial $edid
            $serialSource = ''
            if ($serialNum) {
                $serialSource = 'DisplayID product serial'
                $displayIdSerials++
            }

            # Extract model name and serial from EDID descriptor blocks
            $name = ''
            for ($i = 54; $i -lt 126; $i += 18) {
                if ($edid[$i] -eq 0 -and $edid[$i+1] -eq 0 -and $edid[$i+2] -eq 0) {
                    $descriptorType = $edid[$i+3]

                    # 0xFC = Monitor name
                    if ($descriptorType -eq 0xFC) {
                        [byte[]]$nameBytes = $edid[($i+5)..($i+17)]
                        $name = Convert-EdidText $nameBytes
                    }

                    # 0xFF = base EDID serial string (maximum 13 characters).
                    if ($descriptorType -eq 0xFF -and -not $serialNum) {
                        [byte[]]$serialBytes = $edid[($i+5)..($i+17)]
                        $serialNum = Convert-EdidText $serialBytes
                        if ($serialNum) {
                            $serialSource = 'Base EDID text'
                        }
                    }
                }
            }

            # WmiMonitorID normally mirrors base EDID, but some display drivers
            # expose a useful value here when the registry descriptor is absent.
            if (-not $serialNum) {
                $wmiMatch = $wmiMonitorIds | Where-Object {
                    $_.InstanceName -and $_.InstanceName.StartsWith($deviceID, [System.StringComparison]::OrdinalIgnoreCase)
                } | Select-Object -First 1
                if ($wmiMatch) {
                    $serialNum = Convert-WmiText $wmiMatch.SerialNumberID
                    if ($serialNum) { $serialSource = 'WmiMonitorID' }
                    if (-not $name) { $name = Convert-WmiText $wmiMatch.UserFriendlyName }
                }
            }

            # The four-byte binary value is a separate EDID identifier. Keep it
            # as the final fallback and label it so it is not confused with an
            # ASCII serial printed on the monitor.
            if (-not $serialNum) {
                $numericSerial = [System.BitConverter]::ToUInt32($edid, 12)
                if ($numericSerial -ne 0) {
                    $serialNum = $numericSerial.ToString()
                    $serialSource = 'Base EDID numeric'
                } else {
                    $serialNum = 'N/A'
                    $serialSource = 'Not provided by monitor'
                }
            }

            if (-not $name) {
                $name = 'Unknown Model'
            }

            Write-Output ('MONITOR|' + (Clean-MonitorField $manufacturer) + '|' +
                (Clean-MonitorField $name) + '|' + (Clean-MonitorField $serialNum) + '|' +
                (Clean-MonitorField $serialSource))
            $monitorRows++
        } else {
            $edidMissing++
        }
    } catch {
        $parseErrors++
        # Skip monitors we can't read
    }
}

Write-Output ('INFO|EDID reads: ' + $edidReads + ', EDID override blocks: ' + $overrideReads +
    ', DisplayID serials: ' + $displayIdSerials + ', missing/invalid EDID: ' + $edidMissing +
    ', parse errors: ' + $parseErrors + ', monitor rows: ' + $monitorRows)
";
        }

        private static string MapManufacturerToName(string manufacturerCode)
        {
            switch (manufacturerCode)
            {
                case "LEN": return "Lenovo";
                case "ACI": return "ASUS";
                case "LGD": return "LG";
                case "SDC": return "Surface Display";
                case "SEC": return "Epson";
                case "SAM": return "Samsung";
                case "SNY": return "Sony";
                case "GSM": return "LG (Goldstar) TV";
                case "GWY": return "Gateway 2000";
                case "ITE": return "Integrated Tech Express";
                case "DEL": return "Dell";
                case "HPN": return "HP";
                case "AOC": return "AOC";
                case "BNQ": return "BenQ";
                case "MSI": return "MSI";
                case "PHL": return "Philips";
                case "VSC": return "ViewSonic";
                case "ACR": return "Acer";
                default: return "Unknown: " + manufacturerCode;
            }
        }
    }
}
