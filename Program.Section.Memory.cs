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
        private static async Task<MemoryReportData> CollectMemoryReportDataAsync()
        {
            return await Task.Run(() => {
                MemoryReportData reportData = new MemoryReportData();

                try
                {
                    Task memoryArrayTask = Task.Run(() => CollectMemoryArrayInfo(reportData));
                    Task physicalMemoryTask = Task.Run(() => CollectPhysicalMemoryModuleInfo(reportData));

                    Task.WaitAll(memoryArrayTask, physicalMemoryTask);
                    reportData.TotalSystemMemory = CollectTotalSystemMemory();
                }
                catch
                {
                    reportData.ErrorMessage = "Error retrieving memory information";
                }

                return reportData;
            });
        }

        private static void CollectMemoryArrayInfo(MemoryReportData reportData)
        {
            try
            {
                using (ManagementClass managementClass = new ManagementClass("Win32_PhysicalMemoryArray"))
                using (ManagementObjectCollection collection = managementClass.GetInstances())
                {
                    foreach (ManagementObject obj in collection)
                    {
                        try
                        {
                            MemoryArrayInfo arrayInfo = new MemoryArrayInfo();

                            if (obj["MaxCapacity"] != null)
                            {
                                arrayInfo.MaxCapacityGB = Convert.ToDouble(obj["MaxCapacity"]) / (1024 * 1024);
                            }

                            if (obj["MemoryDevices"] != null)
                            {
                                arrayInfo.MemoryDevices = obj["MemoryDevices"].ToString();
                            }

                            reportData.MemoryArrays.Add(arrayInfo);
                        }
                        catch
                        {
                            reportData.MemoryArrays.Add(new MemoryArrayInfo { ErrorMessage = "Error retrieving data" });
                        }
                    }
                }
            }
            catch
            {
                reportData.MemoryArrayErrorMessage = "Error accessing memory array";
            }
        }

        private static void CollectPhysicalMemoryModuleInfo(MemoryReportData reportData)
        {
            try
            {
                using (ManagementClass managementClass = new ManagementClass("Win32_PhysicalMemory"))
                using (ManagementObjectCollection collection = managementClass.GetInstances())
                {
                    foreach (ManagementObject obj in collection)
                    {
                        try
                        {
                            PhysicalMemoryModuleInfo moduleInfo = new PhysicalMemoryModuleInfo();

                            if (obj["Capacity"] != null)
                            {
                                moduleInfo.CapacityGB = Convert.ToDouble(obj["Capacity"]) / (1024 * 1024 * 1024);
                            }

                            if (obj["Manufacturer"] != null)
                            {
                                moduleInfo.Manufacturer = obj["Manufacturer"].ToString();
                            }

                            if (obj["BankLabel"] != null)
                            {
                                moduleInfo.BankLabel = obj["BankLabel"].ToString();
                            }

                            if (obj["ConfiguredClockSpeed"] != null)
                            {
                                moduleInfo.ClockSpeed = $"{obj["ConfiguredClockSpeed"]} MHz";
                            }

                            if (obj["DeviceLocator"] != null)
                            {
                                moduleInfo.DeviceLocator = obj["DeviceLocator"].ToString();
                            }

                            if (obj["SerialNumber"] != null)
                            {
                                moduleInfo.SerialNumber = obj["SerialNumber"].ToString();
                            }

                            reportData.PhysicalMemoryModules.Add(moduleInfo);
                        }
                        catch
                        {
                            reportData.PhysicalMemoryModules.Add(new PhysicalMemoryModuleInfo { ErrorMessage = "Error retrieving memory module data" });
                        }
                    }
                }
            }
            catch
            {
                reportData.PhysicalMemoryErrorMessage = "Error accessing memory modules";
            }
        }

        private static string CollectTotalSystemMemory()
        {
            try
            {
                using (ManagementClass managementClass = new ManagementClass("Win32_ComputerSystem"))
                using (ManagementObjectCollection collection = managementClass.GetInstances())
                {
                    foreach (ManagementObject obj in collection)
                    {
                        if (obj["TotalPhysicalMemory"] != null)
                        {
                            double totalRam = Convert.ToDouble(obj["TotalPhysicalMemory"]) / 1073741824;
                            return $"{totalRam:F2} GB";
                        }

                        break;
                    }
                }
            }
            catch
            {
                return "Error retrieving total memory";
            }

            return "Unknown";
        }

        private static string RenderMemoryInfoHtml(MemoryReportData memoryData)
        {
            if (memoryData == null || !string.IsNullOrWhiteSpace(memoryData?.ErrorMessage))
            {
                return TableRowText("RAM", "Error retrieving memory information");
            }

            StringBuilder completeInfo = new StringBuilder();
            completeInfo.Append(TableRowText("Total System Memory", memoryData.TotalSystemMemory));
            completeInfo.Append(TableRowHtml("Memory Details", RenderMemoryArrayHtml(memoryData) + "<br>" + RenderPhysicalMemoryModulesHtml(memoryData)));

            return completeInfo.ToString();
        }

        private static string RenderMemoryArrayHtml(MemoryReportData memoryData)
        {
            StringBuilder arrayInfo = new StringBuilder();
            arrayInfo.Append("<h3>Physical Memory Array</h3><table border='1' class='memory-array-table' data-table-type='memory-array'>");
            arrayInfo.Append("<thead><tr><th>MaxCapacity</th><th>MemoryDevices</th></tr></thead><tbody>");

            if (!string.IsNullOrWhiteSpace(memoryData.MemoryArrayErrorMessage))
            {
                arrayInfo.Append($"<tr><td colspan='2'>{HtmlText(memoryData.MemoryArrayErrorMessage)}</td></tr>");
            }
            else if (memoryData.MemoryArrays.Count == 0)
            {
                arrayInfo.Append("<tr><td colspan='2'>No memory array information available</td></tr>");
            }
            else
            {
                foreach (MemoryArrayInfo memoryArray in memoryData.MemoryArrays)
                {
                    if (!string.IsNullOrWhiteSpace(memoryArray.ErrorMessage))
                    {
                        arrayInfo.Append($"<tr><td colspan='2'>{HtmlText(memoryArray.ErrorMessage)}</td></tr>");
                        continue;
                    }

                    arrayInfo.Append($"<tr data-memory-array-item><td data-field=\"capacity\">{memoryArray.MaxCapacityGB:F2} GB</td><td data-field=\"devices\">{HtmlText(memoryArray.MemoryDevices)}</td></tr>");
                }
            }

            arrayInfo.Append("</tbody></table>");
            return arrayInfo.ToString();
        }

        private static string RenderPhysicalMemoryModulesHtml(MemoryReportData memoryData)
        {
            StringBuilder moduleInfo = new StringBuilder();
            moduleInfo.Append("<h3>Physical Memory</h3><table border='1' class='memory-modules-table' data-table-type='memory-modules'>");
            moduleInfo.Append("<thead><tr><th>Manufacturer</th><th>BankLabel</th><th>ClockSpeed</th><th>DeviceLocator</th><th>Capacity</th><th>SerialNumber</th></tr></thead><tbody>");

            if (!string.IsNullOrWhiteSpace(memoryData.PhysicalMemoryErrorMessage))
            {
                moduleInfo.Append($"<tr><td colspan='6'>{HtmlText(memoryData.PhysicalMemoryErrorMessage)}</td></tr>");
            }
            else if (memoryData.PhysicalMemoryModules.Count == 0)
            {
                moduleInfo.Append("<tr><td colspan='6'>No physical memory modules detected</td></tr>");
            }
            else
            {
                foreach (PhysicalMemoryModuleInfo module in memoryData.PhysicalMemoryModules)
                {
                    if (!string.IsNullOrWhiteSpace(module.ErrorMessage))
                    {
                        moduleInfo.Append($"<tr><td colspan='6'>{HtmlText(module.ErrorMessage)}</td></tr>");
                        continue;
                    }

                    moduleInfo.Append($"<tr data-memory-module=\"{HtmlAttr(module.DeviceLocator)}\"><td data-field=\"manufacturer\">{HtmlText(module.Manufacturer)}</td><td data-field=\"bank\">{HtmlText(module.BankLabel)}</td><td data-field=\"speed\">{HtmlText(module.ClockSpeed)}</td><td data-field=\"location\">{HtmlText(module.DeviceLocator)}</td><td data-field=\"capacity\">{module.CapacityGB:F2} GB</td><td data-field=\"serial\">{HtmlText(module.SerialNumber)}</td></tr>");
                }
            }

            moduleInfo.Append("</tbody></table>");
            return moduleInfo.ToString();
        }
    }
}
