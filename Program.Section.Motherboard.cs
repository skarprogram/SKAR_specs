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
        private static async Task<MotherboardReportData> CollectMotherboardReportDataAsync()
        {
            return await Task.Run(() => {
                MotherboardReportData reportData = new MotherboardReportData();

                Task baseBoardTask = Task.Run(() => CollectMotherboardBaseBoardInfo(reportData));
                Task biosTask = Task.Run(() => CollectMotherboardBiosInfo(reportData));
                Task detailedInfoTask = Task.Run(() => CollectMotherboardDetailInfo(reportData));

                Task.WaitAll(baseBoardTask, biosTask, detailedInfoTask);
                return reportData;
            });
        }

        private static void CollectMotherboardBaseBoardInfo(MotherboardReportData reportData)
        {
            try
            {
                using (ManagementClass managementClass = new ManagementClass("Win32_BaseBoard"))
                using (ManagementObjectCollection collection = managementClass.GetInstances())
                {
                    foreach (ManagementObject obj in collection)
                    {
                        if (obj["Manufacturer"] != null && !string.IsNullOrWhiteSpace(obj["Manufacturer"].ToString()))
                        {
                            reportData.Manufacturer = obj["Manufacturer"].ToString();
                        }

                        if (obj["Product"] != null && !string.IsNullOrWhiteSpace(obj["Product"].ToString()))
                        {
                            reportData.Model = obj["Product"].ToString();
                        }

                        break;
                    }
                }
            }
            catch
            {
                reportData.BaseBoardErrorMessage = "Error retrieving information";
            }
        }

        private static void CollectMotherboardBiosInfo(MotherboardReportData reportData)
        {
            try
            {
                using (ManagementClass managementClass = new ManagementClass("Win32_BIOS"))
                using (ManagementObjectCollection collection = managementClass.GetInstances())
                {
                    foreach (ManagementObject obj in collection)
                    {
                        if (obj["SMBIOSBIOSVersion"] != null)
                        {
                            reportData.BiosVersion = obj["SMBIOSBIOSVersion"].ToString();
                        }

                        break;
                    }
                }
            }
            catch
            {
                reportData.BiosErrorMessage = "Error retrieving information";
            }
        }

        private static void CollectMotherboardDetailInfo(MotherboardReportData reportData)
        {
            try
            {
                using (ManagementClass managementClass = new ManagementClass("Win32_BaseBoard"))
                using (ManagementObjectCollection collection = managementClass.GetInstances())
                {
                    foreach (ManagementObject obj in collection)
                    {
                        MotherboardDetailInfo detailInfo = new MotherboardDetailInfo();

                        if (obj["Manufacturer"] != null)
                        {
                            detailInfo.Manufacturer = obj["Manufacturer"].ToString();
                        }

                        if (obj["Model"] != null)
                        {
                            detailInfo.Model = obj["Model"].ToString();
                        }

                        if (obj["Product"] != null)
                        {
                            detailInfo.Product = obj["Product"].ToString();
                        }

                        if (obj["SerialNumber"] != null)
                        {
                            detailInfo.SerialNumber = obj["SerialNumber"].ToString();
                        }

                        if (obj["Version"] != null)
                        {
                            detailInfo.Version = obj["Version"].ToString();
                        }

                        reportData.Detail = detailInfo;
                        break;
                    }
                }
            }
            catch
            {
                reportData.DetailErrorMessage = "Error retrieving detailed motherboard information";
            }
        }

        private static string RenderMotherboardInfoHtml(MotherboardReportData motherboardData)
        {
            StringBuilder result = new StringBuilder();

            if (!string.IsNullOrWhiteSpace(motherboardData.BaseBoardErrorMessage))
            {
                result.Append(TableRowText("Motherboard Manufacturer:", motherboardData.BaseBoardErrorMessage));
                result.Append(TableRowText("Motherboard Model:", motherboardData.BaseBoardErrorMessage));
            }
            else
            {
                result.Append(TableRowText("Motherboard Manufacturer:", motherboardData.Manufacturer));
                result.Append(TableRowText("Motherboard Model:", motherboardData.Model));
            }

            result.Append(TableRowText("BIOS/UEFI Ver.:",
                string.IsNullOrWhiteSpace(motherboardData.BiosErrorMessage) ? motherboardData.BiosVersion : motherboardData.BiosErrorMessage));
            result.Append(RenderMotherboardDetailHtml(motherboardData));

            return result.ToString();
        }

        private static string RenderMotherboardDetailHtml(MotherboardReportData motherboardData)
        {
            if (!string.IsNullOrWhiteSpace(motherboardData.DetailErrorMessage))
            {
                return TableRowText("WMI Motherboard:", motherboardData.DetailErrorMessage);
            }

            StringBuilder motherboardInfo = new StringBuilder();
            motherboardInfo.Append("<table border='1'>");

            if (motherboardData.Detail == null)
            {
                motherboardInfo.Append("<tr><td colspan='2'>No detailed motherboard information available</td></tr>");
            }
            else
            {
                motherboardInfo.Append($"<tr><td>Manufacturer</td><td>{HtmlText(motherboardData.Detail.Manufacturer)}</td></tr>");
                motherboardInfo.Append($"<tr><td>Model</td><td>{HtmlText(motherboardData.Detail.Model)}</td></tr>");
                motherboardInfo.Append($"<tr><td>Product</td><td>{HtmlText(motherboardData.Detail.Product)}</td></tr>");
                motherboardInfo.Append($"<tr><td>SerialNumber</td><td>{HtmlText(motherboardData.Detail.SerialNumber)}</td></tr>");
                motherboardInfo.Append($"<tr><td>Version</td><td>{HtmlText(motherboardData.Detail.Version)}</td></tr>");
            }

            motherboardInfo.Append("</table>");
            return TableRowHtml("WMI Motherboard:", motherboardInfo.ToString());
        }
    }
}
