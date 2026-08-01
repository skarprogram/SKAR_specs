using System.Collections.Generic;
// SPDX-License-Identifier: GPL-3.0-only

using System.Diagnostics;
using System.Text;
using System.Threading.Tasks;

namespace SKAR_specs
{
    partial class Program
    {
        private static string officeLicenseName = "";
        private const int ExternalCommandTimeoutMs = 10000;
        private const int QuickExternalCommandTimeoutMs = 5000;
        private const int ProcessStreamDrainTimeoutMs = 10000;
        private const int NetworkTimeoutMs = 3000;
        private const int PrinterDnsTimeoutMs = 1000;
        private const string ExternalIpSummaryToken = "__SKAR_EXTERNAL_IP_SUMMARY__";
        private const string HardDriveSummaryToken = "__SKAR_HARD_DRIVE_SUMMARY__";
        private const uint AttachParentProcess = 0xFFFFFFFF;
        private static readonly object DebugLogSync = new object();
        private static readonly StringBuilder DebugLogBuilder = new StringBuilder();
        private static readonly Stopwatch AppStopwatch = Stopwatch.StartNew();
        private static readonly object ExternalIpLookupSync = new object();
        private static Task<List<ExternalIpServiceResult>> externalIpLookupTask;
        private static bool debugEnabled = false;
        private static bool debugConsoleAttached = false;
        private static int completedSections = 0;
    }
}
