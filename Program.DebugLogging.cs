// SPDX-License-Identifier: GPL-3.0-only

using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace SKAR_specs
{
    partial class Program
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool AttachConsole(uint dwProcessId);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool FreeConsole();

        private static void InitializeDebugMode(string[] args)
        {
            debugEnabled = args != null && args.Any(a =>
                string.Equals(a, "-debug", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(a, "--debug", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(a, "/debug", StringComparison.OrdinalIgnoreCase));

            if (!debugEnabled)
            {
                return;
            }

            try
            {
                debugConsoleAttached = AttachConsole(AttachParentProcess);
                if (debugConsoleAttached)
                {
                    Console.OutputEncoding = Encoding.UTF8;
                }
            }
            catch
            {
                debugConsoleAttached = false;
                // Debug logging still goes into the report if console attachment fails.
            }

            DebugLog("Debug mode enabled. Section timings will be printed as collection runs.");
        }

        private static void DebugLog(string message)
        {
            if (!debugEnabled)
            {
                return;
            }

            lock (DebugLogSync)
            {
                string line = $"{DateTime.Now:HH:mm:ss.fff} +{AppStopwatch.Elapsed:mm\\:ss\\.fff} {message}";
                DebugLogBuilder.AppendLine(line);
                try
                {
                    if (debugConsoleAttached)
                    {
                        Console.WriteLine(line);
                    }
                }
                catch
                {
                    // Keep report logging even if console output is unavailable.
                }
            }
        }
    }
}
