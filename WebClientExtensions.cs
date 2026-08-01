// SPDX-License-Identifier: GPL-3.0-only

using System;
using System.Net;
using System.Threading.Tasks;

namespace SKAR_specs
{
    public static class WebClientExtensions
    {
        public static async Task<string> DownloadStringWithTimeoutAsync(this WebClient client, string address, int timeoutMilliseconds = 3000)
        {
            var tcs = new TaskCompletionSource<string>();
            DownloadStringCompletedEventHandler completedHandler = null;
            completedHandler = (s, e) => {
                client.DownloadStringCompleted -= completedHandler;
                if (e.Error != null)
                    tcs.TrySetException(e.Error);
                else if (e.Cancelled)
                    tcs.TrySetCanceled();
                else
                    tcs.TrySetResult(e.Result);
            };
            client.DownloadStringCompleted += completedHandler;
            client.DownloadStringAsync(new Uri(address));
            var timeoutTask = Task.Delay(timeoutMilliseconds);
            var completedTask = await Task.WhenAny(tcs.Task, timeoutTask);
            if (completedTask == timeoutTask)
            {
                client.DownloadStringCompleted -= completedHandler;
                client.CancelAsync();
                throw new TimeoutException($"Download from {address} timed out after {timeoutMilliseconds}ms");
            }
            return await tcs.Task;
        }
    }
}
