using System;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Threading;

namespace Microsoft.SnippetDesigner
{
    public enum LogType
    {
        Information,
        Warning,
        Error
    }

    public class Logger : ILogger
    {
        private readonly IServiceProvider _serviceProvider;

        public Logger(IServiceProvider serviceProvider) => _serviceProvider = serviceProvider;

        public void Log(string message, string source, LogType logType) => ThreadHelper.JoinableTaskFactory.Run(async delegate
         {
             await LogAsync(message, source, logType);
         });

        public void Log(string message, string source, Exception e) => ThreadHelper.JoinableTaskFactory.Run(async delegate
         {
             await LogAsync(message, source, e);
         });

        public async System.Threading.Tasks.Task LogAsync(string message, string source, LogType logType)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            if (_serviceProvider.GetService(typeof(SVsActivityLog)) is not IVsActivityLog log)
            {
                return;
            }

            _ = log.LogEntry((uint)ToEntryType(logType), source, message);
        }

        public async System.Threading.Tasks.Task LogAsync(string message, string source, Exception e)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            const string format = "Message: {0} \n Exception Message: {1} \n Stack Trace: {2}";
            if (_serviceProvider.GetService(typeof(SVsActivityLog)) is not IVsActivityLog log)
            {
                return;
            }

            _ = log.LogEntry((uint)__ACTIVITYLOG_ENTRYTYPE.ALE_ERROR, source,
                string.Format(CultureInfo.CurrentCulture, format, message, e.Message, e.StackTrace));
        }

        public void MessageBox(string title, string message, LogType logType) => ThreadHelper.JoinableTaskFactory.Run(async delegate
         {
             await MessageBoxAsync(title, message, logType);
         });

        public async System.Threading.Tasks.Task MessageBoxAsync(string title, string message, LogType logType)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            var icon = OLEMSGICON.OLEMSGICON_INFO;
            if (logType == LogType.Error)
            {
                icon = OLEMSGICON.OLEMSGICON_CRITICAL;
            }
            else if (logType == LogType.Warning)
            {
                icon = OLEMSGICON.OLEMSGICON_WARNING;
            }

            var uiShell = (IVsUIShell)_serviceProvider.GetService(typeof(SVsUIShell));
            if (uiShell != null)
            {
                var clsid = Guid.Empty;
                uiShell.ShowMessageBox(0, ref clsid, title, message, string.Empty,
                    0, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST,
                    icon, 0, out var result);
            }
        }

        private static __ACTIVITYLOG_ENTRYTYPE ToEntryType(LogType logType) => logType switch
        {
            LogType.Information => __ACTIVITYLOG_ENTRYTYPE.ALE_INFORMATION,
            LogType.Warning => __ACTIVITYLOG_ENTRYTYPE.ALE_WARNING,
            LogType.Error => __ACTIVITYLOG_ENTRYTYPE.ALE_ERROR,
            _ => __ACTIVITYLOG_ENTRYTYPE.ALE_INFORMATION,
        };
    }
}
