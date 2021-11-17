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
        private readonly IServiceProvider serviceProvider;

        public Logger(IServiceProvider serviceProvider) => this.serviceProvider = serviceProvider;

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
            var log = serviceProvider.GetService(typeof(SVsActivityLog)) as IVsActivityLog;
            if (log == null)
            {
                return;
            }

            var hr = log.LogEntry((uint)ToEntryType(logType), source, message);
        }

        public async System.Threading.Tasks.Task LogAsync(string message, string source, Exception e)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            var format = "Message: {0} \n Exception Message: {1} \n Stack Trace: {2}";
            var log = serviceProvider.GetService(typeof(SVsActivityLog)) as IVsActivityLog;
            if (log == null)
            {
                return;
            }

            var hr = log.LogEntry((uint)__ACTIVITYLOG_ENTRYTYPE.ALE_ERROR, source, string.Format(CultureInfo.CurrentCulture, format, message, e.Message, e.StackTrace));
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

            var uiShell = (IVsUIShell)serviceProvider.GetService(typeof(SVsUIShell));
            if (uiShell != null)
            {
                var clsid = Guid.Empty;
                uiShell.ShowMessageBox(0, ref clsid, title, message, string.Empty,
                    0, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST,
                    icon, 0, out var result);
            }
        }

        private __ACTIVITYLOG_ENTRYTYPE ToEntryType(LogType logType)
        {
            switch (logType)
            {
                case LogType.Information:
                    return __ACTIVITYLOG_ENTRYTYPE.ALE_INFORMATION;

                case LogType.Warning:
                    return __ACTIVITYLOG_ENTRYTYPE.ALE_WARNING;

                case LogType.Error:
                    return __ACTIVITYLOG_ENTRYTYPE.ALE_ERROR;

                default:
                    return __ACTIVITYLOG_ENTRYTYPE.ALE_INFORMATION;
            }
        }
    }
}
