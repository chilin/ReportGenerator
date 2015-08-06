using System;
using System.Collections.Generic;
using System.Linq;
using log4net;

namespace ReportGeneratorApp.Log
{
    public class Logger
    {
        public static void Debug(string message)
        {
            if (message == null) return;
            var log = LogManager.GetLogger("ReportLog");
            if (log.IsDebugEnabled)
            {
                log.Debug(string.Format("Message   : {0}", message));
            }
            log = null;
        }

        public static void Info(string message)
        {
            if (message == null) return;
            var log = LogManager.GetLogger("ReportLog");
            if (log.IsInfoEnabled)
            {
                log.Info(string.Format("Message   : {0}", message));
            }
            log = null;
        }

        public static void Warn(string message)
        {
            if (message == null) return;
            var log = LogManager.GetLogger("ReportLog");
            if (log.IsWarnEnabled)
            {
                log.Warn(string.Format("Message   : {0}", message));
            }
            log = null;
        }

        public static void Error(string message)
        {
            if (message == null) return;
            var log = LogManager.GetLogger("ReportLog");
            if (log.IsErrorEnabled)
            {
                log.Error(string.Format("Message   : {0}", message));
            }
            log = null;
        }

        public static void Error(Exception exception)
        {
            if (exception == null) return;
            var log = LogManager.GetLogger("ReportLog");
            if (log.IsErrorEnabled)
            {
                log.Error(string.Format(@"Source    : {0}
Method    : {1}
Message   : {2}
Trace     : {3}",
                exception.Source ?? string.Empty,
                exception.TargetSite == null ? string.Empty : exception.TargetSite.Name,
                exception.Message,
                exception.StackTrace ?? string.Empty));
            }
            log = null;
        }

        public static void Fatal(string message)
        {
            if (message == null) return;
            var log = LogManager.GetLogger("ReportLog");
            if (log.IsFatalEnabled)
            {
                log.Fatal(string.Format("Message   : {0}", message));
            }
            log = null;
        }

        public static void Fatal(Exception exception)
        {
            if (exception == null) return;
            var log = LogManager.GetLogger("ReportLog");
            if (log.IsFatalEnabled)
            {
                log.Fatal(string.Format(@"Source    : {0}
Method    : {1}
Message   : {2}
Trace     : {3}",
                exception.Source ?? string.Empty,
                exception.TargetSite == null ? string.Empty : exception.TargetSite.Name,
                exception.Message,
                exception.StackTrace ?? string.Empty));
            }
            log = null;
        }
    }
}