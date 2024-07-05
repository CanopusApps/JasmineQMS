using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Practices.EnterpriseLibrary.Logging;
namespace TEPL.QMS.Common
{
    public static class LoggerBlock
    {
        public static void WriteTraceLog(Exception ex)
        {
            LogEntry logEntryObject = new LogEntry();
            logEntryObject.EventId = 1400;
            logEntryObject.Priority = 1;
            logEntryObject.Severity = TraceEventType.Error;
            logEntryObject.Message = ex.Message;
            //Logger.SetLogWriter(new LogWriterFactory().Create());

            logEntryObject.Categories.Add("Error");

            Logger.Write(logEntryObject);
        }
        public static void WriteLog(string strMessage)
        {
            LogEntry logEntry = new LogEntry();
            logEntry.EventId = 100;
            logEntry.Priority = 2;
            logEntry.Severity = TraceEventType.Information;
            logEntry.Message = strMessage;
            //Logger.SetLogWriter(new LogWriterFactory().Create());
            Logger.Write(logEntry);
        }
        public static void LogInformation(int eventId, DateTime timeStamp, string message, string documentName)
        {
            LogEntry logEntryObject = new LogEntry();
            logEntryObject.EventId = eventId;
            logEntryObject.Title = documentName;
            logEntryObject.TimeStamp = timeStamp;
            logEntryObject.MachineName = System.Environment.MachineName;
            logEntryObject.Severity = TraceEventType.Information;
            logEntryObject.Message = message;
            logEntryObject.Categories.Add("Exception");
            //Logger.SetLogWriter(new LogWriterFactory().Create());
            Logger.Write(logEntryObject);
        }
    }
}
