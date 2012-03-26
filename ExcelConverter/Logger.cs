using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;

namespace ExcelConverter
{
    public class Logger
    {

        public static void Write(string message, EventLogEntryType entryType)
        {
            string sSource;
            string sLog;

            sSource = "Excel Converter Service";
            sLog = "Service";

            if (!EventLog.SourceExists(sSource))
                EventLog.CreateEventSource(sSource, sLog);
            EventLog.WriteEntry(sSource, message, entryType);
        }

        public static void Write(string message)
        {
            Write(message, EventLogEntryType.Information);
        }

    }
}
