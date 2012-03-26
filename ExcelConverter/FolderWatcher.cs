using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;

namespace ExcelConverter
{
    class FolderWatcher : FileSystemWatcher
    {
        int _timeout;
        public FolderWatcher(string folderPath, int timeout)
        {
            Path = folderPath;
            _timeout = timeout;
            startActivityMonitoring();
        }

        public FolderWatcher(string folderPath)
        {
            Path = folderPath;
            _timeout = 5;
            startActivityMonitoring();
        }

        ~FolderWatcher()
        {
            abortActivityMonitoring();
        }

        private void startActivityMonitoring()
        {
            // Make sure you use the OR on each Filter because we need to monitor
            // all of those activities

            NotifyFilter = System.IO.NotifyFilters.DirectoryName;

            NotifyFilter =
            NotifyFilter | System.IO.NotifyFilters.FileName;
            NotifyFilter =
            NotifyFilter | System.IO.NotifyFilters.Attributes;

            // Now hook the triggers(events) to our handler (eventRaised)
            Changed += new FileSystemEventHandler(eventRaised);
            Created += new FileSystemEventHandler(eventRaised);
            Deleted += new FileSystemEventHandler(eventRaised);
            // Occurs when a file or directory is renamed in the specific path
            Renamed += new RenamedEventHandler(eventRaised);

            // And at last.. We connect our EventHandles to the system API (that is all
            // wrapped up in System.IO)
            try
            {
                EnableRaisingEvents = true;
            }
            catch (ArgumentException e)
            {
                abortActivityMonitoring();
                Logger.Write(@"Argument exception: " + e.Message, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        private void abortActivityMonitoring()
        {
            EnableRaisingEvents = false;
        }

        /// <summary>
        /// Triggered when an event is raised from the folder activity monitoring.
        /// All types exists in System.IO
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e">containing all data send from the event that got executed.</param>
        private void eventRaised(object sender, FileSystemEventArgs e)
        {
            string sEvent = @"";

            switch (e.ChangeType)
            {
                case WatcherChangeTypes.Changed:
                    if (convertFileIfNecessary(e.FullPath))
                    {
                        //Do Nothing.  It worked.
                    }
                    break;
                case WatcherChangeTypes.Created:
                    if (convertFileIfNecessary(e.FullPath))
                    {
                        //Do Nothing.  It worked.
                    }
                    break;
                case WatcherChangeTypes.Deleted:
                    //Optional: Do stuff here for this condition
                    break;
                case WatcherChangeTypes.Renamed:
                    RenamedEventArgs re = (RenamedEventArgs) e;
                    if (convertFileIfNecessary(e.FullPath))
                    {
                        string convertedOldFilePath = System.IO.Path.GetDirectoryName(re.OldFullPath) + @"\" + System.IO.Path.GetFileNameWithoutExtension(re.OldFullPath) + @".xlsx";
                        if (File.Exists(convertedOldFilePath))
                            File.Delete(convertedOldFilePath);
                    }
                    break;
                default: // Another action
                    //Optional: Do stuff here for this condition
                    break;
            }
            if (sEvent.Length > 0)
                Logger.Write(sEvent);
        }

        private bool convertFileIfNecessary(string filePath)
        {
            if (System.IO.Path.GetExtension(filePath).ToLower() == @".xls")
            {
                //ExcelStuff(e.FullPath);
                ExcelFile _xlFile = new ExcelFile(filePath, _timeout);
                return _xlFile.Convert();
            }
            return false;
        }
    }
}
