using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace ExcelConverter
{
    class Parser
    {
        private static string usageString = @"Usage: ExcelConverter.exe [-t <timeout in seconds>] <path to watch folder>";

        public struct Parameters
        {
            public int timeout;
            public string folderToWatch;
        }

        private static int getTimeoutFromString(string timeoutString)
        {
            int _timeout = -1;

            if (timeoutString.StartsWith(@"/"))
                timeoutString = timeoutString.Substring(1);

            if (Int32.TryParse(timeoutString, out _timeout))
            {
                if (_timeout > 60)
                    _timeout = 60;
                else if (_timeout < 1)
                {
                    _timeout = 1;
                }
            }
            return _timeout;
        }

        public static Parameters Parse(string[] args)
        {
            if (args == null || args.Length <= 0)
            {
                if (Constants.watchFolderPath == null || Constants.watchFolderPath.Equals(@""))
                {
                    throw new Exception("Path of folder to watch not specified.\r\n" + usageString);
                }
                return new Parameters() { folderToWatch = Constants.watchFolderPath, timeout = Constants.defaultTimeout };
            }
            
            List<string> argList = args.ToList();

            int _timeout = Constants.defaultTimeout;

            for (int i=0; i<argList.Count; i++)
            {
                if (args[i].StartsWith(@"/"))
                    args[i] = args[i].Substring(1);

                if (args[i].ToLower().Equals(@"-t") && argList.Count > i+1)
                {
                    _timeout = getTimeoutFromString(args[i + 1]);
                }
            }

            string _folderToWatch = args[args.Length - 1];

            if (argList.Count < 1 || !Directory.Exists(_folderToWatch))
            {
                _folderToWatch = Constants.watchFolderPath;
                if (_folderToWatch == null || !Directory.Exists(_folderToWatch))
                    throw new Exception(String.Format("{0} not found.\r\nPlease ensure that the start parameters identify a valid watch folder and that Local Service has full control access to it.\r\n", _folderToWatch, System.Diagnostics.EventLogEntryType.Error)+usageString);
            }

            return new Parameters() { timeout = _timeout, folderToWatch = _folderToWatch };
        }
    }
}
