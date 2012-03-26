using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Threading;

namespace ExcelConverter
{
    class ExcelFile
    {
        string _fileName;
        int _timeout;

        public ExcelFile(string fileName, int timeout)
        {
            _fileName = fileName;
            _timeout = timeout;

            // Add handler for unhandled exceptions (because file writing can be unpredictable)
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(UnhandledException);
        }

        public ExcelFile(string fileName)
        {
            _fileName = fileName;
            _timeout = 5;

            // Add handler for unhandled exceptions (because file writing can be unpredictable)
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(UnhandledException);
        }

        private string ResolveExcelCnvLocation()
        {
            string filePath = @"";
            string programFiles = Environment.GetEnvironmentVariable("ProgramFiles(x86)") + "\\Microsoft Office\\";
            try
            {
                string[] officeDirs = Directory.GetDirectories(programFiles, @"Office*");
                for (int i = 0; i < officeDirs.Length; i++)
                {
                    if (File.Exists(officeDirs[i] + "\\excelcnv.exe"))
                        filePath = officeDirs[i] + "\\excelcnv.exe";
                }
            }
            catch (Exception e)
            {
                LogExceptionDetails(e);
            }
            if (filePath.Equals(@""))
            {
                filePath = Constants.excelcnvLocation;
                if (!File.Exists(filePath))
                    LogExceptionDetails(new Exception("Failed to resolve excelcnv.exe location.\r\nPlease ensure that the file exists and LocalService has permissions to execute it."));
            }
            return filePath;
        }

        public bool Convert()
        {
            /// <summary>
            /// Store the Application object we can use in the member functions.
            /// </summary>
            bool ret = false;
            string newFileName = Path.GetDirectoryName(_fileName) + @"\" + Path.GetFileNameWithoutExtension(_fileName) + ".xlsx";
            string processFilePath = ResolveExcelCnvLocation();
            string processArguments = " -oice \"" + _fileName + "\" \"" + newFileName + "\"";

            Process process = new Process();
            process.StartInfo.FileName = processFilePath;
            process.StartInfo.Arguments = processArguments;
            process.StartInfo.UseShellExecute = false;

            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.RedirectStandardError = true;

            try
            {
                process.Start();

                process.BeginOutputReadLine();
                process.BeginErrorReadLine();

                if (process.WaitForExit(_timeout * 1000))
                {
                    ret = true;
                }
                else
                {
                    process.Kill();
                }
            }
            catch (Exception e)
            {
                LogExceptionDetails(e);
            }
            finally
            {
                process.Close();
            }
            return ret;
        }

        static void UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {// All exceptions thrown by additional threads are handled in this method
            LogExceptionDetails(e.ExceptionObject as Exception);
        }

        static void LogExceptionDetails(Exception e)
        {
            Logger.Write(e.Message, System.Diagnostics.EventLogEntryType.Error);
        }

    }
}
