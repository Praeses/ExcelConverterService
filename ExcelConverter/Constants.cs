using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ServiceProcess;

namespace ExcelConverter
{
    class Constants
    {
        //These can also be set as service startup parameters
        public const string watchFolderPath = @"F:\ExcelConverter";
        public const int defaultTimeout = 5;                                 
        public const string excelcnvLocation = "\"C:\\program files (x86)\\Microsoft Office\\Office14\\excelcnv.exe\"";
        public const ServiceStartMode startMode = ServiceStartMode.Automatic;
    }
}
