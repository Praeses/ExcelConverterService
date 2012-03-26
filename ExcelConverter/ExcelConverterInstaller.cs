using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.ServiceProcess;


namespace ExcelConverter
{
    [RunInstaller(true)]
    public partial class ExcelConverterInstaller : System.Configuration.Install.Installer
    {
        public ExcelConverterInstaller()
        {
            InitializeComponent();

            var processInstaller = new ServiceProcessInstaller();
            var serviceInstaller = new ServiceInstaller();

            processInstaller.Account = ServiceAccount.LocalService;
            
            serviceInstaller.DisplayName = @"Excel Converter Service";
            serviceInstaller.StartType = Constants.startMode;

            serviceInstaller.ServiceName = @"Excel Converter Service";

            this.Installers.Add(processInstaller);
            this.Installers.Add(serviceInstaller);

        }
    }
}
