using System.ServiceProcess;
using System;

namespace ExcelConverter
{
	public partial class ExcelConverterService : ServiceBase
	{

		FolderWatcher _folderWatcher;

		public ExcelConverterService()
		{
			InitializeComponent();
			this.ServiceName = @"Excel Converter Service";
		}

		protected override void OnStart(string[] args)
		{
            Parser.Parameters _params = Parser.Parse(args);
            _folderWatcher = new FolderWatcher(_params.folderToWatch, _params.timeout);
            base.OnStart(args);
            Logger.Write(String.Format(@"Excel Converter started with params: timeout = {0}, watch folder = {1}.", _params.timeout.ToString(), _params.folderToWatch));
		}

		protected override void OnStop()
		{
			base.OnStop();
		}
	}
}
