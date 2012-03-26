REQUIREMENTS:
-----------------------------------------------------------------------
- ExcelCnv.exe (Microsoft-provided Excel converter)
	-- Comes installed with Office OR the free Office compatibility pack (http://www.microsoft.com/download/en/details.aspx?id=3)
	-- Located in %ProgramFiles(x86)%\Office\<Office Version>" <-- ExcelFile class attempts to autoresolve this
- .NET Framework 3.5 Client Profile (gets bundled in Setup.Exe upon build of ExcelConverterSetup


BUILD AND INSTALL INSTRUCTIONS:
-----------------------------------------------------------------------
- Edit Constants.cs
	-- If you set the startup type to "Automatic" you can ensure that the service continues to run in case of reboot; however,
	   the watch folder path in Constants.cs MUST be correct for this to work because I have not yet found a way to specify
	   default service parameters for automatic start.
- Optional: Make code changes (shouldn't be necessary unless you're smashing bugs)
- Build ExcelConverter
- Build ExcelConverterSetup
- Open the ExcelConverterSetup release folder and copy "ExcelConverterSetup.msi" and "setup.exe" to the target server
- Run "setup.exe" on the target server to install the .NET framework dependency; it will automatically call ExcelConverterSetup.msi.


POST-INSTALLATION SETUP INSTRUCTIONS:
-----------------------------------------------------------------------
- Ensure that "<Machine Name>\Local Service" has FULL CONTROL permissions to the watch folder
- Ensure that "<Machine Name>\Local Service" has READ and EXECUTE permissions on ExcelCnv.exe (listed above in Requirements)
- On Windows Server 2008, set ExcelCnv.exe to run in "XP Compatibility Mode" (learned this the hard way)
	-- Right click excelcnv.exe and select Properties
	-- Select "Edit for All Users"
	-- Check the box to "Run in Compatibility Mode For" and select "Windows XP"
- You can specify the timeout and watch folder path in the service startup parameters rather than rebuilding every time they need to change
	-- Syntax for multiple parameters uses a "/" before each parameter
	-- Example 1: "C:\WatchFolder"			<-- manually starts the process to watch "C:\WatchFolder" and use the default timeout of 5 seconds
	-- Example 2: /-t /10 /"C:\WatchFolder"	<-- manually starts the process to watch "C:\WatchFolder" and a 10-second timeout


HELPFUL PRE- AND POST-BUILD COMMANDS FOR DEBUG:
-----------------------------------------------------------------------
- Right click "ExcelConverter" project and open the build options.  Paste the following into the boxes to automatically deploy your
  code upon build.
- Pre-build commands to automatically stop the service if it is running and reinstall the newly built code:
	net stop "Excel Converter Service"
	C:\Windows\Microsoft.NET\Framework\v4.0.30319\InstallUtil.exe /u "$(TargetPath)"
	Exit /b 0
- Post-build commands to automatically start the newly installed service after successful build (substitute the path to your watch folder):
	C:\Windows\Microsoft.NET\Framework\v4.0.30319\InstallUtil.exe /ShowCallStack "$(TargetPath)"
	net start "Excel Converter Service" /-t /5 /"C:\Users\svines\WatcherTest"