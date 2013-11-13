Enumerator.prototype.toArray = function ()
{
	var Result = [];
	for (;!this.atEnd();this.moveNext())
	Result.push(this.item())
	return Result;
}

// Add ECMA262-5 method binding if not supported natively
//
if (!('bind' in Function.prototype)) {
    Function.prototype.bind= function(owner) {
        var that= this;
        if (arguments.length<=1) {
            return function() {
                return that.apply(owner, arguments);
            };
        } else {
            var args= Array.prototype.slice.call(arguments, 1);
            return function() {
                return that.apply(owner, arguments.length===0? args : args.concat(Array.prototype.slice.call(arguments)));
            };
        }
    };
}

// Add ECMA262-5 string trim if not supported natively
//
if (!('trim' in String.prototype)) {
    String.prototype.trim= function() {
        return this.replace(/^\s+/, '').replace(/\s+$/, '');
    };
}

// Add ECMA262-5 Array methods if not supported natively
//
if (!('indexOf' in Array.prototype)) {
    Array.prototype.indexOf= function(find, i /*opt*/) {
        if (i===undefined) i= 0;
        if (i<0) i+= this.length;
        if (i<0) i= 0;
        for (var n= this.length; i<n; i++)
            if (i in this && this[i]===find)
                return i;
        return -1;
    };
}

if (!('lastIndexOf' in Array.prototype)) {
    Array.prototype.lastIndexOf= function(find, i /*opt*/) {
        if (i===undefined) i= this.length-1;
        if (i<0) i+= this.length;
        if (i>this.length-1) i= this.length-1;
        for (i++; i-->0;) /* i++ because from-argument is sadly inclusive */
            if (i in this && this[i]===find)
                return i;
        return -1;
    };
}

if (!('forEach' in Array.prototype)) {
    Array.prototype.forEach= function(action, that /*opt*/) {
        for (var i= 0, n= this.length; i<n; i++)
            if (i in this)
                action.call(that, this[i], i, this);
    };
}

if (!('map' in Array.prototype)) {
    Array.prototype.map= function(mapper, that /*opt*/) {
        var other= new Array(this.length);
        for (var i= 0, n= this.length; i<n; i++)
            if (i in this)
                other[i]= mapper.call(that, this[i], i, this);
        return other;
    };
}

if (!('filter' in Array.prototype)) {
    Array.prototype.filter= function(filter, that /*opt*/) {
        var other= [], v;
        for (var i=0, n= this.length; i<n; i++)
            if (i in this && filter.call(that, v= this[i], i, this))
                other.push(v);
        return other;
    };
}

if (!('every' in Array.prototype)) {
    Array.prototype.every= function(tester, that /*opt*/) {
        for (var i= 0, n= this.length; i<n; i++)
            if (i in this && !tester.call(that, this[i], i, this))
                return false;
        return true;
    };
}

if (!('some' in Array.prototype)) {
    Array.prototype.some= function(tester, that /*opt*/) {
        for (var i= 0, n= this.length; i<n; i++)
            if (i in this && tester.call(that, this[i], i, this))
                return true;
        return false;
    };
}

var FSO = fso = new ActiveXObject("Scripting.FileSystemObject");
var WshShell= new ActiveXObject("WScript.Shell");
WshShell.CurrentDirectory= (WScript.ScriptFullName+"").replace(/[^\\]+$/g,"");

function GenerateString(L)
{
	if (!L) L=8;
	return new ActiveXObject('Scriptlet.TypeLib').Guid.replace(/[^\w]+/g,"").slice(0,L);
}

function WMIQuery(Moniker, Query)
{
	var Service =  GetObject(Moniker),
	Items = Service.ExecQuery(Query);
	return new Enumerator(Items).toArray();
}

// Autoit Macro

//@AppDataCommonDir Path to Application Data +
function GetAppDataCommonDir()
{
	var AppDataCommonDir = 0x023;
	var oShell = new ActiveXObject("Shell.Application");
	Result = oShell.NameSpace(AppDataCommonDir).Self.Path;
	oShell = null;
	return Result;
}
//WScript.Echo(GetAppDataCommonDir());
//WScript.Quit();

//@AppDataDir Path to current user's Application Data +
function GetAppDataDir()
{
	return WshShell.ExpandEnvironmentStrings("%APPDATA%");
}
//WScript.Echo(GetAppDataDir());
//WScript.Quit();

//@Exe The full path and filename of the AutoIt executable currently running. For compiled scripts it is the path of the compiled script. +
function GetExe()
{
	return WScript.FullName;
}
//WScript.Echo(GetExe());
//WScript.Quit();

//@PID PID of the process running the script. +
function GetPID()
{
	var ChildProc = new ActiveXObject("WScript.Shell").Exec("rundll32 kernel32,Sleep").ProcessId;
	var test=GetObject("winmgmts:\\\\.\\root\\cimv2:win32_process.Handle='" +ChildProc+ "'");
	var Result = test.ParentProcessId;
	test.Terminate();
	return Result;
}
//WScript.Echo(GetPID());
//WScript.Quit();

//@CommonFilesDir Path to Common Files folder +
function GetCommonFilesDir()
{
	var GetCommonFilesDir = 0x017;
	var oShell = new ActiveXObject("Shell.Application");
	Result = oShell.NameSpace(GetCommonFilesDir).Self.Path;
	oShell = null;
	return Result;
}
//WScript.Echo(GetCommonFilesDir());
//WScript.Quit();

//@ComputerName Computer's network name. +
function GetComputerName()
{
	return new ActiveXObject("WScript.Network").ComputerName;
}
//WScript.Echo(GetComputerName());
//WScript.Quit();

//@ComSpec Value of %comspec%, the SPECified secondary COMmand interpreter;
function GetComSpec()
{
	return WshShell.ExpandEnvironmentStrings("%COMSPEC%");
}
//WScript.Echo(GetComSpec());
//WScript.Quit();

//@CPUArch Returns "X86" when the CPU is a 32-bit CPU and "X64" when the CPU is 64-bit. +
function GetCPUArc()
{
	return isFolder("%WinDir%\\syswow64")? "X64": "X86";
}
//WScript.Echo(GetCPUArc());
//WScript.Quit();

//@DesktopCommonDir Path to Desktop +
function GetDesktopCommonDir()
{
	var SpecialDir = 0x019;
	var oShell = new ActiveXObject("Shell.Application");
	Result = oShell.NameSpace(SpecialDir).Self.Path;
	oShell = null;
	return Result;
}
//WScript.Echo(GetDesktopCommonDir());
//WScript.Quit();

//@DesktopDir Path to current user's Desktop +
function GetDesktopDir()
{
	var SpecialDir = 0x0;
	var oShell = new ActiveXObject("Shell.Application");
	Result = oShell.NameSpace(SpecialDir).Self.Path;
	oShell = null;
	return Result;
}
//WScript.Echo(GetDesktopDir());
//WScript.Quit();

//@DocumentsCommonDir Path to Documents 
function GetDocumentsCommonDir()
{
	var SpecialDir = 0x2E;
	var oShell = new ActiveXObject("Shell.Application");
	Result = oShell.NameSpace(SpecialDir).Self.Path;
	oShell = null;
	return Result;
}
//WScript.Echo(GetDocumentsCommonDir());
//WScript.Quit();

//@FavoritesCommonDir Path to Favorites 
function GetFavoritesCommonDir()
{
	var SpecialDir = 0x1F;
	var oShell = new ActiveXObject("Shell.Application");
	Result = oShell.NameSpace(SpecialDir).Self.Path;
	oShell = null;
	return Result;
}
//WScript.Echo(GetFavoritesCommonDir());
//WScript.Quit();

//@FavoritesDir Path to current user's Favorites 
function GetFavoritesDir()
{
	var SpecialDir = 0x6;
	var oShell = new ActiveXObject("Shell.Application");
	Result = oShell.NameSpace(SpecialDir).Self.Path;
	oShell = null;
	return Result;
}
//WScript.Echo(GetFavoritesDir());
//WScript.Quit();


//@MyDocumentsDir Path to My Documents target 
function GetMyDocumentsDir()
{
	return new ActiveXObject("Wscript.Shell").SpecialFolders("Mydocuments");
}
//WScript.Echo(GetMyDocumentsDir());
//WScript.Quit();

//@HomePath Directory part of current user's home directory. To get the full path, use in conjunction with @HomeDrive. 
function GetHomePath()
{
	return WshShell.ExpandEnvironmentStrings("%HomePath%");
}
//WScript.Echo(GetHomePath());
//WScript.Quit();

//@HomeDrive Drive letter of drive containing current user's home directory. 
function GetHomeDrive()
{
	return WshShell.ExpandEnvironmentStrings("%userprofile%").replace(/:.+$/, ":");
}
//WScript.Echo(GetHomeDrive());
//WScript.Quit();

//@StartMenuCommonDir Path to Start Menu folder 
function GetStartMenuCommonDir()
{
	var SpecialDir = 0x16;
	var oShell = new ActiveXObject("Shell.Application");
	Result = oShell.NameSpace(SpecialDir).Self.Path;
	oShell = null;
	return Result;
}
//WScript.Echo(GetStartMenuCommonDir());
//WScript.Quit();

//@StartMenuDir Path to current user's Start Menu 
function GetStartMenuDir()
{
	var SpecialDir = 0x1D;
	var oShell = new ActiveXObject("Shell.Application");
	Result = oShell.NameSpace(SpecialDir).Self.Path;
	oShell = null;
	return Result;
}
//WScript.Echo(GetStartMenuDir());
//WScript.Quit();

//@UserProfileDir Path to current user's Profile folder. 
function GetUserProfileDir()
{
	return WshShell.ExpandEnvironmentStrings("%userprofile%");
}
//WScript.Echo(GetUserProfileDir());
//WScript.Quit();

//@SystemDir Path to the Windows' System (or System32) folder 
function GetSystemDir()
{
	var SpecialDir = 0x25;
	var oShell = new ActiveXObject("Shell.Application");
	Result = oShell.NameSpace(SpecialDir).Self.Path;
	oShell = null;
	return Result;
}
//WScript.Echo(GetSystemDir());
//WScript.Quit();

//@TempDir Path to the temporary files folder. 
function GetTempDir()
{
	return WshShell.ExpandEnvironmentStrings("%TEMP%");
}
//WScript.Echo(GetTempDir());
//WScript.Quit();

//@ProgramFilesDir Path to Program Files folder 
function GetProgramFilesDir()
{
	var SpecialDir = 0x26;
	var oShell = new ActiveXObject("Shell.Application");
	Result = oShell.NameSpace(SpecialDir).Self.Path;
	oShell = null;
	return Result;
}
//WScript.Echo(GetProgramFilesDir());
//WScript.Quit();

//@ProgramsCommonDir Path to Start Menu's Programs folder 
function GetProgramsCommonDir()
{
	var SpecialDir = 0x17;
	var oShell = new ActiveXObject("Shell.Application");
	Result = oShell.NameSpace(SpecialDir).Self.Path;
	oShell = null;
	return Result;
}
//WScript.Echo(GetProgramsCommonDir());
//WScript.Quit();

//@ProgramsDir Path to current user's Programs (folder on Start Menu) 
function GetProgramsDir()
{
	var SpecialDir = 0x2;
	var oShell = new ActiveXObject("Shell.Application");
	Result = oShell.NameSpace(SpecialDir).Self.Path;
	oShell = null;
	return Result;
}
//WScript.Echo(GetProgramsDir());
//WScript.Quit();

//@WindowsDir Path to Windows folder 
function GetWindowsDir()
{
	return WshShell.ExpandEnvironmentStrings("%WINDIR%");
}
//WScript.Echo(GetWindowsDir());
//WScript.Quit();

//@WorkingDir Current/active working directory. (Result doesn't contain a trailing backslash) 
function GetWorkingDir()
{
	return WshShell.CurrentDirectory;
}
//WScript.Echo(GetWorkingDir());
//WScript.Quit();

//@ScriptDir Directory containing the running script. (Result doesn't contain a trailing backslash) 
function GetScriptDir()
{
	return (WScript.ScriptFullName+"").replace(/[^\\]+$/g,"");
}
//WScript.Echo(GetScriptDir());
//WScript.Quit();

//@ScriptFullPath Equivalent to @ScriptDir & "\" & @ScriptName 
function GetScriptFullPath()
{
	return WScript.ScriptFullName;
}
//WScript.Echo(GetScriptFullPath());
//WScript.Quit();

//@ScriptName Long filename of the running script. 
function GetScriptName()
{
	return WScript.ScriptName;
}
//WScript.Echo(GetScriptName());
//WScript.Quit();

//@DesktopHeight Height of the desktop screen in pixels. (Vertical resolution) 
function GetDesktopHeight()
{
	var mon = WMIQuery("Winmgmts:\\\\.\\root\\cimv2", "Select * From Win32_DesktopMonitor where DeviceID = 'DesktopMonitor1'")[0];
	return mon.ScreenHeight
}
//WScript.Echo(GetDesktopHeight());
//WScript.Quit();

//@DesktopWidth Width of the desktop screen in pixels. (Horizontal resolution) 
function GetDesktopWidth()
{
	var mon = WMIQuery("winmgmts:\\\\.\\root\\cimv2", "Select * From Win32_DesktopMonitor where DeviceID = 'DesktopMonitor1'")[0];
	return mon.ScreenWidth
}
//WScript.Echo(GetDesktopWidth());
//WScript.Quit();

//@DesktopDepth Depth of the desktop screen in bits per pixel.  
function GetDesktopDepth()
{
	var mon = WMIQuery("winmgmts:\\\\.\\root\\cimv2", "Select * From Win32_DisplayConfiguration")[0];
	return mon.BitsPerPel;
}
//WScript.Echo(GetDesktopDepth());
//WScript.Quit();

//@HOUR Hours value of clock in 24-hour format. Range is 00 to 23 
function GetHOUR()
{
	var d = new Date();
	return d.getHours(); 
}

//@MDAY Current day of month. Range is 01 to 31 
function GetMDAY()
{
	var d = new Date();
	return d.getDate(); 
}

//@MIN Minutes value of clock. Range is 00 to 59 
function GetMIN()
{
	var d = new Date();
	return d.getMinutes(); 
}

//@MON Current month. Range is 01 to 12 
function GetMON()
{
	var d = new Date();
	return d.getMonth()+1; 
}

//@MSEC Milliseconds value of clock.  Range is 00 to 999. The update frequency of this value depends on the timer resolution of the hardware and may not update every millisecond. 
function GetMSEC()
{
	var d = new Date();
	return d.getMilliseconds();
}

//@SEC Seconds value of clock. Range is 00 to 59 
function GetSEC()
{
	var d = new Date();
	return d.getSeconds();
}

//@WDAY Numeric day of week. Range is 1 to 7 which corresponds to Sunday through Saturday. 
function GetWDAY()
{
	var d = new Date();
	return d.getDay()+1;
}

//@YDAY Current day of year. Range is 001 to 366 (or 001 to 365 if not a leap year) 
function GetYDAY()
{
	var d = new Date();
	var yn = d.getFullYear();
	var mn = d.getMonth();
	var dn = d.getDate();
	var d1 = new Date(yn,0,1,12,0,0); // noon on Jan. 1
	var d2 = new Date(yn,mn,dn,12,0,0); // noon on input date
	var ddiff = Math.round((d2-d1)/864e5);
	return ddiff+1;
}
//WScript.Echo(GetYDAY());
//WScript.Quit();

//@YEAR Current four-digit year 
function GetYEAR()
{
	var d = new Date();
	return d.getFullYear();
}

//@UserName ID of the currently logged on user. 
function GetUserName()
{
	return new ActiveXObject("WScript.Network").UserName;
}
//WScript.Echo(GetUserName());
//WScript.Quit();

//@LogonServer Logon server. 
function GetLogonServer()
{
	return WshShell.ExpandEnvironmentStrings("%LOGONSERVER%");
}
//WScript.Echo(GetLogonServer());
//WScript.Quit();

//@IPAddress1 IP address of first network adapter. Tends to return 127.0.0.1 on some computers. 
function GetIPAddress1()
{
	var res = WMIQuery("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\cimv2", "Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")
	return res[0].IPAddress(0);
}
//WScript.Echo(GetIPAddress1());
//WScript.Quit();

function GetMacAddress()
{
	var res = WMIQuery("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\cimv2", "Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")
	return res[0].IPAddress(1);
}
//WScript.Echo(GetMacAddress());
//WScript.Quit();

//@PublicIPAddress Public IP address. 
function GetPublicIPAddress()
{
	try {
		var o = new ActiveXObject("MSXML2.XMLHTTP");
		o.open("GET", "http://ifconfig.me/ip", false);
		o.send();
		return o.responseText;
	} catch (e) {return ""}
}
//WScript.Echo(GetPublicIPAddress());
//WScript.Quit();

//@OSArch Returns one of the following: "X86", "IA64", "X64" - this is the architecture type of the currently running operating system. 
function GetOSArch()
{
	return WshShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%").toUpperCase();
}
//WScript.Echo(GetOSArch());
//WScript.Quit();

//@OSBuild Returns the OS build number. For example, Windows 2003 Server returns 3790 
function GetOSBuild()
{
	var Windows = WMIQuery("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\cimv2", "Select * from Win32_OperatingSystem" )[0];
	return Windows.BuildNumber;
}
//WScript.Echo(GetOSBuild());
//WScript.Quit();

//@OSLang Returns code denoting OS Language.  See Appendix for possible values. 
function GetOSLang()
{
	var Windows = WMIQuery("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\cimv2", "Select * from Win32_OperatingSystem" )[0];
	return Windows.OSLanguage;
}
//WScript.Echo(GetOSLang());
//WScript.Quit();

function GetOSInstallDate()
{
	var dtmConvertedDate = new ActiveXObject("WbemScripting.SWbemDateTime");
	var Windows = WMIQuery("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\cimv2", "Select * from Win32_OperatingSystem" )[0];
	dtmConvertedDate.Value = Windows.InstallDate;
	return dtmConvertedDate.GetVarDate();
}
//WScript.Echo(GetOSInstallDate());
//WScript.Quit();

//@OSServicePack Service pack info in the form of "Service Pack 3". 
function GetOSServicePack()
{
	var Windows = WMIQuery("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\cimv2", "Select * from Win32_OperatingSystem" )[0];
	return "Service Pack " + Windows.ServicePackMajorVersion  + "." + Windows.ServicePackMinorVersion;
}
//WScript.Echo(GetOSServicePack());
//WScript.Quit();

//@OSVersion Returns one of the following: "WIN_2008R2", "WIN_7", "WIN_8", "WIN_2008", "WIN_VISTA", "WIN_2003", "WIN_XP", "WIN_XPe", "WIN_2000". 
function GetOSVersion()
{
	var Windows = WMIQuery("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\cimv2", "Select * from Win32_OperatingSystem" )[0];
	
	if (/windows +7/i.test(Windows.Caption)) return "WIN_7";
	if (/2000/i.test(Windows.Caption)) return "WIN_2000";
	if (/windows +8/i.test(Windows.Caption)) return "WIN_8";
	if (/2008 +R2/i.test(Windows.Caption)) return "WIN_2008R2";
	if (/2008/i.test(Windows.Caption)) return "WIN_2008";
	if (/2003/i.test(Windows.Caption)) return "WIN_2003";
	if (/ +xp +e/i.test(Windows.Caption)) return "WIN_XPe";
	if (/ +xp/i.test(Windows.Caption)) return "WIN_XP";
	if (/vista/i.test(Windows.Caption)) return "WIN_VISTA";
	
	return "UNKNOWN_OS";
}
//WScript.Echo(GetOSVersion());
//WScript.Quit();

//Process Managment

//ProcessClose Terminates a named process.
function ProcessClose(ProcName) //pid or name
{
	if (typeof ProcName == "number")
	{
	try {
			var Proc = WMIQuery("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\CIMV2", "SELECT * FROM Win32_Process WHERE  ProcessId='"+ProcName+"'");
			if (Proc.length>0)
			for (var i=0, l=Proc.length; i<l; i++)
			Proc[i].Terminate();
			return 1;
	} catch(e) {
		return 0;
	}
	}
	
	try {
			var Proc = WMIQuery("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\CIMV2", "SELECT * FROM Win32_Process WHERE Name='"+ProcName+"'");
			for (var i=0, l=Proc.length; i<l; i++)
			Proc[i].Terminate();
			return 1;
	} catch(e) {
		return 0;
	}
}
//WScript.Echo(ProcessClose("notepad.exe"));
//WScript.Quit();

//ProcessExists Checks to see if a specified process exists. 
function ProcessExists(ProcName) //pid or name
{
	if (typeof ProcName == "number")
	{
	try {
			var Proc = WMIQuery("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\CIMV2", "SELECT * FROM Win32_Process WHERE  ProcessId='"+ProcName+"'");
			if (Proc.length>0) return Proc[0].ProcessId;
			return 0;
	} catch(e) {
		return 0;
	}
	}
	
	try {
			var Proc = WMIQuery("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\CIMV2", "SELECT * FROM Win32_Process WHERE Name='"+ProcName+"'");
			if (Proc.length>0) return Proc[0].ProcessId;
			return 0;
	} catch(e) {
		return 0;
	}
}
//WScript.Echo(ProcessExists("notepad.exe"));
//WScript.Quit();

//Pauses script execution until a given process exists.
function ProcessWait(process, timeout)
//process The name of the process to check. 
//timeout [optional] Specifies how long to wait (in seconds). Default is to wait indefinitely. 
{
	timeout = timeout || 0;
	var time_tmp=0, pid=0;
	while (pid=ProcessExists(process))
	{
		WScript.Sleep(250);
		if (timeout!=0 && time_tmp>=timeout) break; else time_tmp+=0.250; 
	}
	return pid;
}
//WScript.Echo(ProcessWait("notepad.exe"));
//WScript.Quit();

//Pauses script execution until a given process does not exist.
function ProcessWaitClose(process, timeout)
//process The name of the process to check. 
//timeout [optional] Specifies how long to wait (in seconds). Default is to wait indefinitely. 
{
	timeout = timeout || 0;
	var time_tmp=0, pid=0;
	while (!(pid=ProcessExists(process)))
	{
		WScript.Sleep(250);
		if (timeout!=0 && time_tmp>=timeout) break; else time_tmp+=0.250; 
	}
	return pid;
}
//WScript.Echo(ProcessWaitClose("notepad.exe"));
//WScript.Quit();


// File, Directory and Disk Managment

//DirCopy Copies a directory and all sub-directories and files (Similar to xcopy).
function DirCopy(SourceDir, DestDir, flag)
{
	SourceDir = WshShell.ExpandEnvironmentStrings(SourceDir);
	CreateDir(DestDir);
	
	flag = flag || 0;
	FSO.CopyFolder(SourceDir, DestDir, flag);
	return 1;
}

//DirCreate Creates a directory/folder.
function CreateDir(Path)
{
	Path = WshShell.ExpandEnvironmentStrings(Path);
	__FolderCreate(Path);
	return 1;
}

function __FolderCreate(FolderPath)
{
	if (!fso.FolderExists(FolderPath))
	{
		__FolderCreate(fso.GetParentFolderName(FolderPath));
		fso.CreateFolder(FolderPath);
	}
}

//FileDelete Delete one or more files. (принудительное удаление. принимает как массивы файлов, так и просто путь) поддерживает удаление по маске.
function FileDelete(Path)
{
    if (/Array/i.test(Path.constructor+"")) 
    {
		for (var i=0, l=Path.length;i<l;i++)
		DeleteFile(Path[i]);
		
		return true;
    }

	Path = WshShell.ExpandEnvironmentStrings(Path);
	try {
		FSO.DeleteFile(Path, true); return 1;
	}catch (e) {return 0}
}

//FileExists Checks if a file or directory exists.
function FileExists(Path)
{
	Path = WshShell.ExpandEnvironmentStrings(Path);
	return FSO.FolderExists(Path) || FSO.FileExists(Path);
}

//DirMove Moves a directory and all sub-directories and files.
function DirMove(SourceDir, DestDir)
{
	try {
		FSO.MoveFolder(SourceDir, DestDir, true);
		return 1;
	} catch(e) {return 0;}
	
}

//FileMove Moves one or more files
function FileMove(SourceFile, DestFile)
{
	if (/Array/i.test(Path.constructor+"")) 
    {
		for (var i=0, l=Path.length;i<l;i++)
		FileMove(Path[i]);
		return 1;
    }
	
	try {
		FSO.MoveFile(SourceDir, DestDir);
		return 1;
	} catch(e) {return 0;}
	
}

//DirRemove Deletes a directory/folder.
function DirRemove(Path)
{
	Path = WshShell.ExpandEnvironmentStrings(Path.replace(/\\+$/, ""));
	var main_folder = fso.GetFolder(Path);

	function DirWithSubFolders(_folder)
	{
		new Enumerator(_folder.SubFolders).toArray().forEach(function (f){try {DirWithSubFolders(_folder);} catch (e) {} });
		try {_folder.Delete(true);} catch(e) {}
	}

	DirWithSubFolders(main_folder);
	try {main_folder.Delete(true);} catch(e) {}
}
//DirRemove("%TMP%\\Fine.SSR11");
//WScript.Quit();

//FileCopy Copies one or more files. 
function FileCopy(source, dest, flag)
{
	flag = flag || 0;
	try {
		FSO.CopyFile(source, dest, flag)
		return 1;
	} catch(e) {return 0;}
}

// DriveGetDrive Returns an array containing the enumerated drives.
function DriveGetDrive(type, available_flag)
{
	type = type || "ALL";
	type = type.toUpperCase();
	var Drives=[];
	var D = (new Enumerator(FSO.Drives).toArray().toString().replace(/A:,/,"").split(","));
	
	available_flag = available_flag || 0;
	if (!available_flag) Drives = D; 
	
	if (available_flag)
	{
		Drives=[];
		for (var i=0; i<D.length; i++)
		{
			try {
				if (FSO.GetDrive(D[i]).IsReady) {Drives.push(D[i]); ;}
			} catch (e) {}
		}
	}
	
	if (type=="ALL") return Drives;
	
	//type "CDROM", "REMOVABLE", "FIXED", "NETWORK", "RAMDISK", or "UNKNOWN"
/*
    0 - неизвестное устройство.
    1 - устройство со сменным носителем.
    2 - жёсткий диск.
    3 - сетевой диск.
    4 - CD-ROM.
    5 - RAM-диск.
*/
	var types=["UNKNOWN", "REMOVABLE", "FIXED", "NETWORK", "CDROM", "RAMDISK"];
	
	var Result = [];
	type = types.indexOf(type);
	
	return Drives.filter(function (d){return FSO.GetDrive(d).DriveType==type});
}
//WScript.Echo(DriveGetDrive("ALL", true));
//WScript.Echo(DriveGetDrive("ALL"));
//WScript.Quit()

//FileGetVersion Returns the "File" version information.
function FileGetVersion(filespec){
	return ""+fso.GetFileVersion(WshShell.ExpandEnvironmentStrings(filespec));
}

//FileChangeDir Changes the current working directory.
function FileChangeDir(dir){
	try {
		WshShell.CurrentDirectory=WshShell.ExpandEnvironmentStrings(dir);
		return 1;
	} catch(e) {
		return 0;
	}
	
}

//DriveGetFileSystem Returns File System Type of a drive.
function DriveGetFileSystem(Drive)
{
	Drive=Drive.replace(/:.*$/, ":\\");
	return FSO.GetDrive(Drive).FileSystem;
}

//DriveGetLabel Returns Volume Label of a drive, if it has one.
function DriveGetLabel(Drive)
{
	Drive=Drive.replace(/:.*$/, ":\\");
	return FSO.GetDrive(Drive).VolumeName;
}
//WScript.Echo(DriveGetLabel(GetHomeDrive()));
//WScript.Quit();


//DriveGetSerial Returns Serial Number of a drive.
function DriveGetSerial(Drive)
{
	Drive=Drive.replace(/:.*$/, ":\\");
	return FSO.GetDrive(Drive).SerialNumber;
}

//DriveGetType Returns drive type.
function DriveGetType(Drive)
{
	Drive=Drive.replace(/:.*$/, ":\\");
/*
    0 - неизвестное устройство.
    1 - устройство со сменным носителем.
    2 - жёсткий диск.
    3 - сетевой диск.
    4 - CD-ROM.
    5 - RAM-диск.
*/
	var types=["UNKNOWN", "REMOVABLE", "FIXED", "NETWORK", "CDROM", "RAMDISK"];
	return types[FSO.GetDrive(Drive).DriveType];
}
//WScript.Echo(DriveGetType(GetHomeDrive()));
//WScript.Quit();

//DriveMapAdd Maps a network drive.
function DriveMapAdd(strLocalDrive, strRemoteShare, persistent, strUser, strPassword)
{
	try {
		persistent = persistent || "";
		strUser = strUser || ""; 
		strPassword = strPassword || "";
	
		var objNetwork = new ActiveXObject("WScript.Network");
		objNetwork.MapNetworkDrive(strLocalDrive, strRemoteShare, persistent, strUser, strPassword);
		return 1;
	} catch(e) {
		return 0;
	}
}

//DriveMapDel Disconnects a network drive.
function DriveMapDel(strLocalDrive)
{
	try {
		var objNetwork = new ActiveXObject("WScript.Network");
		objNetwork.RemoveNetworkDrive(strName, true);
		return 1;
	} catch(e) {
		return 0;
	}
}

//DriveMapGet Retrieves the details of a mapped drive.
function DriveMapGet(strDrive)
{
	var strDrive = strDrive.replace(/:.*$/, ":\\");
	try {
		return WMIQuery("winmgmts:\\\\.\\root\\cimv2", "Select * from Win32_LogicalDisk Where DeviceID = '"+strDrive+"'")[0].ProviderName;
	} catch(e) {
		return strDrive;
	}
}

//DriveSetLabel Sets the Volume Label of a drive.
function DriveSetLabel(Drive, NewLabel)
{
	try {
			Drive=Drive.replace(/:.*$/, ":\\");
			FSO.GetDrive(Drive).VolumeName = NewLabel;
			return 1;
	} catch(e) {
		return 0;
	}
}
//DriveSetLabel(GetHomeDrive(), "Hello");
//WScript.Quit();

//DriveSpaceFree Returns the free disk space of a path in Megabytes.
function DriveSpaceFree(Drive)
{
	try {
			Drive=Drive.replace(/:.*$/, ":");
			return Math.round(FSO.GetDrive(Drive).FreeSpace/1024/1024);
	} catch(e) {
		return 0;
	}
}
//WScript.Echo(DriveSpaceFree("C:\\"));
//WScript.Quit();

//DriveSpaceTotal Returns the total disk space of a path in Megabytes.
function DriveSpaceTotal(Drive)
{
	try {
			Drive=Drive.replace(/:.*$/, ":");
			return Math.round(FSO.GetDrive(Drive).TotalSize/1024/1024);
	} catch(e) {
		return 0;
	}
}

//DriveStatus Returns the status of the drive as a string.
function DriveStatus(strDrive)
{
	var strDrive = strDrive.replace(/:.*$/, ":\\");
	try {
		FSO.GetDrive(strDrive);
	} catch(e) {
		return "INVALID";
	}
	
	return FSO.GetDrive(strDrive).IsReady? "READY":"NOTREADY";
}
//WScript.Echo(DriveStatus("d:"));
//WScript.Quit();
