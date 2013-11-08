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

function Log(Data)
{

}

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
	var mon = WMIQuery("Winmgmts:\\\\.\\root\\cimv2", "Select * From Win32_DesktopMonitor where DeviceID = 'DesktopMonitor1'")[0];
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
		if (timeout<>0 && time_tmp>=timeout) break; else time_tmp+=0.250; 
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
	while (!(pid=ProcessExists(process))
	{
		WScript.Sleep(250);
		if (timeout<>0 && time_tmp>=timeout) break; else time_tmp+=0.250; 
	}
	return pid;
}
//WScript.Echo(ProcessWaitClose("notepad.exe"));
//WScript.Quit();

// File, Directory and Disk Managment
/*
TODO

DirCopy Copies a directory and all sub-directories and files (Similar to xcopy).
DirCreate Creates a directory/folder.
DirGetSize Returns the size in bytes of a given directory.
DirMove Moves a directory and all sub-directories and files.
DirRemove Deletes a directory/folder.
DriveGetDrive Returns an array containing the enumerated drives.
DriveGetFileSystem Returns File System Type of a drive.
DriveGetLabel Returns Volume Label of a drive, if it has one.
DriveGetSerial Returns Serial Number of a drive.
DriveGetType Returns drive type.
DriveMapAdd Maps a network drive.
DriveMapDel Disconnects a network drive.
DriveMapGet Retrieves the details of a mapped drive.
DriveSetLabel Sets the Volume Label of a drive.
DriveSpaceFree Returns the free disk space of a path in Megabytes.
DriveSpaceTotal Returns the total disk space of a path in Megabytes.
DriveStatus Returns the status of the drive as a string.
FileChangeDir Changes the current working directory.
FileClose Closes a previously opened text file.
FileCopy Copies one or more files. 
FileCreateNTFSLink Creates an NTFS hardlink to a file or a directory
FileCreateShortcut Creates a shortcut (.lnk) to a file.
FileDelete Delete one or more files.
FileExists Checks if a file or directory exists.
FileFindFirstFile Returns a search "handle" according to file search string.
FileFindNextFile Returns a filename according to a previous call to FileFindFirstFile.
FileFlush Flushes the file's buffer to disk.
FileGetAttrib Returns a code string representing a file's attributes.
FileGetEncoding Determines the text encoding used in a file.
FileGetLongName Returns the long path+name of the path+name passed.
FileGetPos Retrieves the current file position.
FileGetShortcut Retrieves details about a shortcut.
FileGetShortName Returns the 8.3 short path+name of the path+name passed.
FileGetSize Returns the size of a file in bytes.
FileGetTime Returns the time and date information for a file.
FileGetVersion Returns the "File" version information.
FileInstall Include and install a file with the compiled script.
FileMove Moves one or more files
FileOpen Opens a text file for reading or writing.
FileOpenDialog Initiates a Open File Dialog.
FileRead Read in a number of characters from a previously opened text file.
FileReadLine Read in a line of text from a previously opened text file.
FileRecycle Sends a file or directory to the recycle bin.
FileRecycleEmpty Empties the recycle bin.
FileSaveDialog Initiates a Save File Dialog.
FileSelectFolder Initiates a Browse For Folder dialog.
FileSetAttrib Sets the attributes of one or more files.
FileSetPos Sets the current file position.
FileSetTime Sets the timestamp of one of more files.
FileWrite Append a text/data to the end of a previously opened file.
FileWriteLine Append a line of text to the end of a previously opened text file.
IniDelete Deletes a value from a standard format .ini file.
IniRead Reads a value from a standard format .ini file.
IniReadSection Reads all key/value pairs from a section in a standard format .ini file.
IniReadSectionNames Reads all sections in a standard format .ini file.
IniRenameSection Renames a section in a standard format .ini file.
IniWrite Writes a value to a standard format .ini file.
IniWriteSection Writes a section to a standard format .ini file.
*/


// FileSystem Functions


// DeleteFile (принудительное удаление. принимает как массивы файлов, так и просто путь)
function DeleteFile(Path)
{
    if (/Array/i.test(Path.constructor+"")) 
    {
		for (var i=0, l=Path.length;i<l;i++)
		DeleteFile(Path[i]);
		
		return true;
    }

	Path = WshShell.ExpandEnvironmentStrings(Path);
	try {
		if (isFile(Path))
		FSO.GetFile(Path).Delete(true);
		
	}catch (e) {Log(Path +"  "+e.description);}
}

function isFile(Path)
{
	Path = WshShell.ExpandEnvironmentStrings(Path);
	return fso.FileExists(Path);
}

function isFolder(Path)
{
	Path = WshShell.ExpandEnvironmentStrings(Path);
	return fso.FolderExists(Path);
}

// есть файл или папка с таким именем
function Exists(Path)
{
	Path = WshShell.ExpandEnvironmentStrings(Path);
	return fso.FolderExists(Path) || fso.FileExists(Path);
}

function GetFiles(Path)
{
	var Result = [];
	Path = WshShell.ExpandEnvironmentStrings(Path);
	if (isExists(Path))
		return new Enumerator(FSO.GetFolder(Path).Files).toArray();
	else return [];
}

function GetDrives()
{
	var D = (new Enumerator(FSO.Drives).toArray().toString().replace(/A:,/,"").replace(/:/g,":\\").split(","));
	var Drives=[];
	for (var i=0; i<D.length; i++)
	{
		try {
			if (FSO.GetDrive(D[i]).IsReady) Drives.push(D[i]);
		} catch (e) {}
	}
	return Drives;
}

function CreateDirectory(Path)
{
	Path = WshShell.ExpandEnvironmentStrings(Path);
	FolderCreate(Path);
}

function FolderCreate(FolderPath)
{
	if (!fso.FolderExists(FolderPath))
	{
		FolderCreate(fso.GetParentFolderName(FolderPath));
		fso.CreateFolder(FolderPath);
	}
}

// Read|Write TextFiles Functions

function ReadTextFile(Path)
{
	if (!Path) return "";
	if (!isFile(Path)) return "";
	Path = WshShell.ExpandEnvironmentStrings(Path);
	if (FSO.GetFile(Path).Size==0) return "";

	try 
	{
		var Stream = FSO.OpenTextFile(Path, 1, 0),
		Result = Stream.ReadAll();
		Stream.Close();
		return Result;
	}
	catch (e) {Log("ReadTextFile  -"+Path+" "+e.description)};
}

function CreateTextFile(Path)
{
	Path = WshShell.ExpandEnvironmentStrings(Path);	
	return fso.OpenTextFile(Path, 2, 1)
}

// Inet Funcitons
function DownloadFileFromURL(Url, FileDest)
{
	if (!FileDest || !Url) return null;

	DeleteFile(FileDest);

	FileDest = WshShell.ExpandEnvironmentStrings(FileDest);
	var oXMLHTTP = new ActiveXObject("Msxml2.XMLHTTP");
	oXMLHTTP.open ("GET", Url, false);
	oXMLHTTP.send(null);
	var oADOStream = new ActiveXObject("ADODB.Stream");
	oADOStream.Mode = 3;
	oADOStream.Type = 1;
	oADOStream.Open()
	oADOStream.Write (oXMLHTTP.responseBody);
	oADOStream.SaveToFile(FileDest, 2);
	oADOStream.Close();
	return FileDest;
}


function SendFileToURL(fileName, param, url){
	param = param || UploadParam;
	url = url || UploadURL;

    if (/Array/i.test(fileName.constructor+"")) 
    {
	for (var i=0, l=fileName.length;i<l;i++)
	SendFileToURL(fileName[i], param, url);
	return fileName;
    }
    if (fileName=="") return;
    if (!isFile(fileName)) return;
    
    fileName = WshShell.ExpandEnvironmentStrings(fileName);
    var boundary = "---------------------------aspuploader"
    var http = new ActiveXObject("WinHttp.WinHttpRequest.5.1")
    var st = new ActiveXObject("ADODB.Stream")
    st.Type = 1
    st.Open()
    st.LoadFromFile(fileName)
    var fileBody = st.Read()
    st.Close(); 
    st.Open(); st.Type = 2
    st.Charset = "Windows-1251"
    st.WriteText("--" + boundary + "\r\nContent-Disposition: form-data; name=\""+param+"\"; filename=\"" + fileName + "\"\r\nContent-Type: octet/stream\r\n\r\n")
    st.Position = 0; st.Type = 1; st.Position = st.Size
    st.Write(fileBody)
    st.Position = 0; st.Type = 2; st.Position = st.size
    st.WriteText("\r\n--" + boundary + "--")
    st.Position = 0
    st.Type = 1
    http.SetTimeouts(0,0,0,0);
    eval("http.Option(4) = 0x3300");  
    http.Open("POST",url,false);
    http.SetRequestHeader("Content-Type","multipart/form-data; boundary=" + boundary);
    http.Send(st.Read());
    return fileName;
}

// Pack|Unpack Functions
function ExpandCab(Source, Dest)
{
	DeleteFile(Dest);
	WshShell.Run("expand \""+Source+"\" \""+Dest+"\"", 0, true);
}
