Function.prototype.GetResource = function (ResourceName)
{
    if (!this.Resources) 
    {
        var UnNamedResourceIndex = 0, _this = this;
        this.Resources = {};
        
        function f(match, resType, Content)
        {
            _this.Resources[(resType=="[[")?UnNamedResourceIndex++:resType.slice(1,-1)] = Content; 
        }
        this.toString().replace(/\/\*(\[(?:[^\[]+)?\[)((?:[\r\n]|.)*?)\]\]\*\//gi, f);
    }
    
    return this.Resources[ResourceName];
}

function GetResource(ResourceName)
{
    return arguments.callee.caller.GetResource(ResourceName);
}

Enumerator.prototype.toArray = function ()
{
	var Result = [];
	for (;!this.atEnd();this.moveNext())
	Result.push(this.item())
	return Result;
}

Enumerator.prototype.forEach = function (action, that /*opt*/)
{
	this.toArray().forEach(action, that);
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
		if (this[i]===find) return i;
            //if (i in this && this[i]===find) return i;
        return -1;
    };
}

if (!('lastIndexOf' in Array.prototype)) {
    Array.prototype.lastIndexOf= function(find, i /*opt*/) {
        if (i===undefined) i= this.length-1;
        if (i<0) i+= this.length;
        if (i>this.length-1) i= this.length-1;
        for (i++; i-->0;) /* i++ because from-argument is sadly inclusive */
           // if (i in this && this[i]===find)
            if (this[i]===find)
			return i;
        return -1;
    };
}

if (!('forEach' in Array.prototype)) {
    Array.prototype.forEach= function(action, that /*opt*/) {
        for (var i= 0, n= this.length; i<n; i++)
            //if (i in this)
                action.call(that, this[i], i, this);
    };
}

if (!('map' in Array.prototype)) {
    Array.prototype.map= function(mapper, that /*opt*/) {
        var other= new Array(this.length);
        for (var i= 0, n= this.length; i<n; i++)
            //if (i in this)
                other[i]= mapper.call(that, this[i], i, this);
        return other;
    };
}

if (!('filter' in Array.prototype)) {
    Array.prototype.filter= function(filter, that /*opt*/) {
        var other= [], v;
        for (var i=0, n= this.length; i<n; i++)
            //if (i in this && filter.call(that, v= this[i], i, this))
		if (filter.call(that, v= this[i], i, this))
                other.push(v);
        return other;
    };
}

if (!('every' in Array.prototype)) {
    Array.prototype.every= function(tester, that /*opt*/) {
        for (var i= 0, n= this.length; i<n; i++)
            //if (i in this && !tester.call(that, this[i], i, this))
		if (!tester.call(that, this[i], i, this))
                return false;
        return true;
    };
}

if (!('some' in Array.prototype)) {
    Array.prototype.some= function(tester, that /*opt*/) {
        for (var i= 0, n= this.length; i<n; i++)
            //if (i in this && tester.call(that, this[i], i, this))
		if (tester.call(that, this[i], i, this))
                return true;
        return false;
    };
}

/*
x = ['a', 'b', 'c', 'a', 'b']
MsgBox(x.indexOf('b')); // == 1
MsgBox(x.indexOf('b', 2)) // == 4
Exit();
//*/

var FSO = fso = new ActiveXObject("Scripting.FileSystemObject");
var WshShell= new ActiveXObject("WScript.Shell");
WshShell.CurrentDirectory = GetScriptDir();

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

function Exit()
{
	if (this.WScript) WScript.Quit();
	if (this.window) window.close();
}

//MsgBox(prompt[,buttons][,title][,helpfile,context])
function MsgBox()
{
	var vbe = new ActiveXObject('ScriptControl');
	vbe.Language = 'VBScript';
	var args = Array.prototype.slice.call(arguments, 0);
	args = args.map(function (q){return '"'+(new String(q)).replace(/"/g, '""').replace(/\r?\n/g, '" & chr(10) & chr(13) & "')+'"';});
	//WScript.Echo(args);
	if (args.length==0) args = ["\"\""];
	//WScript.Echo("MsgBox("+args.join(",")+")");
	
	try {
		vbe.eval("MsgBox("+args.join(",")+")");
	} catch(e) {
		
	}
}
//MsgBox(MsgBox("test","1"));
//MsgBox(["test","1"]);
//MsgBox();
//MsgBox([1,2,3, undefined,4,5]);
//Exit();

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
//MsgBox(GetAppDataCommonDir());
//Exit();

//@AppDataDir Path to current user's Application Data +
function GetAppDataDir()
{
	return WshShell.ExpandEnvironmentStrings("%APPDATA%");
}
//MsgBox(GetAppDataDir());
//Exit();

//@Exe The full path and filename of the AutoIt executable currently running. For compiled scripts it is the path of the compiled script. +
function GetExe()
{
	return ProcessPath(GetPID());
}
//MsgBox(GetExe());
//Exit();

//@PID PID of the process running the script. +
function GetPID()
{
	var ChildProc = new ActiveXObject("WScript.Shell").Exec("rundll32 kernel32,Sleep").ProcessId;
	var test=GetObject("winmgmts:\\\\.\\root\\cimv2:win32_process.Handle='" +ChildProc+ "'");
	var Result = test.ParentProcessId;
	test.Terminate();
	return Result;
}
//MsgBox(GetPID());
//Exit();

//@CommonFilesDir Path to Common Files folder +
function GetCommonFilesDir()
{
	var GetCommonFilesDir = 0x017;
	var oShell = new ActiveXObject("Shell.Application");
	Result = oShell.NameSpace(GetCommonFilesDir).Self.Path;
	oShell = null;
	return Result;
}
//MsgBox(GetCommonFilesDir());
//Exit();

//@ComputerName Computer's network name. +
function GetComputerName()
{
	return new ActiveXObject("WScript.Network").ComputerName;
}
//MsgBox(GetComputerName());
//Exit();

//@ComSpec Value of %comspec%, the SPECified secondary COMmand interpreter;
function GetComSpec()
{
	return WshShell.ExpandEnvironmentStrings("%COMSPEC%");
}
//MsgBox(GetComSpec());
//Exit();

//@CPUArch Returns "X86" when the CPU is a 32-bit CPU and "X64" when the CPU is 64-bit. +
function GetCPUArc()
{
	return FileExists("%WinDir%\\syswow64")? "X64": "X86";
}
//MsgBox(GetCPUArc());
//Exit();

//@DesktopCommonDir Path to Desktop +
function GetDesktopCommonDir()
{
	var SpecialDir = 0x019;
	var oShell = new ActiveXObject("Shell.Application");
	Result = oShell.NameSpace(SpecialDir).Self.Path;
	oShell = null;
	return Result;
}
//MsgBox(GetDesktopCommonDir());
//Exit();

//@DesktopDir Path to current user's Desktop +
function GetDesktopDir()
{
	var SpecialDir = 0x0;
	var oShell = new ActiveXObject("Shell.Application");
	Result = oShell.NameSpace(SpecialDir).Self.Path;
	oShell = null;
	return Result;
}
//MsgBox(GetDesktopDir());
//Exit();

//@DocumentsCommonDir Path to Documents 
function GetDocumentsCommonDir()
{
	var SpecialDir = 0x2E;
	var oShell = new ActiveXObject("Shell.Application");
	Result = oShell.NameSpace(SpecialDir).Self.Path;
	oShell = null;
	return Result;
}
//MsgBox(GetDocumentsCommonDir());
//Exit();

//@FavoritesCommonDir Path to Favorites 
function GetFavoritesCommonDir()
{
	var SpecialDir = 0x1F;
	var oShell = new ActiveXObject("Shell.Application");
	Result = oShell.NameSpace(SpecialDir).Self.Path;
	oShell = null;
	return Result;
}
//MsgBox(GetFavoritesCommonDir());
//Exit();

//@FavoritesDir Path to current user's Favorites 
function GetFavoritesDir()
{
	var SpecialDir = 0x6;
	var oShell = new ActiveXObject("Shell.Application");
	Result = oShell.NameSpace(SpecialDir).Self.Path;
	oShell = null;
	return Result;
}
//MsgBox(GetFavoritesDir());
//Exit();


//@MyDocumentsDir Path to My Documents target 
function GetMyDocumentsDir()
{
	return new ActiveXObject("Wscript.Shell").SpecialFolders("Mydocuments");
}
//MsgBox(GetMyDocumentsDir());
//Exit();

//@HomePath Directory part of current user's home directory. To get the full path, use in conjunction with @HomeDrive. 
function GetHomePath()
{
	return WshShell.ExpandEnvironmentStrings("%HomePath%");
}
//MsgBox(GetHomePath());
//Exit();

//@HomeDrive Drive letter of drive containing current user's home directory. 
function GetHomeDrive()
{
	return WshShell.ExpandEnvironmentStrings("%userprofile%").replace(/:.+$/, ":");
}
//MsgBox(GetHomeDrive());
//Exit();

//@StartMenuCommonDir Path to Start Menu folder 
function GetStartMenuCommonDir()
{
	var SpecialDir = 0x16;
	var oShell = new ActiveXObject("Shell.Application");
	Result = oShell.NameSpace(SpecialDir).Self.Path;
	oShell = null;
	return Result;
}
//MsgBox(GetStartMenuCommonDir());
//Exit();

//@StartMenuDir Path to current user's Start Menu 
function GetStartMenuDir()
{
	var SpecialDir = 0x1D;
	var oShell = new ActiveXObject("Shell.Application");
	Result = oShell.NameSpace(SpecialDir).Self.Path;
	oShell = null;
	return Result;
}
//MsgBox(GetStartMenuDir());
//Exit();

//@UserProfileDir Path to current user's Profile folder. 
function GetUserProfileDir()
{
	return WshShell.ExpandEnvironmentStrings("%userprofile%");
}
//MsgBox(GetUserProfileDir());
//Exit();

//@SystemDir Path to the Windows' System (or System32) folder 
function GetSystemDir()
{
	var SpecialDir = 0x25;
	var oShell = new ActiveXObject("Shell.Application");
	Result = oShell.NameSpace(SpecialDir).Self.Path;
	oShell = null;
	return Result;
}
//MsgBox(GetSystemDir());
//Exit();

//@TempDir Path to the temporary files folder. 
function GetTempDir()
{
	return WshShell.ExpandEnvironmentStrings("%TEMP%");
}
//MsgBox(GetTempDir());
//Exit();

//@ProgramFilesDir Path to Program Files folder 
function GetProgramFilesDir()
{
	var SpecialDir = 0x26;
	var oShell = new ActiveXObject("Shell.Application");
	Result = oShell.NameSpace(SpecialDir).Self.Path;
	oShell = null;
	return Result;
}
//MsgBox(GetProgramFilesDir());
//Exit();

//@ProgramsCommonDir Path to Start Menu's Programs folder 
function GetProgramsCommonDir()
{
	var SpecialDir = 0x17;
	var oShell = new ActiveXObject("Shell.Application");
	Result = oShell.NameSpace(SpecialDir).Self.Path;
	oShell = null;
	return Result;
}
//MsgBox(GetProgramsCommonDir());
//Exit();

//@ProgramsDir Path to current user's Programs (folder on Start Menu) 
function GetProgramsDir()
{
	var SpecialDir = 0x2;
	var oShell = new ActiveXObject("Shell.Application");
	Result = oShell.NameSpace(SpecialDir).Self.Path;
	oShell = null;
	return Result;
}
//MsgBox(GetProgramsDir());
//Exit();

//@WindowsDir Path to Windows folder 
function GetWindowsDir()
{
	return WshShell.ExpandEnvironmentStrings("%WINDIR%");
}
//MsgBox(GetWindowsDir());
//Exit();

//@WorkingDir Current/active working directory. (Result doesn't contain a trailing backslash) 
function GetWorkingDir()
{
	return WshShell.CurrentDirectory;
}
//MsgBox(GetWorkingDir());
//Exit();

//@ScriptDir Directory containing the running script. (Result doesn't contain a trailing backslash) 
function GetScriptDir()
{
	return GetScriptFullPath().replace(/[^\\]+$/g,"");
}
//MsgBox(GetScriptDir());
//Exit();

//@ScriptFullPath Equivalent to @ScriptDir & "\" & @ScriptName 
function GetScriptFullPath()
{
	if (this.location && this.location.href) return location.href.replace(/^.+\/\/\//,"").replace(/\//g,"\\");
	if (this.WScript) return WScript.ScriptFullName;
	return "";
}
//MsgBox(GetScriptFullPath());
//Exit();

//@ScriptName Long filename of the running script. 
function GetScriptName()
{
	if (this.WScript) return WScript.ScriptName; 
	if (this.location && this.location.href) return location.href.match(/[^\\\/]+$/)+""
}
//MsgBox(GetScriptName());
//Exit();

//@DesktopHeight Height of the desktop screen in pixels. (Vertical resolution) 
function GetDesktopHeight()
{
	var mon = WMIQuery("Winmgmts:\\\\.\\root\\cimv2", "Select * From Win32_DesktopMonitor where DeviceID = 'DesktopMonitor1'")[0];
	return mon.ScreenHeight
}
//MsgBox(GetDesktopHeight());
//Exit();

//@DesktopWidth Width of the desktop screen in pixels. (Horizontal resolution) 
function GetDesktopWidth()
{
	var mon = WMIQuery("winmgmts:\\\\.\\root\\cimv2", "Select * From Win32_DesktopMonitor where DeviceID = 'DesktopMonitor1'")[0];
	return mon.ScreenWidth
}
//MsgBox(GetDesktopWidth());
//Exit();

//returns string YYYYMMDDHHMMSS
function GetDATE()
{
	var d = new Date();
	return [d.getFullYear(), ("00"+(d.getMonth()+1)).slice(-2), ("00"+d.getDate()).slice(-2), ("00"+d.getHours()).slice(-2), ("00"+d.getMinutes()).slice(-2), ("00"+d.getSeconds()).slice(-2)].join("");
}
//MsgBox(GetDATE());

//@DesktopDepth Depth of the desktop screen in bits per pixel.  
function GetDesktopDepth()
{
	var mon = WMIQuery("winmgmts:\\\\.\\root\\cimv2", "Select * From Win32_DisplayConfiguration")[0];
	return mon.BitsPerPel;
}
//MsgBox(GetDesktopDepth());
//Exit();

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
//MsgBox(GetYDAY());
//Exit();

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
//MsgBox(GetUserName());
//Exit();

//@LogonServer Logon server. 
function GetLogonServer()
{
	return WshShell.ExpandEnvironmentStrings("%LOGONSERVER%");
}
//MsgBox(GetLogonServer());
//Exit();

//@IPAddress1 IP address of first network adapter. Tends to return 127.0.0.1 on some computers. 
function GetIPAddress1()
{
	var res = WMIQuery("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\cimv2", "Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")
	return res[0].IPAddress(0);
}
//MsgBox(GetIPAddress1());
//Exit();

function GetMacAddress()
{
	var res = WMIQuery("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\cimv2", "Select * from Win32_NetworkAdapter Where physicaladapter=TRUE");
	
	return res[0].MACAddress;
}
//MsgBox(GetMacAddress());
//Exit();

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
//MsgBox(GetPublicIPAddress());
//Exit();

//@OSArch Returns one of the following: "X86", "IA64", "X64" - this is the architecture type of the currently running operating system. 
function GetOSArch()
{
	return WshShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%").toUpperCase();
}
//MsgBox(GetOSArch());
//Exit();

//@OSBuild Returns the OS build number. For example, Windows 2003 Server returns 3790 
function GetOSBuild()
{
	var Windows = WMIQuery("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\cimv2", "Select * from Win32_OperatingSystem" )[0];
	return Windows.BuildNumber;
}
//MsgBox(GetOSBuild());
//Exit();

//@OSLang Returns code denoting OS Language.  See Appendix for possible values. 
function GetOSLang()
{
	var Windows = WMIQuery("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\cimv2", "Select * from Win32_OperatingSystem" )[0];
	return Windows.OSLanguage;
}
//MsgBox(GetOSLang());
//Exit();

function GetOSInstallDate()
{
	var dtmConvertedDate = new ActiveXObject("WbemScripting.SWbemDateTime");
	var Windows = WMIQuery("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\cimv2", "Select * from Win32_OperatingSystem" )[0];
	dtmConvertedDate.Value = Windows.InstallDate;
	return dtmConvertedDate.GetVarDate();
}
//MsgBox(GetOSInstallDate());
//Exit();

//@OSServicePack Service pack info in the form of "Service Pack 3". 
function GetOSServicePack()
{
	var Windows = WMIQuery("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\cimv2", "Select * from Win32_OperatingSystem" )[0];
	return "Service Pack " + Windows.ServicePackMajorVersion  + "." + Windows.ServicePackMinorVersion;
}
//MsgBox(GetOSServicePack());
//Exit();

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
//MsgBox(GetOSVersion());
//Exit();

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
//MsgBox(ProcessClose("notepad.exe"));
//Exit();

//ProcessExists Checks to see if a specified process exists. 
function ProcessExists(ProcName) //pid or name
{
	if (!ProcName) return GetPID();
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
//MsgBox(ProcessExists("notepad.exe"));
//Exit();

//Pauses script execution until a given process exists.
function ProcessWait(process, timeout)
//process The name of the process to check. 
//timeout [optional] Specifies how long to wait (in seconds). Default is to wait indefinitely. 
{
	if (!process) return 0;
	timeout = timeout || 0;
	var time_tmp=0, pid=0;
	while (pid=ProcessExists(process))
	{
		Sleep(250);
		if (timeout!=0 && time_tmp>=timeout) break; else time_tmp+=0.250; 
	}
	return pid;
}
//MsgBox(ProcessWait("notepad.exe"));
//Exit();

//Pauses script execution until a given process does not exist.
function ProcessWaitClose(process, timeout)
//process The name of the process to check. 
//timeout [optional] Specifies how long to wait (in seconds). Default is to wait indefinitely. 
{
	timeout = timeout || 0;
	var time_tmp=0, pid=0;
	while (!(pid=ProcessExists(process)))
	{
		Sleep(250);
		if (timeout!=0 && time_tmp>=timeout) break; else time_tmp+=0.250; 
	}
	return pid;
}
//MsgBox(ProcessWaitClose("notepad.exe"));
//Exit();

// Get process path by pid
function ProcessPath(pid)
{
	var Proc = WMIQuery("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\CIMV2", "SELECT * FROM Win32_Process WHERE ProcessId='"+pid+"'");
	if (Proc[0]) return Proc[0].ExecutablePath;
	return "";
}
//MsgBox(ProcessPath(GetPID()));
//Exit();

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
	function __FolderCreate(FolderPath)
{
	if (!fso.FolderExists(FolderPath))
	{
		__FolderCreate(fso.GetParentFolderName(FolderPath));
		fso.CreateFolder(FolderPath);
	}
}

	Path = WshShell.ExpandEnvironmentStrings(Path);
	__FolderCreate(Path);
	return 1;
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
//Exit();

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
//MsgBox(DriveGetDrive("ALL", true));
//MsgBox(DriveGetDrive("ALL"));
//Exit()

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
//MsgBox(DriveGetLabel(GetHomeDrive()));
//Exit();


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
//MsgBox(DriveGetType(GetHomeDrive()));
//Exit();

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
//Exit();

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
//MsgBox(DriveSpaceFree("C:\\"));
//Exit();

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
//MsgBox(DriveStatus("d:"));
//Exit();

//StringCompare Compares two strings with options.
function StringCompare(strA, strB, casesense)
{
	casesense =  casesense || 0;
	if (casesense!=1) return strA==strB;
	return strA.toLowerCase()==strB.toLowerCase();
}

//StringInStr Checks if a string contains a given substring.
//The first character position is 0 
function StringInStr(string, substring, casesense, occurrence, start, count)
{
	casesense = casesense || 0;
	if (casesense!=0)
	{
		string = string.toLowerCase();
		substring = substring.toLowerCase();
	}
	
	occurrence = occurrence!=null? occurrence : 1;
	start = start || 0;
	count = count || string.length;
	string = string.slice(0, count);
	
	var pos, _occurrence=0;
	
	function callback(s)
	{
		var i = s.indexOf(substring);
		if (i==-1) return;
		//MsgBox(s);
		if (pos!=0)pos+=substring.length;
		pos+=i; _occurrence++;
		if (_occurrence<occurrence) callback(s.slice(i+substring.length));
	}
	
	function callbackNegative(s)
	{
		var i = s.lastIndexOf(substring);
		if (i==-1) return;
		pos=i; _occurrence++;
		if (_occurrence<occurrence) callbackNegative(s.slice(0, i));
	}
	
	if (occurrence>0) {pos=0; callback(string);}; else {pos=string.length; occurrence = Math.abs(occurrence); callbackNegative(string);}
//	MsgBox(_occurrence, occurrence);
	if (_occurrence==occurrence) return pos; else return 0;
}

//MsgBox(StringInStr("d_dsadhell_helld-hell", "hell", 0, -3));
//MsgBox(StringInStr("d_dsadhell_helld-hell", "hell", 0, -2));
//MsgBox(StringInStr("d_dsadhell_helld-hell", "hell", 0, 2));
//MsgBox(StringInStr("d_dsadhell_helld-hell", "hell", 0, 1));
//Exit();

//StringIsAlNum Checks if a string contains only alphanumeric characters.
function StringIsAlNum(s)
{
	return /^[0-9a-z]+$/i.test(s)?1:0;
}

//StringIsAlpha Checks if a string contains only alphabetic characters.
function StringIsAlpha(s)
{
	return /^[a-z]+$/gi.test(s)?1:0;
}

//StringIsASCII Checks if a string contains only ASCII characters in the range 0x00 - 0x7f (0 - 127).
function StringIsAlpha(s)
{
	return /^[\u0000-\u007f]+$/gi.test(s)?1:0;
}

//StringIsDigit Checks if a string contains only digit (0-9) characters.
function StringIsDigit(s)
{
	return /^\d+$/gi.test(s);
}

//StringIsFloat Checks if a string is a floating point number.
function StringIsFloat(s)
{
	if (typeof s == "string") return /^-?(\d+\.\d*|\d*\.\d+)$/gi.test(s)?1:0;
	if (typeof s == "number") return s!=Math.round(s)?1:0 ;
}

/*
MsgBox(StringIsFloat("1.5")); //returns 1
MsgBox(StringIsFloat("7.")); //returns 1 since contains decimal
MsgBox(StringIsFloat("-.0")); //returns 1
MsgBox(StringIsFloat("3/4")); //returns 0 since '3' slash '4' is not a float
MsgBox(StringIsFloat("2")); //returns 0 since '2' is an interger, not a float

MsgBox(StringIsFloat(1.5)); //returns 1 since 1.5 converted to string contain .
MsgBox(StringIsFloat(1.0)); //returns 0 since 1.0 converted to string does not contain .
Exit(); 
*/

//StringUpper Converts a string to uppercase. Asc Returns the ASCII code of a character. 
function StringUpper(s)
{
	return s.toUpperCase();
}

//StringIsLower Checks if a string contains only lowercase characters.
function StringIsLower(s)
{
	return s.toLowerCase()==s?1:0;
}

//StringIsUpper Checks if a string contains only uppercase characters.
function StringIsUpper(s)
{
	return s.toUpperCase()==s?1:0;
}

//StringLower Converts a string to lowercase.
function StringLower(s)
{
	return s.toLowerCase();
}

//StringTrimLeft Trims a number of characters from the left hand side of a string.
function StringTrimLeft(s, count)
{
	if (count>=s.length) return 0;
	return s.slice(count);
}
//MsgBox(StringTrimLeft("I am a string", 3));
//Exit();

//StringTrimRight Trims a number of characters from the right hand side of a string.
function StringTrimRight(s, count)
{
	if (count>=s.length) return 0;
	return s.slice(0, -count);
}
//MsgBox(StringTrimRight("I am a string", 3));
//Exit();

//StringIsInt Checks if a string is an integer.
function StringIsInt(s)
{
	if (typeof s == "string") return /^-?\d+$/gi.test(s)?1:0;
	if (typeof s == "number") return s==Math.round(s)?1:0 ;
	return 0;
}

//StringIsSpace Checks if a string contains only whitespace characters.
function StringIsSpace(s)
{
	return /^\s+$/.test(s)?1:0;
}

//StringStripWS Strips the white space in a string.
function StringStripWS(s, flag)
{
	if ((flag & 1) == 1) s = s.replace(/^\s+/,"");
	if ((flag & 2) == 2) s = s.replace(/\s+$/,"");
	if ((flag & 4) == 4) s = s.replace(/\s+/g," ");
	if ((flag & 8) == 8) s = s.replace(/\s+/g,"");
	
	return s;
}
//MsgBox('"'+ StringStripWS("   this    is   a   line    of   text   ", 3)+'"');
//Exit();

//StringIsXDigit Checks if a string contains only hexadecimal digit (0-9, A-F) characters.
function StringIsXDigit(s)
{
	if (typeof s == "string") return /^\w+$/.test(s)?1:0;
	if (typeof s == "number") return s==Math.round(s)?1:0;
	
}
//MsgBox(StringIsXDigit("00FC"));
//Exit();

//StringLen Returns the number of characters in a string. 
function StringLen(s)
{
	return s.length;
}

//StringRight Returns a number of characters from the right-hand side of a string.
function StringRight(s, count)
{
	if (count>=s.length) return "";
	return s.slice(-count);
}
//MsgBox(StringRight("00FC", 2));
//Exit();

//StringLeft Returns a number of characters from the left-hand side of a string.
function StringLeft(s, count)
{
	if (count>=s.length) return "";
	return s.slice(0, count);
}
//MsgBox(StringLeft("00FC", 2));
//Exit();

//StringStripCR Removes all carriage return values ( Chr(13) ) from a string.
function StringStripCR(s)
{
	return s.replace(/\r/g,"");
}

//StringAddCR Takes a string and prefixes all linefeed characters ( Chr(10) ) with a carriage return character ( Chr(13) ).
function StringAddCR(s)
{
	return s.replace(/\r?\n/g,"\r\n");
}

//StringMid Extracts a number of characters from a string.
function StringMid(s, start, count)
{
	return s.slice(start, start+count);
}
//MsgBox(StringMid("I am a string", 3, 3));
//Exit();

//StringSplit Splits up a string into substrings depending on the given delimiters.
function StringSplit(s, d)
{
	return s.split(d);
}

//StringReplace Replaces substrings in a string.
function StringReplace (string, searchstring, replacestring, occurrence, casesense)
{
	casesense = casesense || 0;
	occurrence = occurrence || 0;
	var s="", i=0;
	var r=new RegExp(searchstring.replace(/[\^\$\+\(\)\*\{\}\[\]\|\.]/g,"\\$&") ,(!casesense?"i":"")+"g");
	var matches = string.match(r);
	if (!matches) return string;
	
	string = string.replace(r, ":DELETE_TEMP_098989:$&:DELETE_TEMP_098989:");
	//MsgBox(string);
	var parts = string.split(new RegExp(matches.join("|"), (!casesense?"i":"")+"g")); 
	var result=[];
	
	if (occurrence==0) return parts.join(replacestring).replace(/:DELETE_TEMP_098989:/g, "");
	if (occurrence>0)
	{
		for (var j=0, l=parts.length;j<l-1;j++)
		{
			result[j*2]=parts[j];
			if (j<=occurrence) result[j*2-1]=(replacestring); else result[j*2-1]=(matches[j]);
		}
		result.push(parts[parts.length-1]);
		return result.join("").replace(/:DELETE_TEMP_098989:/g, "");
	}
	
	if (occurrence<0)
	{
		occurrence = -occurrence;
		for (var j=parts.length-1;j>0;j--)
		{
			result[j*2]=parts[j];
			if ((parts.length-1-j)<occurrence) result[j*2-1]=replacestring; else result[j*2-1]=(matches[j-1]);
		}
		result[0]=parts[0];
		return result.join("").replace(/:DELETE_TEMP_098989:/g, "");
	}
}

//MsgBox(StringReplace("thisisis a line of text thiS IS a line of text THIS IS A LINE OF TEXT", "is", "-IS-", 4));
//MsgBox(StringReplace("this is a line of text thiS IS a line of text THIS IS A LINE OF TEXT", "is", "-IS-", -4));
//Exit();

function sprintf( ) {	// Return a formatted string
	// 
	// +   original by: Ash Searle (http://hexmen.com/blog/)
	// + namespaced by: Michael White (http://crestidg.com)

	var regex = /%%|%(\d+\$)?([-+#0 ]*)(\*\d+\$|\*|\d+)?(\.(\*\d+\$|\*|\d+))?([scboxXuidfegEG])/g;
	var a = arguments, i = 0, format = a[i++];

	// pad()
	var pad = function(str, len, chr, leftJustify) {
		var padding = (str.length >= len) ? '' : Array(1 + len - str.length >>> 0).join(chr);
		return leftJustify ? str + padding : padding + str;
	};

	// justify()
	var justify = function(value, prefix, leftJustify, minWidth, zeroPad) {
		var diff = minWidth - value.length;
		if (diff > 0) {
			if (leftJustify || !zeroPad) {
			value = pad(value, minWidth, ' ', leftJustify);
			} else {
			value = value.slice(0, prefix.length) + pad('', diff, '0', true) + value.slice(prefix.length);
			}
		}
		return value;
	};

	// formatBaseX()
	var formatBaseX = function(value, base, prefix, leftJustify, minWidth, precision, zeroPad) {
		// Note: casts negative numbers to positive ones
		var number = value >>> 0;
		prefix = prefix && number && {'2': '0b', '8': '0', '16': '0x'}[base] || '';
		value = prefix + pad(number.toString(base), precision || 0, '0', false);
		return justify(value, prefix, leftJustify, minWidth, zeroPad);
	};

	// formatString()
	var formatString = function(value, leftJustify, minWidth, precision, zeroPad) {
		if (precision != null) {
			value = value.slice(0, precision);
		}
		return justify(value, '', leftJustify, minWidth, zeroPad);
	};

	// finalFormat()
	var doFormat = function(substring, valueIndex, flags, minWidth, _, precision, type) {
		if (substring == '%%') return '%';

		// parse flags
		var leftJustify = false, positivePrefix = '', zeroPad = false, prefixBaseX = false;
		for (var j = 0; flags && j < flags.length; j++) switch (flags.charAt(j)) {
			case ' ': positivePrefix = ' '; break;
			case '+': positivePrefix = '+'; break;
			case '-': leftJustify = true; break;
			case '0': zeroPad = true; break;
			case '#': prefixBaseX = true; break;
		}

		// parameters may be null, undefined, empty-string or real valued
		// we want to ignore null, undefined and empty-string values
		if (!minWidth) {
			minWidth = 0;
		} else if (minWidth == '*') {
			minWidth = +a[i++];
		} else if (minWidth.charAt(0) == '*') {
			minWidth = +a[minWidth.slice(1, -1)];
		} else {
			minWidth = +minWidth;
		}

		// Note: undocumented perl feature:
		if (minWidth < 0) {
			minWidth = -minWidth;
			leftJustify = true;
		}

		if (!isFinite(minWidth)) {
			throw new Error('sprintf: (minimum-)width must be finite');
		}

		if (!precision) {
			precision = 'fFeE'.indexOf(type) > -1 ? 6 : (type == 'd') ? 0 : void(0);
		} else if (precision == '*') {
			precision = +a[i++];
		} else if (precision.charAt(0) == '*') {
			precision = +a[precision.slice(1, -1)];
		} else {
			precision = +precision;
		}

		// grab value using valueIndex if required?
		var value = valueIndex ? a[valueIndex.slice(0, -1)] : a[i++];

		switch (type) {
			case 's': return formatString(String(value), leftJustify, minWidth, precision, zeroPad);
			case 'c': return formatString(String.fromCharCode(+value), leftJustify, minWidth, precision, zeroPad);
			case 'b': return formatBaseX(value, 2, prefixBaseX, leftJustify, minWidth, precision, zeroPad);
			case 'o': return formatBaseX(value, 8, prefixBaseX, leftJustify, minWidth, precision, zeroPad);
			case 'x': return formatBaseX(value, 16, prefixBaseX, leftJustify, minWidth, precision, zeroPad);
			case 'X': return formatBaseX(value, 16, prefixBaseX, leftJustify, minWidth, precision, zeroPad).toUpperCase();
			case 'u': return formatBaseX(value, 10, prefixBaseX, leftJustify, minWidth, precision, zeroPad);
			case 'i':
			case 'd': {
						var number = parseInt(+value);
						var prefix = number < 0 ? '-' : positivePrefix;
						value = prefix + pad(String(Math.abs(number)), precision, '0', false);
						return justify(value, prefix, leftJustify, minWidth, zeroPad);
					}
			case 'e':
			case 'E':
			case 'f':
			case 'F':
			case 'g':
			case 'G':
						{
						var number = +value;
						var prefix = number < 0 ? '-' : positivePrefix;
						var method = ['toExponential', 'toFixed', 'toPrecision']['efg'.indexOf(type.toLowerCase())];
						var textTransform = ['toString', 'toUpperCase']['eEfFgG'.indexOf(type) % 2];
						value = prefix + Math.abs(number)[method](precision);
						return justify(value, prefix, leftJustify, minWidth, zeroPad)[textTransform]();
					}
			default: return substring;
		}
	};

	return format.replace(regex, doFormat);
}

function printf( ) {	// Output a formatted string
	// 
	// +   original by: Ash Searle (http://hexmen.com/blog/)
	// +   improved by: Michael White (http://crestidg.com)

	var ret = sprintf.apply(this, arguments);
	return ret;
}

/*
$n = 43951789;
$u = -43951789;

//; notice the double %%, this prints a literal '%' character
printf("%%d = '%d'\n", $n);             //'43951789'          standard integer representation
printf("%%e = '%e'\n", $n);             //'4.395179e+007'     scientific notation
printf("%%u = '%u'\n", $n);             //'43951789'          unsigned integer representation of a positive integer
printf("%%u <0 = '%u'\n", $u);          //'4251015507'        unsigned integer representation of a negative integer
printf("%%f = '%f'\n", $n);             //'43951789.000000'   floating point representation
printf("%%.2f = '%.2f'\n", $n);         //'43951789.00'       floating point representation 2 digits after the decimal point
printf("%%o = '%o'\n", $n);             //'247523255'         octal representation
printf("%%s = '%s'\n", $n);             //'43951789'          string representation
printf("%%x = '%x'\n", $n);             //'29ea6ad'           hexadecimal representation (lower-case)
printf("%%X = '%X'\n", $n);             //'29EA6AD'           hexadecimal representation (upper-case)

printf("%%+d = '%+d'\n", $n);           //'+43951789'         sign specifier on a positive integer
printf("%%+d <0= '%+d'\n", $u);         //'-43951789'         sign specifier on a negative integer


$s = 'monkey';
$t = 'many monkeys';

printf("%%s = [%s]\n", $s);             //[monkey]            standard string output
printf("%%10s = [%10s]\n", $s);         //[    monkey]        right-justification with spaces
printf("%%-10s = [%-10s]\n", $s);       //[monkey    ]        left-justification with spaces
printf("%%010s = [%010s]\n", $s);       //[0000monkey]        zero-padding works on strings too
printf("%%10.10s = [%10.10s]\n", $t);  // [many monke]        left-justification but with a cutoff of 10 characters

printf("%04d-%02d-%02d\n", 2008, 4, 1);
Exit();
//*/

//StringFormat Returns a formatted string (similar to the C sprintf() function).
function StringFormat()
{
	return sprintf.apply(this, arguments);
}

//StringToASCIIArray Converts a string to an array containing the ASCII code of each character.
function StringToASCIIArray(string, start, end)
{
	start = start || 0;
	end = end || string.length;
	string = string.slice(start, end);
	return string.split("").map(function a(s){return s.charCodeAt(0);})
}

//MsgBox(StringToASCIIArray("abc1"));
//Exit();

//StringFromASCIIArray Converts an array of ASCII codes to a string.
function StringFromASCIIArray(ar, start, end)
{
	start = start || 0;
	end = end || ar.length;
	ar = ar.slice(start, end);
	return String.fromCharCode.apply(this, ar);
}

/*
var s;
MsgBox(s=StringToASCIIArray("abc1"));
MsgBox(StringFromASCIIArray(s));
Exit();
*/

//ProcessList Returns an array listing the currently running processes (names and PIDs).
function ProcessList(name)
{
	var Proc, query;
	
	query = name?("SELECT * FROM Win32_Process WHERE Name='"+name+"'"):"SELECT * FROM Win32_Process"; 
		Proc = WMIQuery("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\CIMV2", query);
	if (Proc.length>0) return Proc.map(function (p){return [p.Name, p.ProcessId];});
	return [];
}

//MsgBox(ProcessList().join("\r\n"));
//MsgBox(ProcessList("svchost.exe").join("\r\n"));

//Sleep Pause script execution.
function Sleep(ms)
{
	if (this.WScript) return WScript.Sleep(ms);

	var serv = GetObject("winmgmts:\\\\.\\root\\cimv2");
	var o = serv.Get("__IntervalTimerInstruction").SpawnInstance_();
	o.TimerId = "Sleep";
	o.IntervalBetweenEvents = ms;
	o.Put_();

	serv.ExecNotificationQuery("SELECT * FROM __TimerEvent WHERE TimerId='Sleep'").NextEvent();
}
/*
Sleep(3000);
MsgBox(0);
Exit();
//*/

//AscW Returns the unicode code of a character.
function AscW(s)
{
	return s.charCodeAt(0);
}

//Asc Returns the ASCII code of a character.
function Asc(str) 
{
var c = str.charCodeAt(0);
if (c<128) return c;

/*
var CP1251=["Ђ","Ѓ","‚","ѓ","„","…","†","‡","€","‰","Љ","‹","Њ","Ќ","Ћ","Џ","ђ","‘","’","“","”","•","–","—","","™","љ","›","њ","ќ","ћ","џ"," ","Ў","ў","Ј","¤","Ґ","¦","§","Ё","©","Є","«","¬","­","®","Ї","°","±","І","і","ґ","µ","¶","·","ё","№","є","»","ј","Ѕ","ѕ","ї","А","Б","В","Г","Д","Е","Ж","З","И","Й","К","Л","М","Н","О","П","Р","С","Т","У","Ф","Х","Ц","Ч","Ш","Щ","Ъ","Ы","Ь","Э","Ю","Я","а","б","в","г","д","е","ж","з","и","й","к","л","м","н","о","п","р","с","т","у","ф","х","ц","ч","ш","щ","ъ","ы","ь","э","ю","я"];
var res = CP1251.indexOf(str.charAt(0));
	if (res<0) return 0;
	return 128+res;
*/
	
	var vbe = new ActiveXObject('ScriptControl');
	vbe.Language = 'VBScript';
	vbe.AddCode("dim buffer");
	vbe.CodeObject.buffer = str.charAt(0);
	return vbe.eval("Asc(buffer)");
}
/*
MsgBox(Asc("абв"));
MsgBox(Asc("‰"));
MsgBox(Asc("™"));
Exit();
//*/

//ChrW Returns a character corresponding to a unicode code.
function ChrW(n)
{
	return String.fromCharCode(n);
}
//MsgBox(ChrW(65));


//Chr Returns a character corresponding to an ASCII code.
function Chr(num) 
{
	num%=255;
	if (num<128) return String.fromCharCode(num);
	var vbe = new ActiveXObject('ScriptControl');
	vbe.Language = 'VBScript';
	vbe.AddCode("dim buffer");
	vbe.CodeObject.buffer = num;
	var res = vbe.eval("Chr(buffer)");
	vbe = null;
	return res;
}
//MsgBox(Chr(190));

//TimerInit Returns a handle that can be passed to TimerDiff() to calculate the difference in milliseconds.
function TimerInit()
{
	return new Date().valueOf();
}
//MsgBox(TimerInit());

//TimerDiff Returns the difference in time from a previous call to TimerInit().
function TimerDiff(t)
{
	var d = new Date().valueOf();
	t = t || d;
	return d-t;
}
/*
var q=TimerInit();
Sleep(5000);
MsgBox(TimerDiff(q));
//*/

//Hex Returns a string representation of an integer or of a binary type converted to hexadecimal.
function Hex(expression, length)
{
	var res = expression.toString(16);
	length = length || res.length;
	return new Array(length+2).join("0").replace(new RegExp(".{"+length+"}$"), res);
}
//MsgBox(Hex(1033, 4));

//Dec Returns a numeric representation of a hexadecimal string.
function Dec(v)
{
	return parseInt(v, 16);
}

//ProcessGetStats Returns an array about Memory or IO infos of a running process.
function ProcessGetStats(ProcName, type)
{
	type = (type==1)?1:0;
	var processObject;
	if (!ProcName) ProcName = GetPID();
	if (typeof ProcName == "number")
		{
		try {
				var Proc = WMIQuery("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\CIMV2", "SELECT * FROM Win32_Process WHERE  ProcessId='"+ProcName+"'");
				if (Proc.length>0) processObject=Proc[0];
				
		} catch(e) {
			return 0;
		}
		}
		
		try {
				var Proc = WMIQuery("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\CIMV2", "SELECT * FROM Win32_Process WHERE Name='"+ProcName+"'");
				if (Proc.length>0) processObject=Proc[0];
				
		} catch(e) {
			return 0;
		}
		
	if (type==0) return [processObject.WorkingSetSize, processObject.PeakWorkingSetSize]; 
	return [processObject.ReadOperationCount,processObject.WriteOperationCount,processObject.OtherOperationCount, processObject.ReadTransferCount,  processObject.WriteTransferCount, processObject.OtherTransferCount];
	
}
/*
WshShell.Run("notepad");
MsgBox(ProcessGetStats());
MsgBox(ProcessGetStats("notepad.exe",0));
MsgBox(ProcessGetStats("notepad.exe", 1));
ProcessClose("notepad.exe");
//*/


//ProcessSetPriority Changes the priority of a process
function ProcessSetPriority(ProcName, priority )
{
	if (ProcName==null) return 0;
	if (priority==null) return 0;
	
		var processObject;
	if (!ProcName) ProcName = GetPID();
	if (typeof ProcName == "number")
		{
		try {
				var Proc = WMIQuery("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\CIMV2", "SELECT * FROM Win32_Process WHERE  ProcessId='"+ProcName+"'");
				if (Proc.length>0) processObject=Proc[0];
				
		} catch(e) {
			return 0;
		}
		}
		
		try {
				var Proc = WMIQuery("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\CIMV2", "SELECT * FROM Win32_Process WHERE Name='"+ProcName+"'");
				if (Proc.length>0) processObject=Proc[0];
				
		} catch(e) {
			return 0;
		}
/*
0 - Idle/Low
1 - Below Normal
2 - Normal
3 - Above Normal
4 - High
5 - Realtime (Use with caution, may make the system unstable)
*/
	var d=[128, 64, 32, 32768, 256];
	try {
		processObject.SetPriority(d[priority%5]);
		return 1;
	} catch(e) {
		return 0;
	}
}
/*
WshShell.Run("notepad.exe");
Sleep(2000)
MsgBox(ProcessSetPriority("notepad.exe", 0));
ProcessClose("notepad.exe");
//*/


//FileGetTime Returns the time and date information for a file.
function FileGetTime(filename, option, format)
{
	option = option || 0;
	format = format || 0;
	
	try {
		var f = FSO.GetFile(filename);
		var res = "";
		if (option==0) res = f.DateLastModified;
		if (option==1) res = f.DateCreated;
		if (option==2) res = f.DateLastAccessed;
		
		res = (new Date(res));
		res = [res.getFullYear(), res.getMonth()+1, res.getDate(), res.getHours(), res.getMinutes(), res.getSeconds()];
		if (format==0) return res;
		return res.map(function (i){if ((i+"").length<3) return ("00"+i).slice(-2); else return i;}).join("");	
	} catch(e) {
		return 0; 
	}
	//*/
}
//MsgBox(FileGetTime(GetScriptFullPath(), 1, 1));
//Exit();

//IsAdmin Checks if the current user has full administrator privileges.
function IsAdmin()
{
	var path = fso.GetSpecialFolder(1) + "\\IsAdmin.txt";
	try {
	f = fso.CreateTextFile(path,true);
	f.WriteLine("test");
	f.Close();
	fso.DeleteFile(path);
		return true;
	} catch(e) {
		return false;
	}
}
//MsgBox(IsAdmin());
//Exit();

//ClipGet Retrieves text from the clipboard.
function ClipGet()
{
	return new ActiveXObject("HTMLFile").parentWindow.clipboardData.getData("text");
}

//ClipPut Writes text to the clipboard.
function ClipPut(data)
{
	try {
		data = data.replace(/[\\'"]/g,"\\$&");
		new ActiveXObject("WScript.Shell").Run('mshta.exe "javascript:window.clipboardData.setData(\'text\',\''+data+'\');close()"',0);
		return true;
	} catch(e) {
		return false;
	}
}
/*
MsgBox(ClipGet());
MsgBox(ClipPut('"tutu4\\"'));
MsgBox(ClipGet());
Exit();
//*/

//FileSelectFolder Initiates a Browse For Folder dialog.
function FileSelectFolder (dialogText, rootDir, flag, initialDir, hwnd)
{
	//initialDir - ignore
	
	dialogText = dialogText || "";
	rootDir = rootDir || 0;
	hwnd = hwnd || 0;
	initialDir = initialDir || 0;
	
	var objShell = new ActiveXObject("shell.application");
	var objFolder;
	flag = flag || 1;
	flag = flag==1?0x200:flag;
	flag = flag==2?0x40:flag;
	flag = flag==4?0x10:flag;
	
	objFolder = objShell.BrowseForFolder(hwnd, dialogText, flag, rootDir);

	if (objFolder != null)
	{
		return objFolder.Self.Path;
	}
	
	return "";
}
/*
FileSelectFolder("test","",1);
FileSelectFolder("test","",2);
FileSelectFolder("test","",4);
//*/

//Run Runs an external program.
function Run(app, workingdir, show_flag)
{
/*
	SW_SHOWNORMAL	1	Файл будет запущен в обычном режиме.
	SW_MAXIMIZE	3	Файл будет запущен развернутым на весь экран.
	SW_MINIMIZE	6	Файл будет запущен свернутым.
	SW_HIDE	0	Файл будет запущен скрытым.
*/
/*WMI
1 - Window is shown minimized
3 - Window is shown maximized
5 - Window is shown in normal view
12 - Window is hidden and not displayed to the user
*/
	show_flag = show_flag || 0;
	if (workingdir==="" || workingdir===0) workingdir=null; 
	
	
	switch(show_flag)
	{
		case 0:
			show_flag = 12;
			break;
		case 3:
			break;
		case 6:
			show_flag = 1;
			break;
		case 1:
			show_flag = 5;
			break;
		default:
			show_flag = 12;
	}

	var objProcess = GetObject("winmgmts:\\\\.\\root\\cimv2:Win32_Process");
	var objInParams = objProcess.Methods_("Create").InParameters.SpawnInstance_();
	objInParams.CommandLine = app;
	objInParams.CurrentDirectory = workingdir;
	var objService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\CIMV2")
	var objStartup = objService.Get("Win32_ProcessStartup")
	var objConfig = objStartup.SpawnInstance_();
	objConfig.ShowWindow = show_flag; 
	objInParams.ProcessStartupInformation = objConfig;
	var objOutParams = objProcess.ExecMethod_("Create", objInParams);
	return objOutParams.ProcessId || 0;
}
/*
var pid = Run('notepad',"", 1);
MsgBox(pid);
ProcessClose(pid);
Exit();
//*/

//RunWait Runs an external program and pauses script execution until the program finishes.
function RunWait(app, workingdir, show_flag, timeout)
{
	var process = Run(app, workingdir, show_flag);
	ProcessWait(process, timeout);
	return process;
}
/*
var pid = RunWait('notepad',"", 0, 5);
MsgBox(pid);
Exit();
//*/

//FileGetShortName Returns the 8.3 short path+name of the path+name passed.
function FileGetShortName(path)
{
	try {
		if (FSO.FolderExists(path)) return FSO.GetFolder(path).ShortPath;
		if (FSO.FileExists(path)) return FSO.GetFile(path).ShortPath;
		return "";
	} catch(e) {return "";}
}

//MsgBox(FileGetShortName("C:\\Documents and Settings\\Сохранялка почты"));
//Exit();

//FileGetLongName Returns the long path+name of the path+name passed.
function FileGetLongName(p)
{
	var link = WshShell.CreateShortcut("%TMP%\\dummy.lnk");
	link.TargetPath = p;
	var Result = link.TargetPath;
	FileDelete("%TMP%\\dummy.lnk");
	return Result;
}
//MsgBox(FileGetLongName(GetHomeDrive() + "\\PROGRA~1"));
//Exit();

//FileGetSize Returns the size of a file in bytes.
function FileGetSize(path)
{
	try {
		return FSO.GetFile(path).Size;
	} catch(e) {
		return 0;
	}
}
//MsgBox(FileGetSize(GetScriptFullPath()));
//Exit();

//DirGetSize Returns the size in bytes of a given directory.
function DirGetSize(path)
{
	try {
		return FSO.GetFolder(path).Size;
	} catch(e) {
		return 0;
	}
}
//MsgBox(DirGetSize(GetWindowsDir()));
//Exit();

//EnvGet Retrieves an environment variable.
function EnvGet(str)
{
	return WshShell.ExpandEnvironmentStrings(str);
}

//FileCreateShortcut Creates a shortcut (.lnk) to a file.
function FileCreateShortcut (file, lnk, workdir, args, desc, icon, hotkey /*Example: CTRL + ALT + S*/, icon_number, state)
{
	if (!file || !lnk) return 0;
	if (icon_number===null) icon_number = ""; else icon_number =", "+icon_number;
	
	file = WshShell.ExpandEnvironmentStrings(file);
	lnk = WshShell.ExpandEnvironmentStrings(lnk);
	
	try {
		var WshShortcut = WshShell.CreateShortcut(lnk);
		WshShortcut.Arguments = args ||"";
		WshShortcut.Description = desc || "";
		WshShortcut.HotKey = hotkey || "";
		WshShortcut.IconLocation = icon!=""?(icon+icon_number): "";
		WshShortcut.TargetPath = file; 
		WshShortcut.WindowStyle = state;
		WshShortcut.WorkingDirectory = workdir || "";
		WshShortcut.Save();
		return 1;
	} catch(e) {return 0;}
}

//MsgBox(FileCreateShortcut(GetWindowsDir() + "\\explorer.exe", GetDesktopDir() + "\\Shortcut Example.lnk", GetWindowsDir(), "/e,c:\\", "Tooltip description of the shortcut.", GetSystemDir() + "\\shell32.dll", "CTRL+ALT+t", "15", 7));
//Exit();


//FileGetShortcut Retrieves details about a shortcut.
//return (file, lnk, workdir, args, desc, icon, hotkey, icon_number, state)
function FileGetShortcut(lnk)
{
	var result = [];
	if (!FileExists(lnk)) return 0;
	try {
		var WshShortcut = WshShell.CreateShortcut(lnk);
		result[0]=WshShortcut.TargetPath;
		result[1]=lnk;
		result[2]=WshShortcut.WorkingDirectory;
		result[3]=WshShortcut.Arguments;
		result[4]=WshShortcut.Description;
		result[5]=(WshShortcut.IconLocation+"").replace(/,\s*\d+\s*$/,"");
		result[6]=WshShortcut.HotKey;
		var tmp = (WshShortcut.IconLocation+"").match(/,\s*\d+\s*$/);
		if (tmp) tmp = tmp.toString().replace(/[^\d]/g,""); else tmp = "";
		result[7]=tmp;
		result[8]=WshShortcut.WindowStyle;
		
		return result;
	} catch(e) {return 0;}
}
/*
FileCreateShortcut(GetWindowsDir() + "\\explorer.exe", GetDesktopDir() + "\\cmd.exe.lnk", GetWindowsDir(), "/e,c:\\", "Tooltip description of the shortcut.", GetSystemDir() + "\\shell32.dll", "CTRL+ALT+t", "15", 7);
var res = FileGetShortcut(GetDesktopDir() + "\\cmd.exe.lnk");
MsgBox(res);
res[0]=GetWindowsDir() + "\\System32\\calc.exe";
FileCreateShortcut.apply(null, res);
res = FileGetShortcut(GetDesktopDir() + "\\cmd.exe.lnk");
MsgBox(res);
Exit();
//For test press CTRL+ALT+T
//*/

//FileSearch 
function FileSearch(dir, pattern, callback)
{	
	if (dir==0 || dir=="" || dir==null) dir = WshShell.CurrentDirectory;
	if (pattern==0 || pattern==null) pattern=/./;
	if (typeof pattern=="string") pattern=new RegExp(pattern.replace(/[\.\^\$\%\{\}\[\]\\\+]/g,"\\$&").replace(/[\?\*]/g,".$&")+"$","i");
	if (!callback) return 0;

	var func = arguments.callee;
	var level = 0;
	var exit_flag = false;
	
	if (!func.ids) func.ids = {};
	var set_search_id = GenerateString();
	func.ids[set_search_id]=1; e 
	
	var main_folder_path = WshShell.ExpandEnvironmentStrings(dir);
	try {
	var main_folder = fso.GetFolder(main_folder_path);
	} catch (e) {delete func.ids[set_search_id]; return 0};

	function DirWithSubFolders(_folder){
		if (exit_flag) return;
		level++;
		EnumerateFiles(_folder);
	    var more_folders = new Enumerator(_folder.SubFolders);
	    for (;!more_folders.atEnd();more_folders.moveNext())
		{
			OneFolder = more_folders.item();
			try {
			DirWithSubFolders (OneFolder);
			} catch (e) {};
		}
		level--;
	}

	function EnumerateFiles(_folder){
	    if (exit_flag) return;
		var more_files = new Enumerator(_folder.Files);
	    for (;!more_files.atEnd();more_files.moveNext())
		{
			one_file = more_files.item();
			if (pattern.test(one_file.Path)) callback(one_file.Path, level, set_search_id);
			if (!func.ids[set_search_id]){ exit_flag=true; return 1;}
	    }
	}
	
	DirWithSubFolders(main_folder);
	delete arguments.callee.ids[set_search_id];
	
	return 1;
}

function FileSearchStop(search_id)
{
	delete FileSearch.ids[search_id];
}

/*
var i=0, count=1, ar=[];
FileSearch("%USERPROFILE%", "*.hta", function(f, l, id){i++; if (i<=count) ar.push(f); FileSearchStop(id)});
MsgBox(ar);
Exit();
//*/

function FileSearchAll(dir, pattern, level)
{
	var ar=[];
	level = level || 0;
	FileSearch(dir, pattern, function(f, l, search_id){if (level!=0 && l<=level || level==0) ar.push(f); else FileSearchStop(search_id);});
	return ar;
}
//*
//MsgBox(FileSearchAll("", "*.ht*").join("\r\n"));
//MsgBox(FileSearchAll("%USERPROFILE%", "*.hta").join("\r\n"));
//Exit();
//*/

function FileFindFirst(v, count)
{
	var i=0, count = count || 1, ar=[];
	FileSearch("", v, function(f, l, id){i++; if (i<=count) ar.push(f); FileSearchStop(id)});
	return ar[0]?ar[0]:"";
}

//FileSetTime Sets the timestamp of one of more files.
function FileSetTime(root, mask, datestring, typeOfFunction, deep)
{
/*[FileSetTime[
import System.Environment;
import System.IO;

// searches files by a string type classic mask or regexp type mask. if in-param "callback" is undefined function returns an array of found pathes. 
function Search(root, re, callback, deep)
{
    var e, files=[], dirs=[], result=[];
    
    if (!root) root = Directory.GetCurrentDirectory();
    
    if (!re) re = /./;
    
    if (typeof re == "string" && /^\/.*\/[gmi ]*$/.test(re)) 
    {
        var m = re.match(/(^\/.*\/)([gmi ]*)$/i);
        re = new RegExp(m[1].replace(/^\/|\/$/g,""), m[2]);
    }
    
    if (typeof re == "string") re = new RegExp("("+re.replace(/[\^\\\*\[\]\{\}\(\)\$]/g, "\\$&").replace(/\|/g,")|(").replace(/\?/g, ".").replace(/(\*)/g, ".*")+"$)", "gi");
    
    if (!callback) callback = function (a){result.push(a)};
    
    if (typeof deep=="undefined") deep = true;
    deep = !!deep;
      
    try
    {
        files = Directory.GetFiles(root);
        
        for (var i=0, l=files.length; i<l; i++)
        {
            if (re.test(files[i])) callback(files[i]);
        }
        
        if (!deep) return result;
        
        dirs = Directory.GetDirectories(root);
        for (var i=0, l=dirs.length; i<l; i++)
        {
            try {
            Search(dirs[i], re, callback);
            } catch (e){ };
        }
    }
    catch (e) {}
    return result;
}

var root, mask, deep, datestring;

var params = System.Environment.GetCommandLineArgs();
// root, mask, dateString, typeOfFunction {0,1,2}, deep{true|false}

if (params.length!=6) System.Environment.Exit();

var dateMatch = (params[3]+"00000000000000").match(/^(....)(..)(..)(..)(..)(..)/);
var date = new Date(dateMatch[1], dateMatch[2]-1, dateMatch[3], dateMatch[4], dateMatch[5], dateMatch[6]);

function SetCreationTime(path)
{
    var e; try {File.SetCreationTime(path, date)} catch (e) {};
}

function SetLastAccessTime(path)
{
    var e; try {File.SetLastAccessTime(path, date)} catch (e) {};
}

function SetLastWriteTime(path)
{
    var e; try {File.SetLastWriteTime(path, date)} catch (e) {};
}

var funcs = [SetLastWriteTime, SetCreationTime, SetLastAccessTime];

Search(params[1], params[2], funcs[params[4] % 3], params[5]);
]]*/
    
    function CreateApp()
    {
        var path = WshShell.ExpandEnvironmentStrings("%TMP%\\setfiletime.txt");
        try {
        var f = fso.CreateTextFile(path,true);
        f.WriteLine(FileSetTime.GetResource("FileSetTime"));
        f.Close();
        var CurrentDirectory = WshShell.CurrentDirectory;
        WshShell.CurrentDirectory = WshShell.ExpandEnvironmentStrings("%TMP%");
        WshShell.Run("%WINDIR%\\Microsoft.NET\\Framework\\v2.0.50727\\jsc %TMP%\\setfiletime.txt", 0, 1);
        WshShell.CurrentDirectory = CurrentDirectory;
        
        }
        catch (e) {}
    }

    if (!root) root = WshShell.CurrentDirectory;
    if (!mask) mask = "";
    if (!datestring) datestring="1970010100000000"; 
    if (typeof typeOfFunction=="undefined") typeOfFunction = 0;
    if (typeof deep == "undefined") deep = 1;
    
    
    if (!FSO.FileExists(WshShell.ExpandEnvironmentStrings("%TMP%\\setfiletime.exe"))) {CreateApp();}
    if (FSO.FileExists(WshShell.ExpandEnvironmentStrings("%TMP%\\setfiletime.exe")))
    
	try {
		WshShell.Run("%TMP%\\setfiletime.exe "+'"'+[root, mask, datestring, typeOfFunction, deep].join('" "')+'"', 0, 1);
		return 1;} catch(e) {return 0;}
}
/*
MsgBox(FileGetTime(GetScriptFullPath()));
FileSetTime("", GetScriptFullPath(), "199212010000");
MsgBox(FileGetTime(GetScriptFullPath()));
Exit();
//*/

//FileGetAttrib Returns a code string representing a file's attributes.
function FileGetAttrib(file)
/*
String returned could contain a combination of these letters "RASHNDOCT":
"R" = READONLY = 1
"A" = ARCHIVE = 32
"S" = SYSTEM = 4
"H" = HIDDEN = 2
"N" = NORMAL = 0
"D" = DIRECTORY = 16
"C" = COMPRESSED (NTFS compression, not ZIP compression)=128
*/
{
	try {
	
	var attrs = FSO.GetFile(file).Attributes;
	var res = "";
	if (attrs & 1) res+="R"; 
	if (attrs & 32) res+="A";
	if (attrs & 4) res+="S";
	if (attrs & 2) res+="H";
	if (attrs & 16) res+="D";
	
	if (res=="") res="N";
	
	return res;
	
	} catch(e) {
		return "";
	}
}
//MsgBox(FileGetAttrib(GetScriptFullPath()));
//Exit();

//*
//FileSetAttrib Sets the attributes of one or more files.
function FileSetAttrib (pattern, attr, recurse)
{
//"+-RASHN"
	
	recurse = !!recurse;
	attr = attr.replace(/[^\-\+RASHN]+/gi, "").toUpperCase();
	
	var unsetFlag = attr.match(/(?:-)\w+/g);
	if (!unsetFlag) unsetFlag=""; else unsetFlag = unsetFlag.join("").replace(/\-/g,"").replace(/N/g,"");
	
	var setFlag = attr.match(/^\w+|\+\w+/g);
	if (!setFlag) setFlag=""; else setFlag = setFlag.join("").replace(/\+/g,"");
	
	var mask = {R:1, H:2, S:4, A:32};
	
	setFlag = setFlag.split("");
	unsetFlag = unsetFlag.split("");

	function callback(a)
	{
		var file = FSO.GetFile(a)
		var attrs = file.Attributes;

		function callback_set(s)
		{
			if (s=="N") return attrs = 0;
			attrs = (attrs | mask[s]);
		}
		
		function callback_unset(s)
		{
			if (attrs & mask[s]) attrs = attrs ^ mask[s];
		}
		
		unsetFlag.forEach(callback_unset);
		setFlag.forEach(callback_set);
		
		file.Attributes = attrs;
	}

	FileSearch("", pattern, function(f, l, id){if (recurse && l!=1) FileSearchStop(id); callback(f);});

	return;
}
/*
MsgBox(FileGetAttrib(GetScriptFullPath()));
FileSetAttrib(GetScriptFullPath(),"-A");
MsgBox(FileGetAttrib(GetScriptFullPath()));
Exit();
//*/

//ShellExecute Runs an external program using the ShellExecute API.
//ShellExecute ( "filename" [, "parameters" [, "workingdir" [, "verb" [, showflag]]]] )
function ShellExecute(filename, parameters, workingdir, verb , showflag)
{
	var objShell = new ActiveXObject("shell.application");	
	objShell.ShellExecute(filename, parameters, workingdir, verb, showflag);
}
//MsgBox(ShellExecute("notepad.exe", "", "", "open", 1));

//ShellExecuteWait Runs an external program using the ShellExecute API and pauses script execution until it finishes.
function ShellExecuteWait(filename, parameters, workingdir, verb, showflag)
{
	try {

		var time1 = new Number(GetDATE());
		ShellExecute(filename, parameters, workingdir, verb, showflag);

		var time2 ;
		var Proc = WMIQuery("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\CIMV2", "SELECT * FROM Win32_Process WHERE Name='"+filename.match(/[^\\\/]+$/)+"'");
		
		Sleep(500);
		for (var i=0, l=Proc.length; i<l; i++)
		{
			time2 = new Number(Proc[i].CreationDate.split(".")[0]);
			if ((time2-time1)<3) {ProcessWait(Proc[i].ProcessId); return 1;}
		}
		
		return 1;
	} catch(e) {return 0;}
}
/*
MsgBox(ShellExecuteWait("notepad.exe", "", "", "open", 1));
Exit();
//*/

//MemGetStats Retrieves memory related information.
function MemGetStats()
{
/*
returns Array:
$array[0] = Memory Load (Percentage of memory in use)
$array[1] = Total physical RAM
$array[2] = Available physical RAM
$array[3] = Total Pagefile
$array[4] = Available Pagefile
$array[5] = Total virtual
$array[6] = Available virtual

All memory sizes are in kilobytes.
*/
	
	var Result = [], Proc;
	
	Proc = WMIQuery("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\CIMV2", "SELECT * FROM Win32_OperatingSystem")[0];
	Result[1] = Proc.TotalVisibleMemorySize;
	Result[2] = Proc.FreePhysicalMemory;
	Result[3] = Proc.TotalVirtualMemorySize;
	Result[4] = Proc.FreeSpaceInPagingFiles;
	Result[5] = Proc.TotalVirtualMemorySize;
	Result[6] = Proc.FreeVirtualMemory;

	Result[0] = Math.round(Result[2] / Result[1] * 100);
	
	return Result;
}
/*
MsgBox(MemGetStats());
Exit();
//*/

//RegDelete Deletes a key or value from the registry.
function RegDelete(path)
{
	return WshShell.RegDelete(path)
}
