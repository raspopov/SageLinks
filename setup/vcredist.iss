#ifndef Platform
	#error "Platform" predefined variable must be defined (x64/Win32)
#endif

#if Platform == "x64"
	#define vcredist_exe	"VC_redist.x64.exe"
	#define vcredist_url	"https://aka.ms/vs/16/release/VC_redist.x64.exe"
#else
	#define vcredist_exe	"VC_redist.x86.exe"
	#define vcredist_url	"https://aka.ms/vs/16/release/VC_redist.x86.exe"
#endif

#define vcdir ExtractFilePath(__PATHFILENAME__)
#expr Exec( 'powershell', '-NoProfile -Command (New-Object System.Net.WebClient).DownloadFile(\"' + vcredist_url + '\",\"' + vcdir + '\' + vcredist_exe + '\")' )

[Files]
#if Platform == "x64"
Source: "{#vcdir}\{#vcredist_exe}"; DestDir: "{tmp}"; Flags: deleteafterinstall 64bit; Check: IsWin64; AfterInstall: ExecTemp( '{#vcredist_exe}', '/passive /norestart' );
#else
Source: "{#vcdir}\{#vcredist_exe}"; DestDir: "{tmp}"; Flags: deleteafterinstall 32bit; AfterInstall: ExecTemp( '{#vcredist_exe}', '/passive /norestart' );
#endif

[Code]
procedure ExecTemp(File, Params : String);
var
	nCode: Integer;
begin
	Exec( ExpandConstant( '{tmp}' ) + '\' + File, Params, ExpandConstant( '{tmp}' ), SW_SHOW, ewWaitUntilTerminated, nCode );
end;
