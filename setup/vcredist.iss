#ifndef Platform
	#error "Platform" predefined variable must be defined (x64/Win32)
#endif

#define VCDIR ExtractFilePath(__PATHFILENAME__)

[Files]
#if Platform == "x64"
Source: "{#VCDIR}\vc_redist.x64.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall 64bit; Check: IsWin64; AfterInstall: ExecTemp( 'vc_redist.x64.exe', '/passive /norestart' );
#else
Source: "{#VCDIR}\vc_redist.x86.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall 32bit; AfterInstall: ExecTemp( 'vc_redist.x86.exe', '/passive /norestart' );
#endif

#undef VCDIR

[Code]
var
  NeedsRestart: Boolean;

procedure ExecTemp(File, Params : String);
var
	ResultCode: Integer;
begin
	Exec( ExpandConstant( '{tmp}' ) + '\' + File, Params, ExpandConstant( '{tmp}' ), SW_SHOW, ewWaitUntilTerminated, ResultCode );
  if ResultCode = 3010 then begin
    NeedsRestart := True;
  end
  else if ResultCode = 1638 then begin
    // Alread installed
  end
  else if ResultCode <> 0 then begin
    MsgBox( FmtMessage( SetupMessage( msgErrorFunctionFailed ), [ 'Microsoft Visual C++ Redistributable', IntToStr( ResultCode ) ] ), mbError, MB_OK );
  end;
	// Clean log files
  DelTree( GetTempDir() + '\dd_vcredist*', false, true, false );
end;

function NeedRestart(): Boolean;
begin
  Result := NeedsRestart;
end;
