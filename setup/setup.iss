#if Platform == "x64"
  #define MyBitness		  "64-bit"
#else
  #define MyBitness		  "32-bit"
#endif

#define MyAppExe		    "SageLinks.exe"
#define MyAppSource		  "..\" + Platform + "\Release\" + MyAppExe
#define MyAppName			  GetStringFileInfo( MyAppSource, INTERNAL_NAME )
#define MyAppVersion	  GetFileProductVersion( MyAppSource )
#define MyAppCopyright	GetFileCopyright( MyAppSource )
#define MyAppPublisher	GetFileCompany( MyAppSource )
#define MyAppURL		    GetStringFileInfo( MyAppSource, "Comments" )
#define MyOutput		    LowerCase( StringChange( MyAppName + " " + MyAppVersion + " " + MyBitness, " ", "_" ) )
          
#include "vc_redist.iss"

#include "idp\lang\russian.iss"
#include "idp\idp.iss"

#include "dep\lang\russian.iss"
#include "dep\dep.iss"

[Setup]
AppId={#MyAppName}
AppName={#MyAppName} {#MyBitness}
AppVersion={#MyAppVersion}
VersionInfoVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
AppMutex=Global\{#MyAppName}
AppCopyright={#MyAppCopyright}
DefaultDirName={pf}\{#MyAppPublisher}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=..\{#Platform}\Release
OutputBaseFilename={#MyOutput}
Compression=lzma2/ultra64
SolidCompression=yes
InternalCompressLevel=ultra64
LZMAUseSeparateProcess=yes
PrivilegesRequired=admin
UninstallDisplayIcon={app}\{#MyAppExe},0
DirExistsWarning=no
WizardImageFile=compiler:WizModernImage-IS.bmp
WizardSmallImageFile=compiler:WizModernSmallImage-IS.bmp
SetupMutex=Global\Setup_{#MyAppName}
OutputManifestFile=Setup-Manifest.txt
#if Platform == "x64"
ArchitecturesInstallIn64BitMode=x64
ArchitecturesAllowed=x64
#endif

[Languages]
Name: "en"; MessagesFile: "compiler:Default.isl"
Name: "ru"; MessagesFile: "compiler:Languages\Russian.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "{#MyAppSource}"; DestDir: "{app}"; Flags: replacesameversion uninsrestartdelete
Source: "..\ReadMe.md"; DestName: "ReadMe.txt"; DestDir: "{app}"; Flags: replacesameversion uninsrestartdelete
Source: "..\LICENSE"; DestName: "License.txt"; DestDir: "{app}"; Flags: replacesameversion uninsrestartdelete

[Icons]
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExe}"; Tasks: desktopicon
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExe}"
Name: "{group}\ReadMe.txt"; Filename: "{app}\ReadMe.txt"
Name: "{group}\License.txt"; Filename: "{app}\License.txt"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}";	Filename: "{uninstallexe}"

[Run]
Filename: "{app}\{#MyAppExe}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall runasoriginaluser

[Registry]
Root: HKCU; Subkey: "Software\{#MyAppPublisher}\{#MyAppName}"; Flags: dontcreatekey uninsdeletekey
Root: HKCU; Subkey: "Software\{#MyAppPublisher}"; Flags: dontcreatekey uninsdeletekeyifempty

[UninstallDelete]
Name: "{userstartup}\{#MyAppName}.lnk"; Type: files
Name: "{commonstartup}\{#MyAppName}.lnk"; Type: files
Name: "{app}"; Type: dirifempty
Name: "{pf}\{#MyAppPublisher}"; Type: dirifempty
Name: "{localappdata}\{#MyAppPublisher}\{#MyAppName}"; Type: filesandordirs
Name: "{localappdata}\{#MyAppPublisher}"; Type: dirifempty

[Code]
procedure InitializeWizard();
var
  bWork: Boolean;
begin
  bWork := False;

  if Not (FileExists(ExpandConstant('{sys}\mfc140u.dll')) and FileExists(ExpandConstant('{sys}\vcruntime140.dll'))) then begin
  #if Platform == "Win32"
    if ( not MsiProduct( '{#vcredist_32_productcode}' ) ) then begin
      AddProduct( '{#vcredist_32_exe}', '/quiet /norestart', '{#vcredist_32_title}', '{#vcredist_32_url}', false, false, true );
      bWork := True;
    end;
  #endif
  #if Platform == "x64"
    if ( not MsiProduct( '{#vcredist_64_productcode}' ) ) then begin
      AddProduct( '{#vcredist_64_exe}', '/quiet /norestart', '{#vcredist_64_title}', '{#vcredist_64_url}', false, false, true );
      bWork := True;
    end;
  #endif
  end;

  if bWork then begin
    idpDownloadAfter( wpReady );
    idpSetDetailedMode( True );
  end;
end;

