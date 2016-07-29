#ifdef UNICODE
  #define AW "W"
#else
  #define AW "A"
#endif

#include "lang\default.iss"

[Code]
type
    INSTALLSTATE = Longint;
const
    INSTALLSTATE_INVALIDARG = -2;  // An invalid parameter was passed to the function.
    INSTALLSTATE_UNKNOWN = -1;     // The product is neither advertised or installed.
    INSTALLSTATE_ADVERTISED = 1;   // The product is advertised but not installed.
    INSTALLSTATE_ABSENT = 2;       // The product is installed for a different user.
    INSTALLSTATE_DEFAULT = 5;      // The product is installed for the current user.

function MsiQueryProductState(szProduct: string): INSTALLSTATE;
external 'MsiQueryProductState{#AW}@msi.dll stdcall';

function MsiProduct(const ProductID: string): boolean;
begin
  Result := MsiQueryProductState(ProductID) = INSTALLSTATE_DEFAULT;
end;

type
	TProduct = record
		File: String;
		Title: String;
		Parameters: String;
		InstallClean : boolean;
		MustRebootAfter : boolean;
		AnyVersion : boolean;
	end;

	InstallResult = (InstallSuccessful, InstallRebootRequired, InstallError);

var
	installMemo, downloadMemo, downloadMessage: string;
	products: array of TProduct;
	delayedReboot: boolean;
	DependencyPage: TOutputProgressWizardPage;

procedure AddProduct(FileName, Parameters, Title, URL: string; InstallClean, MustRebootAfter, AnyVersion : boolean);
var
	path: string;
	i: Integer;
begin
	installMemo := installMemo + '%1' + Title + #13;

  path := ExpandConstant('{tmp}{\}') + FileName;

  idpAddFile( URL, path );

  downloadMemo := downloadMemo + '%1' + Title + #13;
  downloadMessage := downloadMessage + '	' + Title + #13;

	i := GetArrayLength(products);
	SetArrayLength(products, i + 1);
	products[i].File := path;
	products[i].Title := Title;
	products[i].Parameters := Parameters;
	products[i].InstallClean := InstallClean;
	products[i].MustRebootAfter := MustRebootAfter;
	products[i].AnyVersion := AnyVersion;
end;

function SmartExec(prod : TProduct; var ResultCode : Integer) : Boolean;
begin
	if (LowerCase(Copy(prod.File,Length(prod.File)-2,3)) = 'exe') then begin
		Result := Exec(prod.File, prod.Parameters, '', SW_SHOWNORMAL, ewWaitUntilTerminated, ResultCode);
	end else begin
		Result := ShellExec('', prod.File, prod.Parameters, '', SW_SHOWNORMAL, ewWaitUntilTerminated, ResultCode);
	end;
end;       

function PendingReboot : Boolean;
var	names: String;
begin
	if (RegQueryMultiStringValue(HKEY_LOCAL_MACHINE, 'SYSTEM\CurrentControlSet\Control\Session Manager', 'PendingFileRenameOperations', names)) then begin
		Result := true;
	end else if ((RegQueryMultiStringValue(HKEY_LOCAL_MACHINE, 'SYSTEM\CurrentControlSet\Control\Session Manager', 'SetupExecute', names)) and (names <> ''))  then begin
		Result := true;
	end else begin
		Result := false;
	end;
end;

const
  ERROR_SUCCESS	                  = 0;    // The action completed successfully.
  ERROR_PRODUCT_VERSION           = 1638; // Another version of this product is already installed.
  ERROR_SUCCESS_REBOOT_INITIATED  =	1641; // The installer has initiated a restart. This message is indicative of a success. 
  ERROR_SUCCESS_REBOOT_REQUIRED   = 3010; // A restart is required to complete the install. This message is indicative of a success.

function InstallProducts: InstallResult;
var
	ResultCode, i, productCount, finishCount: Integer;
begin
	Result := InstallSuccessful;
	productCount := GetArrayLength(products);

	if productCount > 0 then begin
		DependencyPage := CreateOutputProgressPage(ExpandConstant('{cm:depinstall_title}'), ExpandConstant('{cm:depinstall_description}'));
		DependencyPage.Show;

		for i := 0 to productCount - 1 do begin
			if (products[i].InstallClean and (delayedReboot or PendingReboot())) then begin
				Result := InstallRebootRequired;
				break;
			end;

			DependencyPage.SetText(FmtMessage(ExpandConstant('{cm:depinstall_status}'), [products[i].Title]), '');
			DependencyPage.SetProgress(i, productCount);

			if SmartExec(products[i], ResultCode) then begin
				//setup executed; ResultCode contains the exit code
				if (products[i].MustRebootAfter) then begin
					//delay reboot after install if we installed the last dependency anyways
					if (i = productCount - 1) then begin
						delayedReboot := true;
					end else begin
						Result := InstallRebootRequired;
					end;
					break;
				end else if ((ResultCode = ERROR_SUCCESS) or (products[i].AnyVersion and (ResultCode = ERROR_PRODUCT_VERSION))) then begin
					finishCount := finishCount + 1;
				end else if ((ResultCode = ERROR_SUCCESS_REBOOT_INITIATED) or (ResultCode = ERROR_SUCCESS_REBOOT_REQUIRED)) then begin
					delayedReboot := true;
					finishCount := finishCount + 1;
				end else begin
					Result := InstallError;
					break;
				end;
			end else begin
				Result := InstallError;
				break;
			end;
		end;

		//only leave not installed products for error message
		for i := 0 to productCount - finishCount - 1 do begin
			products[i] := products[i+finishCount];
		end;
		SetArrayLength(products, productCount - finishCount);

		DependencyPage.Hide;
	end;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var
	i: Integer;
	s: string;
begin
	delayedReboot := false;

	case InstallProducts() of
		InstallError: begin
			s := ExpandConstant('{cm:depinstall_error}');
			for i := 0 to GetArrayLength(products) - 1 do begin
				s := s + #13 + '	' + products[i].Title;
			end;
			Result := s;
			end;

		InstallRebootRequired: begin
			Result := products[0].Title;
			NeedsRestart := true;
			//write into the registry that the installer needs to be executed again after restart
			RegWriteStringValue(HKEY_CURRENT_USER, 'SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce', 'InstallBootstrap', ExpandConstant('{srcexe}'));
			end;
	end;
end;

function NeedRestart : Boolean;
begin
	if (delayedReboot) then
		Result := true;
end;

function UpdateReadyMemo(Space, NewLine, MemoUserInfoInfo, MemoDirInfo, MemoTypeInfo, MemoComponentsInfo, MemoGroupInfo, MemoTasksInfo: String): String;
var
	s: string;
begin
	if downloadMemo <> '' then
		s := s + ExpandConstant('{cm:depdownload_memo_title}') + ':' + NewLine + FmtMessage(downloadMemo, [Space]) + NewLine;
	if installMemo <> '' then
		s := s + ExpandConstant('{cm:depinstall_memo_title}') + ':' + NewLine + FmtMessage(installMemo, [Space]) + NewLine;
	s := s + MemoDirInfo + NewLine + NewLine + MemoGroupInfo
	if MemoTasksInfo <> '' then
		s := s + NewLine + NewLine + MemoTasksInfo;
	Result := s
end;
