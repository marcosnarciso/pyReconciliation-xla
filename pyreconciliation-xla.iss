[Files]
Source: Pyinex.xll; DestDir: {app}; Flags: confirmoverwrite promptifolder;
;Source: pyRecon.xlam; DestDir: {userappdata}\Microsoft\Suplementos; Flags: confirmoverwrite promptifolder
Source: pyRecon.xlam; DestDir: {app}; Flags: confirmoverwrite promptifolder;
Source: main.pyc; DestDir: {app}; Flags: confirmoverwrite promptifolder;
Source: readme.md; DestDir: {app}; Flags: confirmoverwrite promptifolder; DestName: ReadMe.md
Source: pyRecon\__init__.pyc; DestDir: {app}\pyRecon; Flags: confirmoverwrite promptifolder;
Source: pyRecon\pyRecon.pyc; DestDir: {app}\pyRecon; Flags: confirmoverwrite promptifolder;

[Setup]
AppName=Data Reconciliation Supplement for MS Excel
AppVerName=Data Reconciliation Supplement for MS Excel 0.91a
EnableDirDoesntExistWarning=true
AppID={{55526959-AC1B-4F73-9D57-7CC230726F0D}
MinVersion=5.0.2195
LicenseFile=license.txt
DefaultDirName={pf}\Microsoft Office\{code:OfficeVersion}\pyRecon
AlwaysShowDirOnReadyPage=false
AppVersion=0.91a0
UninstallDisplayName=Data Reconciliation Supplement for MS Excel
DefaultGroupName=Data Reconciliation Supplement for MS Excel
AllowUNCPath=true
AlwaysShowGroupOnReadyPage=false
ShowTasksTreeLines=true
VersionInfoVersion=0.91
VersionInfoCompany=UFBA/TECLIM
VersionInfoProductName=Data Reconciliation Supplement for MS Excel
VersionInfoProductVersion=0.91
AppPublisherURL=https://bitbucket.org/marcosnarciso/pyreconciliation-xla
OutputBaseFilename=setup-0.91a
Compression=lzma
SolidCompression=Yes

[Languages]
Name: en; MessagesFile: "compiler:Default.isl"
Name: ptbr; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Messages]
en.BeveledLabel=English
ptbr.BeveledLabel=Português (Brasil)

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Office\{code:OfficeNumericVersion}\Excel\Add-in Manager"; ValueType: string; ValueName: {app}\pyRecon.xlam; ValueData: ""; Flags: uninsdeletevalue
;Root: HKCU; Subkey: "Software\Microsoft\Office\{code:OfficeNumericVersion}\Excel\Options"; ValueType: string; ValueName: "OPEN1"; ValueData: "{app}/EMSO_XLA.xlam"; Flags: uninsdeletevalue

[Icons]
Name: {group}\Project WebSite; Filename: {app}\website.url; WorkingDir: {app}
Name: {group}\Uninstall Data Reconciliation Supplement for Excel; Filename: {uninstallexe}
Name: {group}\ReadMe; Filename: {app}\ReadMe.md; WorkingDir: {app}
;Name: {userdesktop}\EMSO Reconciliation; Check: Office2003; Filename: {userappdata}\Microsoft\Suplementos\EMSO_XLA.xlam; WorkingDir: {userdesktop}
Name: {commondesktop}\Data Reconciliation; Filename: {app}\pyRecon.xlam; WorkingDir: {commondesktop}

[UninstallDelete]
Type: files; Name: {app}\website.url
Name: {app}; Type: dirifempty

[INI]
Filename: {app}\website.url; Section: InternetShortcut; Key: URL; String: https://bitbucket.org/marcosnarciso/pyreconciliation-xla; Flags: createkeyifdoesntexist

[CustomMessages]
en.officemissing=This setup requires the Microsoft Office. Do you want to download the shareware version now?
ptbr.officemissing=Este programa necessita da Suíte de Aplicativos Microsoft Office instalada. Você quer efetuar o download de uma versão de testes?

[CODE]
function InitializeSetup(): Boolean;
var
    ErrorCode: Integer;
    OfficeInstalled : Boolean;
    Result1 : Boolean;
begin

  OfficeInstalled := RegKeyExists(HKLM,'SOFTWARE\Microsoft\Office\14.0\Excel');
	if OfficeInstalled =true then begin
		Result := true;
	end;

	if OfficeInstalled = false then	begin
		OfficeInstalled := RegKeyExists(HKLM,'SOFTWARE\Microsoft\Office\12.0\Excel');
		if OfficeInstalled =true then	begin
			Result := true;
		end;

    if OfficeInstalled = false then begin
  		OfficeInstalled := RegKeyExists(HKLM,'SOFTWARE\Microsoft\Office\11.0\Excel');
      if OfficeInstalled =true then begin
        Result := true;
      end;

      if OfficeInstalled =false then begin
        Result1 := MsgBox(ExpandConstant('{cm:officemissing}'), mbConfirmation, MB_YESNO) = idYes;
  			if Result1 =false then begin
  				Result:=false;
	   		end
		  	else begin
		  		Result:=false;
    			ShellExec('open','http://care.dlservice.microsoft.com/dl/release/6/E/C/6EC80137-B3C0-4D2C-97E5-82358484D3E5/14.0.4763.1000_ProfessionalPlus_retail_ship_x86_pt-br_exe/ProfessionalPlus.exe?lcid=1046','','',SW_SHOWNORMAL, ewNoWait, ErrorCode);
        end;
      end;
    end;
  end;
end;

function Office2003(): Boolean;
var
sExcelVar: String;
begin
  RegQueryStringValue(HKEY_CLASSES_ROOT, 'Excel.Application\CurVer',  '', sExcelVar);
  if sExcelVar ='Excel.Application.11' then begin
    Result := true;
  end
  else begin
    Result := false;
  end;
end;

function Office2007(): Boolean;
var
sExcelVar: String;
begin
  RegQueryStringValue(HKEY_CLASSES_ROOT, 'Excel.Application\CurVer',  '', sExcelVar);
  if sExcelVar ='Excel.Application.12' then begin
    Result := true;
  end
  else begin
    Result := false;
  end;
end;

function Office2010(): Boolean;
var
sExcelVar: String;
begin
  RegQueryStringValue(HKEY_CLASSES_ROOT, 'Excel.Application\CurVer',  '', sExcelVar);
  if sExcelVar ='Excel.Application.14' then begin
    Result := true;
  end
  else begin
    Result := false;
  end;
end;

function OfficeVersion(Param:String): String;
var
sExcelVar: String;
begin
  RegQueryStringValue(HKEY_CLASSES_ROOT, 'Excel.Application\CurVer',  '', sExcelVar);
  Case sExcelVar of
    'Excel.Application.11': begin
    Result := 'Office11\Bibliote';
    end;
    'Excel.Application.12': begin
    Result := 'Office12\Library';
    end;
    'Excel.Application.14': begin
    Result := 'Office14\Library';
    end;
  end;
end;

function OfficeNumericVersion(Param:String): String;
var
sExcelVar: String;
begin
  RegQueryStringValue(HKEY_CLASSES_ROOT, 'Excel.Application\CurVer',  '', sExcelVar);
  Case sExcelVar of
    'Excel.Application.11': begin
    Result := '11.0';
    end;
    'Excel.Application.12': begin
    Result := '12.0';
    end;
    'Excel.Application.14': begin
    Result := '14.0';
    end;
  end;
end;
