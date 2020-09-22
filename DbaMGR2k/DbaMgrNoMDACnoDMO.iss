; Requires InnoSetup 2.0.18/19
; Richiede InnoSetup 2.0.18/19
; modified xx/11/2002

[Setup]
OutputDir=C:\Programmi\Inno Setup 2\OutPut
OutputBaseFilename =Setup_DbaMgr2k
AppName=DbaMGR2k
AppVerName=DbaMGR2k
AppVersion=0.1.0
AppId=DbaMGR2k
AppCopyright=Copyright (C) 2002 by Insuline Power
AppSupportURL=http://utenti.lycos.it/asql/index.html
AppUpdatesURL=http://utenti.lycos.it/asql/index.html
AppPublisher=Insuline Power
LicenseFile=D:\ADOvb6\DbaMGR2k\DbaMgr2k-EULA.TXT
WizardImageFile=C:\Programmi\Inno Setup 2\WizModernImage.bmp
WizardSmallImageFile=C:\Programmi\Inno Setup 2\WizModernSmallImage.bmp
DefaultDirName={pf}\DbaMGR2k
DefaultGroupName=DbaMGR2k
UninstallDisplayIcon={app}\DbaMGR2k.exe
MinVersion=4.0.950,4.0.1381sp3
AlwaysRestart=true
AlwaysCreateUninstallIcon=true
AlwaysShowComponentsList=false
AdminPrivilegesRequired=true
;MessagesFile=C:\Programmi\Inno Setup 2\Italian.isl
InfoAfterFile=D:\ADOvb6\DbaMGR2k\PostInst.txt
Compression=zip/9
InfoBeforeFile=D:\ADOvb6\DbaMGR2k\Eula2kNO-dmo-mdac.txt


[Tasks]
Name: desktopicon; Description: "Crea un'icona sul &Desktop "; GroupDescription: Additional icons:; MinVersion: 4,4
Name: quicklaunchicon; Description: Crea un'icona &Quick Launch; GroupDescription: Additional icons:; MinVersion: 4,4

[INI]
Filename: {app}\DbaMGR2k.url; Section: InternetShortcut; Key: URL; String: http://utenti.lycos.it/asql/index.html

[Icons]
Name: {group}\DbaMGR2k; Filename: {app}\DbaMGR2k.exe
Name: {commondesktop}\DbaMGR2k; Filename: {app}\DbaMGR2k.exe; MinVersion: 4,4; Tasks: desktopicon
Name: {userappdata}\Microsoft\Internet Explorer\Quick Launch\DbaMGR2k; Filename: {app}\DbaMGR2k.exe; MinVersion: 4,4; Tasks: quicklaunchicon
Name: {group}\Eula; Filename: {app}\DbaMGR2k-EULA.TXT; WorkingDir: {app}; IconIndex: 0; Flags: createonlyiffileexists


[Files]

; start vb CORE files
;Source: C:\WINDOWS\SYSTEM\VB6STKIT.DLL; DestDir: {sys};CopyMode: onlyifdoesntexist;Flags:sharedfile uninsneveruninstall
Source: C:\Programmi\Microsoft Visual Studio\VB98\Wizards\PDWizard\Redist98\COMCAT.DLL; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall regserver allowunsafefiles
Source: C:\Programmi\Microsoft Visual Studio\VB98\Wizards\PDWizard\Redist98\STDOLE2.TLB; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall regtypelib
Source: C:\Programmi\Microsoft Visual Studio\VB98\Wizards\PDWizard\Redist98\ASYCFILT.DLL; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall
Source: C:\Programmi\Microsoft Visual Studio\VB98\Wizards\PDWizard\Redist98\OLEPRO32.DLL; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall regserver
Source: C:\Programmi\Microsoft Visual Studio\VB98\Wizards\PDWizard\Redist98\OLEAUT32.DLL; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall regserver
Source: C:\Programmi\Microsoft Visual Studio\VB98\Wizards\PDWizard\Redist98\VB6IT.DLL; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: sharedfile uninsneveruninstall
Source: C:\Programmi\Microsoft Visual Studio\VB98\Wizards\PDWizard\Redist98\MSVBVM60.DLL; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall regserver
; end vb CORE files

; start vb files
Source: C:\Programmi\Microsoft Visual Studio\VB98\Wizards\PDWizard\Redist98\STDFTIT.DLL; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: sharedfile
Source: C:\Programmi\Microsoft Visual Studio\VB98\Wizards\PDWizard\Redist98\MSSTDFMT.DLL; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: regserver sharedfile
Source: C:\Programmi\Microsoft Visual Studio\VB98\Wizards\PDWizard\Redist98\Msbind.dll; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: regserver sharedfile
Source: C:\Programmi\Microsoft Visual Studio\VB98\Wizards\PDWizard\Redist98\ADODCIT.DLL; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: sharedfile
Source: C:\Programmi\Microsoft Visual Studio\VB98\Wizards\PDWizard\Redist98\MSADODC.OCX; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: regserver sharedfile
Source: C:\Programmi\Microsoft Visual Studio\VB98\Wizards\PDWizard\Redist98\DATGDIT.DLL; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: sharedfile
Source: C:\Programmi\Microsoft Visual Studio\VB98\Wizards\PDWizard\Redist98\MSDATGRD.OCX; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: regserver sharedfile
Source: C:\Programmi\Microsoft Visual Studio\VB98\Wizards\PDWizard\Redist98\FLXGDIT.DLL; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: sharedfile
Source: C:\Programmi\Microsoft Visual Studio\VB98\Wizards\PDWizard\Redist98\MSFLXGRD.OCX; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: regserver sharedfile
Source: C:\Programmi\Microsoft Visual Studio\VB98\Wizards\PDWizard\Redist98\RCHTXIT.DLL; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: sharedfile
Source: C:\Programmi\Microsoft Visual Studio\VB98\Wizards\PDWizard\Redist98\RICHED32.DLL; DestDir: {sys}; CopyMode: alwaysskipifsameorolder
Source: C:\Programmi\Microsoft Visual Studio\VB98\Wizards\PDWizard\Redist98\RICHTX32.OCX; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: regserver sharedfile
Source: C:\Programmi\Microsoft Visual Studio\VB98\Wizards\PDWizard\Redist98\MSCMCIT.DLL; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: sharedfile
Source: C:\Programmi\Microsoft Visual Studio\VB98\Wizards\PDWizard\Redist98\MSCOMCTL.OCX; DestDir: {sys}; CopyMode: alwaysskipifsameorolder; Flags: regserver sharedfile
; end vb files

Source: D:\ADOvb6\DbaMGR2k\Leggimi2k.txt; DestDir: {app}; CopyMode: onlyifdoesntexist
Source: D:\ADOvb6\DbaMGR2k\readme2k.txt; DestDir: {app}; CopyMode: onlyifdoesntexist
Source: D:\ADOvb6\DbaMGR2k\DbaMgr2k.exe; DestDir: {app}; CopyMode: onlyifdoesntexist
Source: D:\ADOvb6\DbaMGR2k\DbaMgr2k-EULA.TXT; DestDir: {app}; CopyMode: alwaysoverwrite
Source: D:\ADOvb6\DbaMGR2k\Depend2k.txt; DestDir: {app}; CopyMode: alwaysoverwrite

[UninstallDelete]
; attenzione, cancella tutti i files...
Type: files; Name: {app}\*.ln2

[Run]
; lettura Note
Filename: {app}\PostInst.txt; Flags: postinstall shellexec skipifsilent skipifdoesntexist

