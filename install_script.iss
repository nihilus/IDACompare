;InnoSetupVersion=4.2.6

[Setup]
AppName=IDACompare
AppVerName=IDACompare v0.3
DefaultDirName=c:\iDefense\IDACompare\
DefaultGroupName=IDACompare
OutputBaseFilename=IDACompare_Setup
OutputDir=./

[Files]
Source: IDA_Compare.plw; DestDir: {app}
Source: IDA_Compare.p64; DestDir: {app}
Source: IDASrvr.plw; DestDir: {app}
Source: blank.mdb; DestDir: {app}
Source: compare.vbs; DestDir: {app}
Source: ida_compare.exe; DestDir: {app}; Flags: ignoreversion
Source: IdaCompare.dll; DestDir: {app}; Flags: regserver ignoreversion
Source: mydoom_example.mdb; DestDir: {app}
Source: Readme.chm; DestDir: {app}
Source: signatures.mdb; DestDir: {app}
Source: ./dependancy\mscomctl.ocx; DestDir: {win}; Flags: uninsneveruninstall regserver
Source: ./dependancy\msscript.ocx; DestDir: {win}; Flags: uninsneveruninstall regserver
Source: ./dependancy\richtx32.ocx; DestDir: {win}; Flags: uninsneveruninstall regserver
Source: ./dependancy\vbDevKit.dll; DestDir: {win}; Flags: uninsneveruninstall regserver
Source: ./dependancy\IDAClientLib.dll; DestDir: {app}; Flags: regserver
Source: ./WinMerge_Plugin\wmIDACompare.dll; DestDir: {app}\WinMerge_Plugin\;  Flags: regserver

[Dirs]
Name: {app}\WinMerge_Plugin

[Run]
Filename: {app}\Readme.chm; WorkingDir: {app}; StatusMsg: View ReadMe File; Flags: shellexec postinstall

[Icons]
Name: {group}\IDACompare.exe; Filename: {app}\ida_compare.exe; WorkingDir: {app}
Name: {group}\ReadMe.chm; Filename: {app}\Readme.chm
Name: {group}\MyDoom Example DataBase; Filename: {app}\ida_compare.exe; Parameters: mydoom_example.mdb; WorkingDir: {app};
Name: {group}\Uninstall; Filename: {app}\unins000.exe; WorkingDir: {app};

[CustomMessages]
NameAndVersion=%1 version %2
AdditionalIcons=Additional icons:
CreateDesktopIcon=Create a &desktop icon
CreateQuickLaunchIcon=Create a &Quick Launch icon
ProgramOnTheWeb=%1 on the Web
UninstallProgram=Uninstall %1
LaunchProgram=Launch %1
AssocFileExtension=&Associate %1 with the %2 file extension
AssocingFileExtension=Associating %1 with the %2 file extension...
