;InnoSetupVersion=4.2.6

[Setup]
AppName=IDACompare
AppVerName=IDACompare v0.3
DefaultDirName=c:\iDefense\IDACompare\
DefaultGroupName=IDACompare
OutputBaseFilename=IDACompare_Setup
OutputDir=./

[Files]
Source: ./IDA_Compare.plw; DestDir: {app}
Source: ./install_plw.bat; DestDir: {app}
Source: ./\vb src\Module1.bas; DestDir: {app}\vb src
Source: ./\vb src\frmPluginSample.frm; DestDir: {app}\vb src
Source: ./\vb src\CPlugin.cls; DestDir: {app}\vb src
Source: ./\vb src\Ida_Compare.vbp; DestDir: {app}\vb src
Source: ./\vb src\clsCmnDlg.cls; DestDir: {app}\vb src
Source: ./\vb src\Ida_Compare.vbw; DestDir: {app}\vb src
Source: ./\vc src\idacompare.cpp; DestDir: {app}\vc src
Source: ./\vc src\idacompare.def; DestDir: {app}\vc src
Source: ./\vc src\idacompare.vcproj; DestDir: {app}\vc src
Source: ./\vc src\idacompare.sln; DestDir: {app}\vc src
Source: ./\blank.mdb; DestDir: {app}
Source: ./\compare.vbs; DestDir: {app}
Source: ./\ida_compare.exe; DestDir: {app}; Flags: ignoreversion
Source: ./\compare\crc.cls; DestDir: {app}\compare
Source: ./\compare\clsCmnDlg.cls; DestDir: {app}\compare
Source: ./\compare\Form1.frm; DestDir: {app}\compare
Source: ./\compare\Project1.vbp; DestDir: {app}\compare
Source: ./\compare\CAlphaSort.cls; DestDir: {app}\compare
Source: ./\compare\Project1.vbw; DestDir: {app}\compare
Source: ./\compare\Module1.bas; DestDir: {app}\compare
Source: ./\compare\Module2.bas; DestDir: {app}\compare
Source: ./\compare\CFunction.cls; DestDir: {app}\compare
Source: ./\compare\frmProfile.frm; DestDir: {app}\compare
Source: ./\compare\frmFind.frm; DestDir: {app}\compare
Source: ./\IdaCompare.dll; DestDir: {app}; Flags: regserver ignoreversion
Source: ./\mydoom_example.mdb; DestDir: {app}
Source: ./\Readme.chm; DestDir: {app}
Source: ./\signatures.mdb; DestDir: {app}
Source: ./\iDefense Labs.url; DestDir: {app}
Source: ./dependancy\mscomctl.ocx; DestDir: {win}; Flags: uninsneveruninstall regserver
Source: ./dependancy\msscript.ocx; DestDir: {win}; Flags: uninsneveruninstall regserver
Source: ./dependancy\richtx32.ocx; DestDir: {win}; Flags: uninsneveruninstall regserver

[Dirs]
Name: {app}\vb src
Name: {app}\vc src
Name: {app}\compare

[Run]
Filename: {app}\Readme.chm; WorkingDir: {app}; StatusMsg: View ReadMe File; Flags: shellexec postinstall

[Icons]
Name: {group}\IDACompare.exe; Filename: {app}\ida_compare.exe; WorkingDir: {app}
Name: {group}\ReadMe.chm; Filename: {app}\Readme.chm
Name: {group}\Example DB; Filename: {app}\ida_compare.exe; Parameters: mydoom_example.mdb; WorkingDir: {app}; IconFilename: {app}\mydoom_example.mdb
Name: {group}\Source\Plugin..vcproj; Filename: {app}\vc src\idacompare.vcproj
Name: {group}\Source\Compare.vbp; Filename: {app}\compare\Project1.vbp
Name: {group}\Source\PluginUI.vbp; Filename: {app}\vb src\Ida_Compare.vbp
Name: {group}\Uninstall; Filename: unins000.exe
;Name: {group}\labs.iDefense.com Website; Filename: {app}\iDefense Labs.url; WorkingDir: {app}

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
