; This is an Inno Setup configuration file
; https://jrsoftware.org/isinfo.php

#define ApplicationVersion GetFileVersion('..\bin\VennDiagramPlotter.exe')

[CustomMessages]
AppName=Venn Digram Plotter

[Messages]
WelcomeLabel2=This will install [name/ver] on your computer.
; Example with multiple lines:
; WelcomeLabel2=Welcome message%n%nAdditional sentence

[Files]
Source: ..\bin\VennDiagramPlotter.exe             ; DestDir: {app}

Source: ..\bin\ControlPrinter.dll                 ; DestDir: {app}
Source: ..\bin\PRISM.dll                          ; DestDir: {app}
Source: ..\bin\stdole.dll                         ; DestDir: {app}
Source: ..\bin\VennDiagrams.dll                   ; DestDir: {app}

Source: ..\bin\VennDiagramPlotter_Settings.xml    ; DestDir: {app}
Source: ..\bin\Venn_Diagram_Examples.ppt          ; DestDir: {app}

Source: ..\Readme.txt                             ; DestDir: {app}
Source: ..\RevisionHistory.txt                    ; DestDir: {app}
Source: Images\delete_16x.ico                     ; DestDir: {app}

[Dirs]
Name: {commonappdata}\VennDiagramPlotter; Flags: uninsalwaysuninstall

[Tasks]
Name: desktopicon; Description: {cm:CreateDesktopIcon}; GroupDescription: {cm:AdditionalIcons}; Flags: unchecked
; Name: quicklaunchicon; Description: {cm:CreateQuickLaunchIcon}; GroupDescription: {cm:AdditionalIcons}; Flags: unchecked

[Icons]
Name: {commondesktop}\Venn Diagram Plotter; Filename: {app}\VennDiagramPlotter.exe; Tasks: desktopicon; Comment: VennDiagramPlotter
Name: {group}\Venn Diagram Plotter; Filename: {app}\VennDiagramPlotter.exe; Comment: Venn Diagram Plotter

[Setup]
AppName=Venn Diagram Plotter
AppVersion={#ApplicationVersion}
;AppVerName=VennDiagramPlotter
AppID=VennDiagramPlotterId
AppPublisher=Pacific Northwest National Laboratory
AppPublisherURL=https://omics.pnl.gov/software
AppSupportURL=https://omics.pnl.gov/software
AppUpdatesURL=https://github.com/PNNL-Comp-Mass-Spec/Venn-Diagram-Plotter
ArchitecturesAllowed=x64 x86
ArchitecturesInstallIn64BitMode=x64
DefaultDirName={autopf}\VennDiagramPlotter
DefaultGroupName=PAST Toolkit
AppCopyright=© PNNL
;LicenseFile=.\License.rtf
PrivilegesRequired=admin
OutputBaseFilename=VennDiagramPlotter_Installer
VersionInfoVersion={#ApplicationVersion}
VersionInfoCompany=PNNL
VersionInfoDescription=Venn Diagram Plotter
VersionInfoCopyright=PNNL
DisableFinishedPage=yes
DisableWelcomePage=no
ShowLanguageDialog=no
ChangesAssociations=no
WizardStyle=modern
EnableDirDoesntExistWarning=no
AlwaysShowDirOnReadyPage=yes
UninstallDisplayIcon={app}\delete_16x.ico
ShowTasksTreeLines=yes
OutputDir=.\Output

[Registry]
;Root: HKCR; Subkey: MyAppFile; ValueType: string; ValueName: ; ValueDataMyApp File; Flags: uninsdeletekey
;Root: HKCR; Subkey: MyAppSetting\DefaultIcon; ValueType: string; ValueData: {app}\wand.ico,0; Flags: uninsdeletevalue

[UninstallDelete]
Name: {app}; Type: filesandordirs
