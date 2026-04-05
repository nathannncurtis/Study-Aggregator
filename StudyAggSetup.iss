; Inno Setup script for Study Aggregator (public edition)

#ifndef MyAppVersion
  #define MyAppVersion "4.0.0"
#endif

#define MyAppName "Study Aggregator"
#define MyAppPublisher "Nathan Curtis"
#define MyAppURL "https://github.com/nathannncurtis/Study-Aggregator"
#define MyAppExeName "Study Aggregator.exe"

[Setup]
AppId={{8B42C1F4-9DA7-4E9B-9F8C-1F4B2E5A6C7D}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={userappdata}\{#MyAppName}
DisableDirPage=yes
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
Compression=lzma
SolidCompression=yes
OutputBaseFilename=StudyAggregatorSetup
WizardStyle=dynamic

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "dist\Study Aggregator\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\Study Aggregator\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
var
  Msg: String;
begin
  if CurStep = ssPostInstall then
  begin
    Msg :=
      'Study Aggregator 4.0 - What''s New' + #13#10 + #13#10 +
      'This release replaces the Python DICOM parser with a new engine written in Rust. ' +
      'The hot path - directory scanning, DICOM tag extraction, ZIP handling, and patient merging - ' +
      'now runs as a native binary (study-agg-engine.exe) that the GUI launches as a subprocess.' + #13#10 + #13#10 +
      'Why it''s faster:' + #13#10 +
      '  - Zero-copy parsing: files are memory-mapped and tags are read in place, with no Python object allocation per file.' + #13#10 +
      '  - Parallelism: directory walking (jwalk) and DICOM parsing (rayon) now use every CPU core instead of one.' + #13#10 +
      '  - Streaming ZIP handling: archive entries are parsed as they''re decompressed and dropped immediately, so a 4GB+ ZIP no longer holds the whole archive in RAM.' + #13#10 +
      '  - No more pydicom / numpy / pyzipper: the installer is significantly smaller and startup is faster.' + #13#10 + #13#10 +
      'What you should notice:' + #13#10 +
      '  - Large folders and ZIPs finish in a fraction of the previous time.' + #13#10 +
      '  - RAM usage stays flat even on multi-gigabyte archives.' + #13#10 +
      '  - Errors from the engine now surface as popups instead of being silently swallowed.';
    MsgBox(Msg, mbInformation, MB_OK);
  end;
end;
