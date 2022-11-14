;#define TEST
#define ProgramVersion '221111P'
#define RevitVersionS '2022:2023'
#define ProgName 'LabelPlacer'

#ifdef TEST
#define RevitAddinDir 'LIMA'
#define RevitAddinSourceDir 'Debug'
#else
#define RevitAddinDir 'Autodesk'
#define RevitAddinSourceDir 'Debug'
#endif

[Setup]
AppName=Label Placer
AppVerName=Label Placer {#ProgramVersion}
WizardStyle=modern
OutputBaseFilename={#ProgName}
DefaultDirName={commonappdata}\{#RevitAddinDir}\RevitAddins\{#ProgName}
Compression=lzma2
SolidCompression=yes
DisableDirPage=no

[Files]
;C:\ProgramData\LIMA\Revit\Addins\2022
Source: "{#ProgName}.addin"; DestDir: "{tmp}"; Flags: dontcopy;
;C:\ProgramData\LIMA\RevitAddins\2022
Source: "..\bin\{#RevitAddinSourceDir}\{#ProgName}.dll"; DestDir: "{tmp}"; Flags: dontcopy;
;Source: "..\bin\Debug\LabelPlacer.dll";

[Dirs]
Name: "{commonappdata}\{#RevitAddinDir}\RevitAddins\LabelPlacer"

[UninstallDelete]
Type: filesandordirs; Name: "{commonappdata}\{#RevitAddinDir}\RevitAddins\{#ProgName}";
Type: dirifempty; Name: "{commonappdata}\{#RevitAddinDir}\RevitAddins";

[Code]
function StrSplit(Text: String; Separator: String): TArrayOfString;
var
  i, p: Integer;
  Dest: Array of String; 
begin
  i := 0;
  repeat
    SetArrayLength(Dest, i+1);
    p := Pos(Separator,Text);
    if p > 0 then begin
      Dest[i] := Copy(Text, 1, p-1);
      Text := Copy(Text, p + Length(Separator), Length(Text));
      i := i + 1;
    end else begin
      Dest[i] := Text;
      Text := '';
    end;
  until Length(Text)=0;
  Result := Dest
end;

function CheckRevitVersionS(): Boolean;
var
  A: AnsiString;
  XMLDocument: Variant;  
  XMLNode: Variant;
  RevitVersionS: Array of String;
  RevitVersion: String;
  I: Integer;
begin
  RevitVersionS := StrSplit (ExpandConstant('{#RevitVersionS}'), ':');
  ExtractTemporaryFiles('{tmp}\LabelPlacer.addin');
  ExtractTemporaryFiles('{tmp}\LabelPlacer.dll');

  for I := 0 to GetArrayLength(RevitVersionS) - 1 do
  begin
    RevitVersion := RevitVersionS[I];
    Result := RegKeyExists(HKCU, 'SOFTWARE\Autodesk\Revit\' + RevitVersion);

    if Result then
    begin
      LoadStringFromFile(ExpandConstant('{tmp}\LabelPlacer.addin'), A);
      XMLDocument := CreateOleObject('Msxml2.DOMDocument.6.0');
      XMLDocument.async := False;
      XMLDocument.LoadXml(A);
      XMLDocument.setProperty('SelectionLanguage', 'XPath');
      XMLNode := XMLDocument.selectSingleNode('//RevitAddIns/AddIn/Assembly');
      XMLNode.text := ExpandConstant('{commonappdata}\{#RevitAddinDir}\RevitAddins\LabelPlacer\' + RevitVersion + '\LabelPlacer.dll');
      XMLDocument.save(ExpandConstant('{commonappdata}\{#RevitAddinDir}\Revit\Addins\' + RevitVersion + '\LabelPlacer.addin'));

      //CreateDir(ExpandConstant('{commonappdata}\{#RevitAddinDir}\RevitAddins'));
      //CreateDir(ExpandConstant('{commonappdata}\{#RevitAddinDir}\RevitAddins\LabelPlacer'));
      CreateDir(ExpandConstant('{commonappdata}\{#RevitAddinDir}\RevitAddins\LabelPlacer\') + RevitVersion);
      FileCopy(ExpandConstant('{tmp}\LabelPlacer.dll'), ExpandConstant('{commonappdata}\{#RevitAddinDir}\RevitAddins\LabelPlacer\') + RevitVersion + '\LabelPlacer.dll', FALSE);
    end;
  end;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
var
  I, Max: Integer;
begin
  if CurPageID = wpReady then begin
    //CheckRevitVersionS()
  end;
  Result := True;
end;

function RemoveRevitAddIn(VerRevit: String): Boolean;
var
  RevitVersionS: Array of String;
  I: Integer;
  RevitVersion: String;
begin
  RevitVersionS := StrSplit (ExpandConstant('{#RevitVersionS}'), ':');

  for I := 0 to GetArrayLength(RevitVersionS) - 1 do
  begin
    RevitVersion := RevitVersionS[I];

    DeleteFile(ExpandConstant('{commonappdata}\{#RevitAddinDir}\Revit\Addins\' + RevitVersion + '\LabelPlacer.addin'));  
  end;

  DelTree(ExpandConstant('{commonappdata}\{#RevitAddinDir}\RevitAddins\LabelPlacer'), TRUE, TRUE, TRUE);
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  I, Max: Integer;
begin
  if CurStep = ssPostInstall then 
  begin
    CheckRevitVersionS()
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  I, Max: Integer;
begin
  if CurUninstallStep = usUninstall then 
  begin
    RemoveRevitAddIn(ExpandConstant('{#RevitVersionS}'));
  end;
end;
