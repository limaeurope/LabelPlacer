;#define TEST

#ifdef TEST
#define RevitAddinDir "LIMA"
#else
#define RevitAddinDir "Autodesk"
#endif

[Setup]
AppName=Label Placer
AppVerName=Label Placer 221111P
WizardStyle=modern
DefaultGroupName=My Program
DefaultDirName={commonappdata}\{#RevitAddinDir}\RevitAddins\LabelPlacer
Compression=lzma2
SolidCompression=yes
DisableDirPage=no

[Files]
;C:\ProgramData\LIMA\Revit\Addins\2022
Source: "LabelPlacer.addin"; DestDir: "{tmp}"; Flags: dontcopy
;C:\ProgramData\LIMA\RevitAddins\2022
Source: "..\bin\Debug\LabelPlacer.dll"; DestDir: "{commonappdata}\LIMA\RevitAddins\LabelPlacer\2022"

[Dirs]
Name: "{commonappdata}\{#RevitAddinDir}\Revit\Addins\LabelPlacer\2022"

[Code]
procedure InitializeWizard;
var
A: AnsiString;
XMLDocument: Variant;  
XMLNode: Variant;  
begin
    ExtractTemporaryFiles('{tmp}\LabelPlacer.addin');
    LoadStringFromFile(ExpandConstant('{tmp}\LabelPlacer.addin'), A);
    XMLDocument := CreateOleObject('Msxml2.DOMDocument.6.0');
    XMLDocument.async := False;
    XMLDocument.LoadXml(A);
    XMLDocument.setProperty('SelectionLanguage', 'XPath');
    XMLNode := XMLDocument.selectSingleNode('//RevitAddIns/AddIn/Assembly');
    XMLNode.text := ExpandConstant('{commonappdata}\{#RevitAddinDir}\Revit\Addins\2022\LabelPlacer\LabelPlacer.dll');
    XMLDocument.save(ExpandConstant('{commonappdata}\{#RevitAddinDir}\Revit\Addins\2022\LabelPlacer.addin'));
end;

