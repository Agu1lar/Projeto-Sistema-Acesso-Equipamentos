[Setup]
AppName=Sistema de Manutenção
AppVersion=1.0
DefaultDirName={pf}\SistemaManutencao
DefaultGroupName=Sistema Manutenção
OutputDir=instalador
OutputBaseFilename=SetupSistema
Compression=lzma
SolidCompression=yes

[Files]
Source: "dist\main.exe"; DestDir: "{app}"
Source: "Logo.png"; DestDir: "{app}"

[Icons]
Name: "{group}\Sistema de Manutenção"; Filename: "{app}\main.exe"
Name: "{commondesktop}\Sistema de Manutenção"; Filename: "{app}\main.exe"