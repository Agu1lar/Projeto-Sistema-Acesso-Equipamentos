[Setup]
AppName=Sistema de Manutencao
AppVersion=1.0
DefaultDirName={pf}\SistemaManutencao
DefaultGroupName=Sistema Manutencao
OutputDir=instalador
OutputBaseFilename=SetupSistema
Compression=lzma
SolidCompression=yes

[Files]
Source: "dist\RelatorioFotografico.exe"; DestDir: "{app}"
Source: "Logo.png"; DestDir: "{app}"

[Icons]
Name: "{group}\Sistema de Manutencao"; Filename: "{app}\RelatorioFotografico.exe"
Name: "{commondesktop}\Sistema de Manutencao"; Filename: "{app}\RelatorioFotografico.exe"
