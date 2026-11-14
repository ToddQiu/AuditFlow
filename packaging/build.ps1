param([string]$Python = 'python', [string]$ISCC = 'C:\Program Files (x86)\Inno Setup 6\ISCC.exe')
$ErrorActionPreference = 'Stop'
Set-Location (Split-Path $PSScriptRoot -Parent)
& $Python -m PyInstaller --noconfirm --clean --windowed --onedir --name AuditFlow --icon assets/audit-a-20261011.ico --add-data 'web;web' --add-data 'assets/audit-a-20261011.ico;assets' --add-data 'research/read_corpus.py;research' --hidden-import openpyxl --hidden-import xlrd --hidden-import olefile --hidden-import pypdf --hidden-import pymupdf --hidden-import lxml.etree --collect-all webview start_workbench.py
if ($LASTEXITCODE -ne 0) { throw 'PyInstaller failed' }
& $ISCC packaging/auditflow.iss
if ($LASTEXITCODE -ne 0) { throw 'Installer compilation failed' }
Get-FileHash dist/AuditFlow-Setup-0.1.0.exe -Algorithm SHA256 | ForEach-Object { "$($_.Hash.ToLower())  AuditFlow-Setup-0.1.0.exe" } | Set-Content dist/SHA256SUMS.txt -Encoding ascii
