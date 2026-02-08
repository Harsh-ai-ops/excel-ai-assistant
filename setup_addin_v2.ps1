$ErrorActionPreference = "Continue"
$folderPath = "C:\ExcelAddins"
$sourceManifest = "c:\Users\harsh\Downloads\files\excel-ai-assistant\manifest.xml"
$compName = $env:COMPUTERNAME

Write-Host "Starting Auto-Setup v2 (Using Computer Name: $compName)..."

# 1. Kill Excel
Write-Host "Closing Excel..."
Get-Process Excel -ErrorAction SilentlyContinue | Stop-Process -Force

# 2. Ensure Folder & Manifest
if (-not (Test-Path $folderPath)) { New-Item -ItemType Directory -Force -Path $folderPath | Out-Null }
Copy-Item -Path $sourceManifest -Destination "$folderPath\manifest.xml" -Force

# 3. Share
$shareName = "ExcelAddins"
# (Share creation logic same as before, likely already exists)
if (Get-Command New-SmbShare -ErrorAction SilentlyContinue) {
    if (-not (Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue)) {
        New-SmbShare -Name $shareName -Path $folderPath -FullAccess "Everyone" -Description "Excel AI Assistant" -ErrorAction SilentlyContinue | Out-Null
    }
}

# 4. Use COMPUTER NAME for Catalog URL
$catalogUrl = "\\$compName\$shareName"
Write-Host "Using Catalog URL: $catalogUrl"

# 5. Update Registry
$catalogGuid = "{F831E07E-4FBD-4EA6-8C37-1234567890AB}"
$regPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security\Trusted Catalogs\$catalogGuid"

if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
Set-ItemProperty -Path $regPath -Name "Url" -Value $catalogUrl
Set-ItemProperty -Path $regPath -Name "Flags" -Value 1 -Type DWord

Write-Host "✅ Registry Updated!"
Write-Host "✅ Catalog set to: $catalogUrl"
Write-Host "--------------------------------"
Write-Host "👉 NOW OPEN EXCEL MANUALLY."
