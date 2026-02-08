$ErrorActionPreference = "Continue" # Don't stop on non-fatal errors
$folderPath = "C:\ExcelAddins"
$sourceManifest = "c:\Users\harsh\Downloads\files\excel-ai-assistant\manifest.xml"

Write-Host "Starting Auto-Setup..."

# 1. Create Folder
if (-not (Test-Path $folderPath)) {
    New-Item -ItemType Directory -Force -Path $folderPath | Out-Null
    Write-Host "Created $folderPath"
}

# 2. Copy Manifest
Copy-Item -Path $sourceManifest -Destination "$folderPath\manifest.xml" -Force
Write-Host "Copied manifest.xml"

# 3. Try to Create Share (Requires Admin)
$shareName = "ExcelAddins"
$catalogUrl = ""

try {
    if (Get-Command New-SmbShare -ErrorAction SilentlyContinue) {
        if (-not (Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue)) {
            New-SmbShare -Name $shareName -Path $folderPath -FullAccess "Everyone" -Description "Excel AI Assistant" -ErrorAction Stop | Out-Null
            Write-Host "Created SMB Share: $shareName"
        } else {
             Write-Host "Share already exists"
        }
        $catalogUrl = "\\localhost\$shareName"
    } else {
        throw "New-SmbShare not available"
    }
} catch {
    Write-Host "⚠️  Unable to create network share (likely needs Admin). Using local path fallback (might not work)."
    $catalogUrl = $folderPath
}

Write-Host "Using Catalog URL: $catalogUrl"

# 4. Update Registry
# GUID for our catalog
$catalogGuid = "{F831E07E-4FBD-4EA6-8C37-1234567890AB}"
$regPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security\Trusted Catalogs\$catalogGuid"

try {
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }
    
    Set-ItemProperty -Path $regPath -Name "Url" -Value $catalogUrl
    Set-ItemProperty -Path $regPath -Name "Flags" -Value 1 -Type DWord # 1 = Show in Menu
    
    Write-Host "✅ Registry Updated Successfully!"
} catch {
    Write-Host "❌ Failed to update registry: $_"
}

Write-Host "DONE. Please Restart Excel."
