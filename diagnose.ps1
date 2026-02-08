$ErrorActionPreference = "Continue"

Write-Host "--- DIAGNOSTIC REPORT ---"

# 1. Computer Details
$compName = $env:COMPUTERNAME
Write-Host "Computer Name: $compName"
Write-Host "User: $env:USERNAME"

# 2. Check WEF\Developer (Dev Mode)
$devKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
Write-Host "`n[Checking Developer Mode Registry]"
if (Test-Path $devKey) {
    Get-Item $devKey | Select-Object -ExpandProperty Property | ForEach-Object {
        $val = Get-ItemProperty -Path $devKey -Name $_
        Write-Host "  Found Manifest ID: $_"
        Write-Host "  Url: $($val.$_)"
    }
}
else {
    Write-Host "  WEF\Developer key NOT found."
}

# 3. Check Trusted Catalogs (Shared Folder Mode)
$catKey = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security\Trusted Catalogs"
Write-Host "`n[Checking Trusted Catalogs]"
if (Test-Path $catKey) {
    Get-ChildItem $catKey | ForEach-Object {
        $props = Get-ItemProperty -Path $_.PSPath
        Write-Host "  Catalog ID: $($_.PSChildName)"
        Write-Host "  Url: $($props.Url)"
        Write-Host "  Flags: $($props.Flags) (Should be 1)"
    }
}
else {
    Write-Host "  Trusted Catalogs key NOT found."
}

# 4. Check Network Shares
Write-Host "`n[Checking Network Shares]"
$sharePathLocal = "\\localhost\ExcelAddins"
$sharePathComp = "\\$compName\ExcelAddins"

if (Test-Path $sharePathLocal) { Write-Host "  Accessible: $sharePathLocal" } else { Write-Host "  ❌ NOT Accessible: $sharePathLocal" }
if (Test-Path $sharePathComp) { Write-Host "  Accessible: $sharePathComp" } else { Write-Host "  ❌ NOT Accessible: $sharePathComp" }

# 5. Check File Existence
$localFile = "C:\ExcelAddins\manifest.xml"
if (Test-Path $localFile) { Write-Host "`n  Manifest exists at $localFile" } else { Write-Host "`n  ❌ Manifest MISSING at $localFile" }

Write-Host "`n--- END REPORT ---"
