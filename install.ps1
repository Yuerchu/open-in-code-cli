# install.ps1 - Build, package, and install Win11 context menu for Claude Code & Codex
# Auto-elevates to admin. Just run: powershell -ExecutionPolicy Bypass -File install.ps1

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$distDir   = Join-Path $scriptDir "dist"
$logFile   = Join-Path $scriptDir "install.log"

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "Requesting admin privileges..." -ForegroundColor Yellow
    $scriptPath = $MyInvocation.MyCommand.Definition
    Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -File `"$scriptPath`"" -Verb RunAs -Wait
    if (Test-Path $logFile) {
        Get-Content $logFile
        Remove-Item $logFile -Force
    }
    exit
}

# ===== Running as admin =====
$results = [System.Collections.ArrayList]::new()
function Log($msg) { [void]$results.Add($msg); Write-Host $msg }

function Find-SdkTool($toolName) {
    $sdkRoot = "C:\Program Files (x86)\Windows Kits\10\bin"
    if (-not (Test-Path $sdkRoot)) { return $null }
    $versions = Get-ChildItem $sdkRoot -Directory | Where-Object { $_.Name -match '^\d+\.' } | Sort-Object Name -Descending
    foreach ($v in $versions) {
        $p = Join-Path $v.FullName "x64\$toolName"
        if (Test-Path $p) { return $p }
    }
    return $null
}

try {
    # 1. Build DLL if needed
    $dllPath = Join-Path $distDir "context_menu.dll"
    if (-not (Test-Path $dllPath)) {
        Log "Building context_menu.dll ..."
        $buildBat = Join-Path $scriptDir "build.bat"
        if (-not (Test-Path $buildBat)) { throw "build.bat not found. Please run from the project root." }
        $buildOut = cmd /c "`"$buildBat`"" 2>&1
        if ($LASTEXITCODE -ne 0) {
            $buildOut | ForEach-Object { Log "  $_" }
            throw "Build failed. Make sure Visual Studio Build Tools 2022 is installed."
        }
        Log "Build complete."
    }

    # 2. Generate icons if missing
    $claudeIco = Join-Path $distDir "claude.ico"
    if (-not (Test-Path $claudeIco)) {
        Log "Generating icons..."
        $iconScript = Join-Path $scriptDir "create_icons.ps1"
        if (Test-Path $iconScript) {
            & $iconScript
            Log "Icons created."
        } else {
            throw "create_icons.ps1 not found and icons are missing."
        }
    }

    # 3. Remove old classic registry entries
    $oldKeys = @(
        "Registry::HKEY_CLASSES_ROOT\Directory\shell\ClaudeCode",
        "Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\ClaudeCode",
        "Registry::HKEY_CLASSES_ROOT\Directory\shell\Codex",
        "Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\Codex"
    )
    foreach ($k in $oldKeys) {
        if (Test-Path $k) {
            Remove-Item -Path $k -Recurse -Force
            Log "Cleaned up old registry: $k"
        }
    }

    # 4. Remove existing package
    $existing = Get-AppxPackage -Name "OpenInCodeCLI.ContextMenu" -ErrorAction SilentlyContinue
    if ($existing) {
        Log "Removing previous installation..."
        Remove-AppxPackage $existing.PackageFullName
    }

    # 5. Find SDK tools
    $makeappx = Find-SdkTool "makeappx.exe"
    $signtool = Find-SdkTool "signtool.exe"
    if (-not $makeappx) { throw "makeappx.exe not found. Install Windows 10/11 SDK." }
    if (-not $signtool) { throw "signtool.exe not found. Install Windows 10/11 SDK." }
    Log "SDK tools: $(Split-Path $makeappx -Parent)"

    # 6. Create self-signed certificate
    $certSubject = "CN=OpenInCodeCLI"
    $certStoreTrust = "Cert:\LocalMachine\TrustedPeople"
    $certStoreMy    = "Cert:\CurrentUser\My"

    # Clean old certs
    Get-ChildItem $certStoreTrust -ErrorAction SilentlyContinue |
        Where-Object { $_.Subject -eq $certSubject } | Remove-Item -Force
    Get-ChildItem $certStoreMy -ErrorAction SilentlyContinue |
        Where-Object { $_.Subject -eq $certSubject } | Remove-Item -Force

    Log "Creating self-signed certificate..."
    $cert = New-SelfSignedCertificate `
        -Type Custom `
        -Subject $certSubject `
        -KeyUsage DigitalSignature `
        -FriendlyName "Open in Code CLI" `
        -CertStoreLocation $certStoreMy `
        -TextExtension @("2.5.29.37={text}1.3.6.1.5.5.7.3.3", "2.5.29.19={text}")
    $thumbprint = $cert.Thumbprint

    # Trust the certificate
    $certFile = Join-Path $scriptDir "OpenInCodeCLI.cer"
    Export-Certificate -Cert "$certStoreMy\$thumbprint" -FilePath $certFile | Out-Null
    Import-Certificate -FilePath $certFile -CertStoreLocation $certStoreTrust | Out-Null
    Remove-Item $certFile -Force
    Log "Certificate ready: $thumbprint"

    # 7. Create MSIX
    $msixPath = Join-Path $scriptDir "OpenInCodeCLI.msix"
    Log "Packaging MSIX..."
    $out = & $makeappx pack /d "$distDir" /p "$msixPath" /nv /o 2>&1
    if ($LASTEXITCODE -ne 0) {
        $out | ForEach-Object { Log "  $_" }
        throw "makeappx failed"
    }

    # 8. Sign MSIX
    Log "Signing MSIX..."
    $out = & $signtool sign /fd SHA256 /sha1 $thumbprint /td SHA256 "$msixPath" 2>&1
    if ($LASTEXITCODE -ne 0) {
        $out | ForEach-Object { Log "  $_" }
        throw "signtool failed"
    }

    # 9. Install
    Log "Installing package..."
    Add-AppxPackage -Path $msixPath -ExternalLocation $distDir

    # 10. Restart Explorer
    Log "Restarting Explorer..."
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    Start-Process explorer

    Log ""
    Log "=== Done! ==="
    Log "Right-click any folder to see 'Claude Code' and 'Codex' in the context menu."
}
catch {
    Log "ERROR: $($_.Exception.Message)"
    Log $_.ScriptStackTrace
}

$results | Out-File -FilePath $logFile -Encoding UTF8
