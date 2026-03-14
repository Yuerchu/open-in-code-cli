# uninstall.ps1 - Remove context menu entries
# Auto-elevates to admin if needed

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$logFile   = Join-Path $scriptDir "uninstall.log"

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

$results = [System.Collections.ArrayList]::new()
function Log($msg) { [void]$results.Add($msg); Write-Host $msg }

try {
    # 1. Remove AppxPackage
    $pkg = Get-AppxPackage -Name "OpenInCodeCLI.ContextMenu" -ErrorAction SilentlyContinue
    if ($pkg) {
        Remove-AppxPackage $pkg.PackageFullName
        Log "Removed package: $($pkg.PackageFullName)"
    } else {
        Log "Package not found (already removed)."
    }

    # 2. Remove certificate
    $certSubject = "CN=OpenInCodeCLI"
    $removed = $false
    foreach ($store in @("Cert:\LocalMachine\TrustedPeople", "Cert:\CurrentUser\My")) {
        Get-ChildItem $store -ErrorAction SilentlyContinue |
            Where-Object { $_.Subject -eq $certSubject } |
            ForEach-Object { Remove-Item $_.PSPath -Force; $removed = $true }
    }
    if ($removed) { Log "Removed certificates." } else { Log "No certificates found." }

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
            Log "Removed: $k"
        }
    }

    # 4. Clean up MSIX file
    $msixPath = Join-Path $scriptDir "OpenInCodeCLI.msix"
    if (Test-Path $msixPath) { Remove-Item $msixPath -Force; Log "Removed MSIX file." }

    # 5. Restart Explorer
    Log "Restarting Explorer..."
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    Start-Process explorer

    Log ""
    Log "=== Uninstall complete! ==="
}
catch {
    Log "ERROR: $($_.Exception.Message)"
}

$results | Out-File -FilePath $logFile -Encoding UTF8
