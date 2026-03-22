#Requires -RunAsAdministrator

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ScriptSourceDir = Split-Path -Path $PSCommandPath -Parent
$MainScriptSourcePath = Join-Path $ScriptSourceDir "BingSpotlight.ps1"
$UninstallScriptSourcePath = Join-Path $ScriptSourceDir "Uninstall-BingSpotlight.ps1"
$InstallDir = "C:\ProgramData\BingSpotlight"
$MainScriptInstallPath = Join-Path $InstallDir "BingSpotlight.ps1"
$UninstallScriptInstallPath = Join-Path $InstallDir "Uninstall-BingSpotlight.ps1"
$ConfigPath = Join-Path $InstallDir "config.json"

function Read-RetentionDays {
    param(
        [int]$DefaultValue = 14
    )

    while ($true) {
        $response = Read-Host ("Combien de jours conserver les images ? [default: {0}]" -f $DefaultValue)

        if ([string]::IsNullOrWhiteSpace($response)) {
            return $DefaultValue
        }

        $parsedValue = 0
        if ([int]::TryParse($response, [ref]$parsedValue) -and $parsedValue -ge 1) {
            return $parsedValue
        }

        Write-Host "Saisis un entier superieur ou egal a 1."
    }
}

function Read-Market {
    param(
        [string]$DefaultValue = "fr-FR"
    )

    while ($true) {
        $response = Read-Host ("Quel marche Bing utiliser ? ex: fr-FR, en-US, de-DE [default: {0}]" -f $DefaultValue)

        if ([string]::IsNullOrWhiteSpace($response)) {
            return $DefaultValue
        }

        $market = $response.Trim()
        if ($market -match '^[a-z]{2}-[A-Z]{2}$') {
            return $market
        }

        Write-Host "Saisis une valeur du type fr-FR, en-US ou de-DE."
    }
}

function Ensure-InstallFolders {
    foreach ($path in @($InstallDir, (Join-Path $InstallDir "logs"), (Join-Path $InstallDir "source"), (Join-Path $InstallDir "rendered"))) {
        if (-not (Test-Path -LiteralPath $path)) {
            New-Item -ItemType Directory -Path $path -Force | Out-Null
        }
    }
}

function Save-Config {
    param(
        [int]$RetentionDays,
        [string]$Market
    )

    $config = [pscustomobject]@{
        Market = $Market
        RetentionDays = $RetentionDays
        RetryCount = 5
        RetryDelaySeconds = 15
    }

    $config | ConvertTo-Json | Set-Content -LiteralPath $ConfigPath -Encoding ASCII
}

function Install-MainScript {
    if (-not (Test-Path -LiteralPath $MainScriptSourcePath)) {
        throw "Source script not found: $MainScriptSourcePath"
    }

    $resolvedSource = [System.IO.Path]::GetFullPath($MainScriptSourcePath)
    $resolvedDestination = [System.IO.Path]::GetFullPath($MainScriptInstallPath)

    if ($resolvedSource -ne $resolvedDestination) {
        Copy-Item -LiteralPath $MainScriptSourcePath -Destination $MainScriptInstallPath -Force
    }
}

function Install-UninstallScript {
    if (-not (Test-Path -LiteralPath $UninstallScriptSourcePath)) {
        throw "Source script not found: $UninstallScriptSourcePath"
    }

    $resolvedSource = [System.IO.Path]::GetFullPath($UninstallScriptSourcePath)
    $resolvedDestination = [System.IO.Path]::GetFullPath($UninstallScriptInstallPath)

    if ($resolvedSource -ne $resolvedDestination) {
        Copy-Item -LiteralPath $UninstallScriptSourcePath -Destination $UninstallScriptInstallPath -Force
    }
}

function Register-BingTask {
    $action = New-ScheduledTaskAction `
        -Execute "powershell.exe" `
        -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -File `"$MainScriptInstallPath`""

    $triggerLogon = New-ScheduledTaskTrigger -AtLogOn
    $triggerLogon.Delay = "PT90S"

    $triggerDaily = New-ScheduledTaskTrigger -Daily -At 10:00AM

    $principal = New-ScheduledTaskPrincipal `
        -UserId "SYSTEM" `
        -LogonType ServiceAccount `
        -RunLevel Highest

    $settings = New-ScheduledTaskSettingsSet `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 10) `
        -MultipleInstances IgnoreNew `
        -StartWhenAvailable `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries

    Register-ScheduledTask `
        -TaskName "BingSpotlight_LockScreen" `
        -TaskPath "\Custom\" `
        -Action $action `
        -Trigger @($triggerLogon, $triggerDaily) `
        -Principal $principal `
        -Settings $settings `
        -Description "Telecharge l'image Bing du jour et l'applique au lock screen LTSC" `
        -Force | Out-Null
}

try {
    $retentionDays = Read-RetentionDays
    $market = Read-Market

    Ensure-InstallFolders
    Install-MainScript
    Install-UninstallScript
    Save-Config -RetentionDays $retentionDays -Market $market
    Register-BingTask

    Write-Host "Installation terminee."
    Write-Host ("Retention configuree : {0} jour(s)." -f $retentionDays)
    Write-Host ("Marche Bing configure : {0}" -f $market)
    Write-Host ("Script installe dans : {0}" -f $MainScriptInstallPath)
    Write-Host ("Desinstallation disponible dans : {0}" -f $UninstallScriptInstallPath)
    Write-Host ("Configuration ecrite dans : {0}" -f $ConfigPath)
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}