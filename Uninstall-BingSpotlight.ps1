#Requires -RunAsAdministrator

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$InstallDir = "C:\ProgramData\BingSpotlight"

$TaskName = "BingSpotlight_LockScreen"
$TaskPath = "\Custom\"
$RegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP"

function Remove-BingTask {
    $task = Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue
    if ($task) {
        Unregister-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Confirm:$false
    }
}

function Clear-LockScreenConfiguration {
    if (-not (Test-Path -LiteralPath $RegPath)) {
        return
    }

    $currentPath = $null
    try {
        $currentPath = (Get-ItemProperty -Path $RegPath -Name "LockScreenImagePath" -ErrorAction SilentlyContinue).LockScreenImagePath
    }
    catch {
    }

    $normalizedInstallDir = [System.IO.Path]::GetFullPath($InstallDir).TrimEnd('\\')
    $normalizedCurrentPath = $null

    if (-not [string]::IsNullOrWhiteSpace($currentPath)) {
        try {
            $normalizedCurrentPath = [System.IO.Path]::GetFullPath($currentPath)
        }
        catch {
        }
    }

    if ($normalizedCurrentPath -and $normalizedCurrentPath.StartsWith($normalizedInstallDir, [System.StringComparison]::OrdinalIgnoreCase)) {
        Remove-ItemProperty -Path $RegPath -Name "LockScreenImagePath" -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path $RegPath -Name "LockScreenImageUrl" -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path $RegPath -Name "LockScreenImageStatus" -ErrorAction SilentlyContinue
    }
}

function Remove-InstallDirectory {
    if (Test-Path -LiteralPath $InstallDir) {
        Remove-Item -LiteralPath $InstallDir -Recurse -Force
    }
}

function Confirm-Uninstall {
    $answer = Read-Host "Confirmer la desinstallation complete de BingSpotlight ? Tape OUI pour continuer"
    return $answer -ceq "OUI"
}

try {
    if (-not (Confirm-Uninstall)) {
        Write-Host "Desinstallation annulee."
        exit 0
    }

    Remove-BingTask
    Clear-LockScreenConfiguration
    Remove-InstallDirectory

    Write-Host "Desinstallation terminee."
    Write-Host ("Dossier supprime : {0}" -f $InstallDir)
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}