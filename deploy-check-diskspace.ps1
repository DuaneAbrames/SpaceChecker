#!ps
#timeout=999999
[CmdletBinding()]
param(
    [string]$GitHubRepo = 'DuaneAbrames/SpaceChecker',
    [string]$AssetRelativePath = 'check-diskspace.ps1',
    [string]$FallbackBranch = 'main',
    [string]$TargetDirectory = 'C:\istools'
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$version = '0.0.0'
$releaseTag = $null
$apiUri = "https://api.github.com/repos/$GitHubRepo/releases/latest"
$downloadUri = $null

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

try {
    Write-Verbose "Querying GitHub release metadata from $apiUri"
    $headers = @{ 'User-Agent' = 'check-diskspace-deployer'; 'Accept' = 'application/vnd.github+json' }
    $latestRelease = Invoke-RestMethod -Uri $apiUri -Headers $headers -UseBasicParsing

    if ($latestRelease.tag_name) {
        $releaseTag = $latestRelease.tag_name
        $version = $latestRelease.tag_name.TrimStart('v')
    } elseif ($latestRelease.name) {
        $releaseTag = $latestRelease.name
        $version = $latestRelease.name.TrimStart('v')
    }

    if ($releaseTag) {
        $downloadUri = "https://raw.githubusercontent.com/$GitHubRepo/$releaseTag/$AssetRelativePath"
    }
} catch {
    Write-Warning "Unable to retrieve GitHub release information: $($_.Exception.Message)"
}

if (-not $downloadUri) {
    $releaseTag = $FallbackBranch
    if ($version -eq '0.0.0') {
        $version = Get-Date -Format 'yyyyMMddHHmmss'
    }
    $downloadUri = "https://raw.githubusercontent.com/$GitHubRepo/$FallbackBranch/$AssetRelativePath"
    Write-Verbose "Falling back to branch '$FallbackBranch' at $downloadUri"
}

$targetPath = Join-Path -Path $TargetDirectory -ChildPath ("check-diskspace v{0}.ps1" -f $version)
$taskNamePrefix = 'Check Disk Space'
$taskName = "{0} (v{1})" -f $taskNamePrefix, $version

Write-Verbose "Resolved version: $version"
Write-Verbose "Download URL: $downloadUri"
Write-Verbose "Target path: $targetPath"

if (-not (Test-Path -Path $TargetDirectory)) {
    Write-Verbose "Creating target directory $TargetDirectory"
    New-Item -Path $TargetDirectory -ItemType Directory -Force | Out-Null
}

Write-Verbose 'Removing prior script versions'
Get-ChildItem -Path $TargetDirectory -Filter 'check-diskspace v*.ps1' -File -ErrorAction SilentlyContinue |
    Where-Object { $_.FullName -ne $targetPath } |
    ForEach-Object {
        try {
            Remove-Item -Path $_.FullName -Force
            Write-Verbose "Removed $($_.FullName)"
        } catch {
            Write-Warning "Failed to remove $($_.FullName): $($_.Exception.Message)"
        }
    }

Write-Verbose 'Ensuring no existing scheduled tasks remain'
Get-ScheduledTask -TaskName "$taskNamePrefix*" -ErrorAction SilentlyContinue |
    ForEach-Object {
        try {
            Unregister-ScheduledTask -TaskName $_.TaskName -Confirm:$false
            Write-Verbose "Removed scheduled task '$($_.TaskName)'"
        } catch {
            Write-Warning "Failed to remove scheduled task '$($_.TaskName)': $($_.Exception.Message)"
        }
    }

Write-Host "Downloading check-diskspace.ps1 (version $version)"
Invoke-WebRequest -Uri $downloadUri -OutFile $targetPath -UseBasicParsing

$escapedTargetPath = $targetPath.Replace('"', '""')
$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$escapedTargetPath`""
$trigger = New-ScheduledTaskTrigger -Daily -At 6:00AM
$principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -Compatibility Win8 -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Write-Verbose "Registering scheduled task '$taskName'"
Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description "Runs check-diskspace.ps1 v$version at 6 AM daily" -Force | Out-Null

Write-Host "Scheduled task '$taskName' created to run $targetPath at 6 AM daily."
Write-Output @{ Version = $version; ScriptPath = $targetPath; TaskName = $taskName }
