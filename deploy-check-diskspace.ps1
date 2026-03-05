#!ps
#timeout=999999
[CmdletBinding()]
param(
    [string]$GitHubRepo = 'DuaneAbrames/SpaceChecker',
    [string]$AssetRelativePath = 'check-diskspace.ps1',
    [string]$FallbackBranch = 'main',
    [string]$ConfigRelativePath = 'check-diskspace-config.json',
    [string]$TargetDirectory,
    [string]$TaskNamePrefix,
    [string]$ScheduleTime
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$scriptDirectory = if ($PSCommandPath) { Split-Path -Path $PSCommandPath -Parent } else { (Get-Location).ProviderPath }
$configPath = Join-Path -Path $scriptDirectory -ChildPath $ConfigRelativePath
$configSourceUrl = $env:CheckDiskSpaceConfig
$forceConfigDownload = [bool]($configSourceUrl -and $configSourceUrl -match '^https?://')

function Get-RemoteScriptMetadata {
    param(
        [Parameter(Mandatory)] [string]$Repo,
        [Parameter(Mandatory)] [string]$AssetPath,
        [Parameter(Mandatory)] [string]$FallbackBranch
    )

    $version = '0.0.0'
    $releaseTag = $null
    $downloadUri = $null
    $apiUri = "https://api.github.com/repos/$Repo/releases/latest"

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $headers = @{ 'User-Agent' = 'deploy-check-diskspace'; 'Accept' = 'application/vnd.github+json' }

    try {
        Write-Verbose "Querying GitHub release metadata from $apiUri"
        $latestRelease = Invoke-RestMethod -Uri $apiUri -Headers $headers -UseBasicParsing

        if ($latestRelease.tag_name) {
            $releaseTag = $latestRelease.tag_name
            $version = $latestRelease.tag_name.TrimStart('v')
        } elseif ($latestRelease.name) {
            $releaseTag = $latestRelease.name
            $version = $latestRelease.name.TrimStart('v')
        }

        if ($releaseTag) {
            $downloadUri = "https://raw.githubusercontent.com/$Repo/$releaseTag/$AssetPath"
        }
    } catch {
        Write-Warning "Unable to retrieve GitHub release information: $($_.Exception.Message)"
    }

    if (-not $downloadUri) {
        $releaseTag = $FallbackBranch
        if ($version -eq '0.0.0') {
            $version = Get-Date -Format 'yyyyMMddHHmmss'
        }
        $downloadUri = "https://raw.githubusercontent.com/$Repo/$FallbackBranch/$AssetPath"
        Write-Verbose "Falling back to branch '$FallbackBranch' at $downloadUri"
    }

    return [pscustomobject]@{
        Version     = $version
        ReleaseTag  = $releaseTag
        DownloadUri = $downloadUri
    }
}

function Ensure-ConfigFile {
    param(
        [Parameter(Mandatory)] [string]$ConfigPath,
        [pscustomobject]$Metadata,
        [Parameter(Mandatory)] [string]$Repo,
        [Parameter(Mandatory)] [string]$FallbackBranch,
        [Parameter(Mandatory)] [string]$RelativePath,
        [string]$ConfigSourceUrl,
        [switch]$ForceDownload
    )

    $downloaded = $false
    if ($ConfigSourceUrl -and $ConfigSourceUrl -match '^https?://') {
        try {
            Invoke-WebRequest -Uri $ConfigSourceUrl -OutFile $ConfigPath -UseBasicParsing -ErrorAction Stop
            Write-Host "Configuration downloaded from $ConfigSourceUrl"
            $downloaded = $true
        } catch {
            Write-Warning "Failed to download configuration from ${ConfigSourceUrl}: $($_.Exception.Message)"
        }
    }

    if ($downloaded) { return }

    if ($ForceDownload -or -not (Test-Path -Path $ConfigPath)) {
        $tagsToTry = @()
        if ($Metadata -and $Metadata.ReleaseTag) { $tagsToTry += $Metadata.ReleaseTag }
        if (-not $tagsToTry.Contains($FallbackBranch)) { $tagsToTry += $FallbackBranch }

        foreach ($tag in $tagsToTry) {
            $configUri = "https://raw.githubusercontent.com/$Repo/$tag/$RelativePath"
            try {
                Invoke-WebRequest -Uri $configUri -OutFile $ConfigPath -UseBasicParsing -ErrorAction Stop
                Write-Host "Configuration downloaded from $configUri"
                $downloaded = $true
                break
            } catch {
                Write-Warning "Failed to download configuration from ${configUri}: $($_.Exception.Message)"
            }
        }
    }

    if (-not (Test-Path -Path $ConfigPath)) {
        throw "Configuration file not found at $ConfigPath. Provide one or set the CheckDiskSpaceConfig environment variable."
    }
}

function Load-Configuration {
    param([Parameter(Mandatory)] [string]$ConfigPath)

    try {
        $raw = Get-Content -Path $ConfigPath -Raw -Encoding UTF8
        return $raw | ConvertFrom-Json
    } catch {
        throw "Failed to parse configuration at ${ConfigPath}: $($_.Exception.Message)"
    }
}

function Get-DailyTriggerTime {
    param([string]$TimeString)

    if ([string]::IsNullOrWhiteSpace($TimeString)) {
        return [DateTime]::Today.AddHours(6)
    }

    $formats = @('hh\:mm','h\:mm','HH\:mm')
    foreach ($fmt in $formats) {
        try {
            $timeSpan = [TimeSpan]::ParseExact($TimeString, $fmt, [System.Globalization.CultureInfo]::InvariantCulture)
            return [DateTime]::Today.Add($timeSpan)
        } catch {}
    }

    try {
        $timeSpan = [TimeSpan]::Parse($TimeString, [System.Globalization.CultureInfo]::InvariantCulture)
        return [DateTime]::Today.Add($timeSpan)
    } catch {}

    Write-Warning "Unable to parse schedule time '$TimeString'; defaulting to 06:00."
    return [DateTime]::Today.AddHours(6)
}

$metadata = Get-RemoteScriptMetadata -Repo $GitHubRepo -AssetPath $AssetRelativePath -FallbackBranch $FallbackBranch
Ensure-ConfigFile -ConfigPath $configPath -Metadata $metadata -Repo $GitHubRepo -FallbackBranch $FallbackBranch -RelativePath $ConfigRelativePath -ConfigSourceUrl $configSourceUrl -ForceDownload:$forceConfigDownload
$config = Load-Configuration -ConfigPath $configPath

if (-not $PSBoundParameters.ContainsKey('GitHubRepo') -and $config.GitHubRepo) { $GitHubRepo = $config.GitHubRepo }
if (-not $PSBoundParameters.ContainsKey('AssetRelativePath') -and $config.AssetRelativePath) { $AssetRelativePath = $config.AssetRelativePath }
if (-not $PSBoundParameters.ContainsKey('FallbackBranch') -and $config.FallbackBranch) { $FallbackBranch = $config.FallbackBranch }
if (-not $PSBoundParameters.ContainsKey('TargetDirectory') -and $config.TargetDirectory) { $TargetDirectory = $config.TargetDirectory }
if (-not $PSBoundParameters.ContainsKey('TaskNamePrefix') -and $config.TaskNamePrefix) { $TaskNamePrefix = $config.TaskNamePrefix }
if (-not $PSBoundParameters.ContainsKey('ScheduleTime') -and $config.ScheduleTime) { $ScheduleTime = $config.ScheduleTime }

if (-not $TargetDirectory) { $TargetDirectory = 'C:\istools' }
if (-not $TaskNamePrefix) { $TaskNamePrefix = 'Check Disk Space' }
if (-not $ScheduleTime) { $ScheduleTime = '06:00' }

$metadata = Get-RemoteScriptMetadata -Repo $GitHubRepo -AssetPath $AssetRelativePath -FallbackBranch $FallbackBranch
Ensure-ConfigFile -ConfigPath $configPath -Metadata $metadata -Repo $GitHubRepo -FallbackBranch $FallbackBranch -RelativePath $ConfigRelativePath -ConfigSourceUrl $configSourceUrl -ForceDownload:$forceConfigDownload

$triggerTime = Get-DailyTriggerTime -TimeString $ScheduleTime

if (-not (Test-Path -Path $TargetDirectory)) {
    Write-Verbose "Creating target directory $TargetDirectory"
    New-Item -Path $TargetDirectory -ItemType Directory -Force | Out-Null
}

Ensure-ConfigFile -ConfigPath (Join-Path $TargetDirectory $ConfigRelativePath) -Metadata $metadata -Repo $GitHubRepo -FallbackBranch $FallbackBranch -RelativePath $ConfigRelativePath -ConfigSourceUrl $configSourceUrl -ForceDownload:$forceConfigDownload

$targetPath = Join-Path -Path $TargetDirectory -ChildPath ("check-diskspace v{0}.ps1" -f $metadata.Version)
Write-Host "Downloading check-diskspace.ps1 (version $($metadata.Version))"
Invoke-WebRequest -Uri $metadata.DownloadUri -OutFile $targetPath -UseBasicParsing

Get-ChildItem -Path $TargetDirectory -Filter 'check-diskspace v*.ps1' -File -ErrorAction SilentlyContinue |
    Where-Object { $_.FullName -ne $targetPath } |
    ForEach-Object {
        try {
            Remove-Item -Path $_.FullName -Force
            Write-Host "Removed script: $($_.FullName)"
        } catch {
            Write-Warning "Failed to remove $($_.FullName): $($_.Exception.Message)"
        }
    }

Get-ScheduledTask -TaskName "$TaskNamePrefix*" -ErrorAction SilentlyContinue |
    ForEach-Object {
        try {
            Unregister-ScheduledTask -TaskName $_.TaskName -Confirm:$false
            Write-Host "Removed scheduled task: $($_.TaskName)"
        } catch {
            Write-Warning "Failed to remove scheduled task '$($_.TaskName)': $($_.Exception.Message)"
        }
    }

$escapedTargetPath = $targetPath.Replace('"', '""')
$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$escapedTargetPath`""
$trigger = New-ScheduledTaskTrigger -Daily -At $triggerTime
$principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -Compatibility Win8 -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

$taskName = "{0} (v{1})" -f $TaskNamePrefix, $metadata.Version
$taskDescription = "Runs check-diskspace.ps1 v{0} at {1} daily" -f $metadata.Version, $triggerTime.ToShortTimeString()
Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description $taskDescription -Force | Out-Null
Write-Host ("Scheduled task '{0}' created to run {1} at {2} daily." -f $taskName, $targetPath, $triggerTime.ToShortTimeString())

Write-Output [pscustomobject]@{
    Version        = $metadata.Version
    ScriptPath     = $targetPath
    TaskName       = $taskName
    ScheduleTime   = $triggerTime.ToShortTimeString()
    ConfigPath     = $configPath
    TargetConfig   = Join-Path $TargetDirectory $ConfigRelativePath
}
