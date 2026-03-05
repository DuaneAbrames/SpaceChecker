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
    [string]$ScheduleTime,
    [STRING]$configSourceUrl = "" #PROBABLY WON'T USE THIS
)
$verbose = $true
##########################################################################################
#  IMPORTANT: YOU CAN SET THIS TO YOUR OWN CONFIG URL OR PROVIDE A LOCAL CONFIG FILE     #
#  OR, MAKE SURE THE ENVIRONMENT VARIABLE 'CheckDiskSpaceConfig' IS SET TO A VALID URL   #
#  THE SCRIPT WILL ALSO TRY TO CONSTRUCT A URL BASED ON THE PRIMARY DNS DOMAIN           #
#  (e.g. https://space.yourdomain.com/check-diskspace-config.json)                       #
##########################################################################################
$configSourceUrl = ""

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$scriptDirectory = if ($PSCommandPath) { Split-Path -Path $PSCommandPath -Parent } else { (Get-Location).ProviderPath }
$configPath = Join-Path -Path $scriptDirectory -ChildPath $ConfigRelativePath

function Resolve-CnameTarget {
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    try {
        $result = Resolve-DnsName -Name $Name -Type CNAME -ErrorAction Stop
        return $result.NameHost
    }
    catch {
        # Not a CNAME or lookup failed — return the original name
        return $Name
    }
}
function Get-PrimaryDnsDomain {
    # Try common sources for the machine's primary DNS domain and return the first usable value.
    $candidates = @()

    if ($env:USERDNSDOMAIN) { $candidates += $env:USERDNSDOMAIN }

    try {
        $csDomain = (Get-WmiObject -Class Win32_ComputerSystem -ErrorAction Stop).Domain
        if ($csDomain -and $csDomain -ne $env:COMPUTERNAME) { $candidates += $csDomain }
    } catch {}

    try {
        $dnsSettings = Get-DnsClientGlobalSetting -ErrorAction Stop
        if ($dnsSettings.PrimaryDnsSuffix) { $candidates += $dnsSettings.PrimaryDnsSuffix }
        if ($dnsSettings.SuffixSearchList) { $candidates += $dnsSettings.SuffixSearchList }
    } catch {}

    return ($candidates | Where-Object { $_ -and $_ -notmatch '^WORKGROUP$' } | Select-Object -First 1).ToLowerInvariant()
}

function Get-ConfigUrlFromPrimaryDomain {
    param([string]$ConfigFileName = $ConfigRelativePath)

    $domain = Get-PrimaryDnsDomain
    if (-not $domain) { return $null }
    if (-not $ConfigFileName) { $ConfigFileName = 'check-diskspace-config.json' }

    $fileName = $ConfigFileName.TrimStart('/','\').TrimStart('/','\').TrimStart('/','\')
    $URLHost = "space.{0}" -f $domain.Trim()
    # If the constructed host is a CNAME, resolve it to get the actual target hostname for 
    # better compatibility with SSL certs and hosting providers.  (also useful for split-DNS setups)
    $URLHost = Resolve-CnameTarget -Name $URLHost
    return "https://$URLHost/$fileName"
}



# Environment variable can override config source URL, and if it looks like a URL, we'll try to download from it first 
if ($env:CheckDiskSpaceConfig) {
    Write-Host "Environment variable 'CheckDiskSpaceConfig' is set to '$($env:CheckDiskSpaceConfig)'"
    $configSourceUrl = $env:CheckDiskSpaceConfig
} else {
    Write-Host "Environment variable 'CheckDiskSpaceConfig' is not set."
}
if ($configSourceUrl -like "") {
    $configSourceUrl = Get-ConfigUrlFromPrimaryDomain -ConfigFileName $ConfigRelativePath
    if ($configSourceUrl) {
        Write-Host "Constructed config source URL based on primary DNS domain: $configSourceUrl"
    } else {
        Write-Host "Could not determine primary DNS domain or construct config source URL."
    }
}

# Persist the resolved config URL for future runs (machine scope).
if ($configSourceUrl) {
    try {
        [Environment]::SetEnvironmentVariable('CheckDiskSpaceConfig', $configSourceUrl, 'Machine')
        Write-Host "Persisted config source URL '$configSourceUrl' to machine environment variable 'CheckDiskSpaceConfig'"
    } catch {
        Write-Warning "Unable to persist CheckDiskSpaceConfig at machine scope: $($_.Exception.Message)"
    }
}

$forceConfigDownload = [bool]($configSourceUrl -and $configSourceUrl -match '^https?://')
if ($forceConfigDownload) {
    Write-Host "Forcing configuration download from $configSourceUrl on each run."
} else {
    Write-Host "No valid config source URL; will rely on existing config file at $configPath or throw if missing."
}



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
        Write-Host "Querying GitHub release metadata from $apiUri"
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
        Write-Host "Falling back to branch '$FallbackBranch' at $downloadUri"
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
    Write-Host "Creating target directory $TargetDirectory"
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

Write-Output ([pscustomobject]@{
    Version        = $metadata.Version
    ScriptPath     = $targetPath
    TaskName       = $taskName
    ScheduleTime   = $triggerTime.ToShortTimeString()
    ConfigPath     = $configPath
    TargetConfig   = Join-Path $TargetDirectory $ConfigRelativePath
})
