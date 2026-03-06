param(
    [switch]$debug,
    [switch]$DisableSelfUpdate,
    [string]$GitHubRepo = 'DuaneAbrames/SpaceChecker',
    [string]$AssetRelativePath = 'check-diskspace.ps1',
    [string]$FallbackBranch = 'main',
    [string]$TargetDirectory,
    [string]$TaskNamePrefix,
    [string]$ScheduleTime,
    [string]$ConfigRelativePath = 'check-diskspace-config.json'
)
function Get-ScriptDirectory {
    param([string]$Path)

    if ($Path) { return Split-Path -Path $Path -Parent }
    return (Get-Location).ProviderPath
}

function Get-LocalScriptVersion {
    param([string]$ScriptPath)

    if (-not $ScriptPath) { return '0.0.0' }
    $name = [System.IO.Path]::GetFileNameWithoutExtension($ScriptPath)
    $match = [regex]::Match($name, 'check-diskspace v(?<ver>.+)$')
    if ($match.Success) { return $match.Groups['ver'].Value }
    return '0.0.0'
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
    $headers = @{ 'User-Agent' = 'check-diskspace-selfupdate'; 'Accept' = 'application/vnd.github+json' }

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
        [switch]$ForceDownload
    )
    $configSourceUrl = $env:CheckDiskSpaceConfig
    Write-Debug "Initial configuration source URL from environment variable: $configSourceUrl"
    if (-not $configSourceUrl) {
        Write-Debug "No configuration source URL provided via environment variable; attempting to derive from primary DNS domain."
        $configSourceUrl = Get-ConfigUrlFromPrimaryDomain -ConfigFileName $ConfigRelativePath
    }   
    $forceConfigDownload = [bool]($configSourceUrl -and $configSourceUrl -match '^https?://')
    if ($forceConfigDownload) {
        Write-Host "Configuration source URL provided via environment variable: $configSourceUrl"
    }
    $downloaded = $false
    if ($ConfigSourceUrl -and $ConfigSourceUrl -match '^https?://') {
        try {
            Invoke-WebRequest -Uri $ConfigSourceUrl -OutFile $ConfigPath -UseBasicParsing -ErrorAction Stop
            Write-Host "Configuration downloaded from $ConfigSourceUrl"
            $downloaded = $true
        } catch {
            Write-Warning "Failed to download configuration from $($ConfigSourceUrl): $($_.Exception.Message)"
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
                Write-Warning "Failed to download configuration from $($configUri): $($_.Exception.Message)"
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
        throw "Failed to parse configuration at $($ConfigPath): $($_.Exception.Message)"
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

function Normalize-StringArray {
    param($Value)

    if ($null -eq $Value) { return @() }
    if ($Value -is [string]) { return @($Value) }
    return [string[]]$Value
}

function Normalize-DriveLetter {
    param([string]$DriveLetter)

    if ([string]::IsNullOrWhiteSpace($DriveLetter)) { return $null }
    $normalized = $DriveLetter.Trim().ToUpper()
    if ($normalized -match '^[A-Z]$') { return "$($normalized):" }
    if ($normalized -match '^[A-Z]:$') { return $normalized }
    return $normalized
}

function Get-ThresholdValue {
    param(
        $Node,
        [string]$PropertyName,
        [double]$DefaultValue
    )

    if ($Node -and $Node.PSObject.Properties[$PropertyName] -and $Node.$PropertyName -ne $null) {
        return [double]$Node.$PropertyName
    }
    return [double]$DefaultValue
}

function Test-ThresholdCondition {
    param(
        [double]$UsedPercentExact,
        [double]$FreeGBExact,
        [double]$PercentUsed,
        [double]$MinFreeGB,
        [string]$Mode
    )

    $modeUpper = if ($Mode) { $Mode.Trim().ToUpper() } else { 'AND' }
    $percentHit = $UsedPercentExact -ge $PercentUsed
    $freeHit = $FreeGBExact -lt $MinFreeGB
    if ($modeUpper -eq 'OR') { return ($percentHit -or $freeHit) }
    return ($percentHit -and $freeHit)
}

function Get-SeverityRank {
    param([string]$Severity)

    switch ($Severity) {
        'OK' { return 0 }
        'WARN' { return 1 }
        'CRIT' { return 2 }
        'EMERG' { return 3 }
        default { return -1 }
    }
}

function Get-HighestSeverity {
    param([object[]]$Items)

    if (-not $Items) { return 'OK' }
    return ($Items | Sort-Object { Get-SeverityRank -Severity $_.Severity } -Descending | Select-Object -First 1).Severity
}

function Get-DiskSeverity {
    param(
        [string]$DriveLetter,
        [double]$SizeBytes,
        [double]$FreeBytes,
        [double]$UsedPercentExact,
        [bool]$IsOSVolume,
        [hashtable]$Config
    )

    $freeGBExact = if ($FreeBytes -ge 0) { $FreeBytes / 1GB } else { 0 }

    if ($IsOSVolume) {
        $os = $Config.OS
        $severity = 'OK'
        if ($freeGBExact -lt [double]$os.FreeGB.Warn) { $severity = 'WARN' }
        if ($freeGBExact -lt [double]$os.FreeGB.Crit) { $severity = 'CRIT' }
        if ($freeGBExact -lt [double]$os.FreeGB.Emerg) { $severity = 'EMERG' }

        if ($freeGBExact -lt [double]$os.FreeGB.Warn -and $os.PercentUsed) {
            $percentSeverity = 'OK'
            if ($UsedPercentExact -ge [double]$os.PercentUsed.Warn) { $percentSeverity = 'WARN' }
            if ($UsedPercentExact -ge [double]$os.PercentUsed.Crit) { $percentSeverity = 'CRIT' }
            if ($UsedPercentExact -ge [double]$os.PercentUsed.Emerg) { $percentSeverity = 'EMERG' }

            if ((Get-SeverityRank -Severity $percentSeverity) -gt (Get-SeverityRank -Severity $severity)) {
                $severity = $percentSeverity
            }
        }

        return $severity
    }

    $data = $Config.Data
    if (Test-ThresholdCondition -UsedPercentExact $UsedPercentExact -FreeGBExact $freeGBExact -PercentUsed ([double]$data.Emerg.PercentUsed) -MinFreeGB ([double]$data.Emerg.MinFreeGB) -Mode $data.Emerg.Mode) { return 'EMERG' }
    if (Test-ThresholdCondition -UsedPercentExact $UsedPercentExact -FreeGBExact $freeGBExact -PercentUsed ([double]$data.Crit.PercentUsed) -MinFreeGB ([double]$data.Crit.MinFreeGB) -Mode $data.Crit.Mode) { return 'CRIT' }
    if (Test-ThresholdCondition -UsedPercentExact $UsedPercentExact -FreeGBExact $freeGBExact -PercentUsed ([double]$data.Warn.PercentUsed) -MinFreeGB ([double]$data.Warn.MinFreeGB) -Mode $data.Warn.Mode) { return 'WARN' }
    return 'OK'
}

function Invoke-DeploymentUpdate {
    param(
        [Parameter(Mandatory)] [pscustomobject]$Metadata,
        [Parameter(Mandatory)] [string]$TargetDirectory,
        [Parameter(Mandatory)] [string]$TaskNamePrefix,
        [Parameter(Mandatory)] [DateTime]$TriggerTime,
        [Parameter(Mandatory)] [string]$Repo,
        [Parameter(Mandatory)] [string]$FallbackBranch,
        [Parameter(Mandatory)] [string]$ConfigRelativePath,
        [string]$ConfigSourceUrl,
        [Parameter()] [hashtable]$BoundParameters
    )

    if (-not (Test-Path -Path $TargetDirectory)) {
        Write-Verbose "Creating target directory $TargetDirectory"
        New-Item -Path $TargetDirectory -ItemType Directory -Force | Out-Null
    }

    $remoteVersion = $Metadata.Version
    $targetPath = Join-Path -Path $TargetDirectory -ChildPath ("check-diskspace v{0}.ps1" -f $remoteVersion)
    Write-Host "Downloading check-diskspace.ps1 (version $remoteVersion)"
    Invoke-WebRequest -Uri $Metadata.DownloadUri -OutFile $targetPath -UseBasicParsing

    Ensure-ConfigFile -ConfigPath (Join-Path $TargetDirectory $ConfigRelativePath) -Metadata $Metadata -Repo $Repo -FallbackBranch $FallbackBranch -RelativePath $ConfigRelativePath 

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
    $trigger = New-ScheduledTaskTrigger -Daily -At $TriggerTime
    $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    $settings = New-ScheduledTaskSettingsSet -Compatibility Win8 -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

    $taskName = "{0} (v{1})" -f $TaskNamePrefix, $remoteVersion
    $taskDescription = "Runs check-diskspace.ps1 v{0} at {1} daily" -f $remoteVersion, $TriggerTime.ToShortTimeString()
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description $taskDescription -Force | Out-Null
    Write-Host ("Scheduled task '{0}' created to run {1} at {2} daily." -f $taskName, $targetPath, $TriggerTime.ToShortTimeString())

    if (-not $BoundParameters) { $BoundParameters = @{} }
    $relaunchParams = @{}
    foreach ($kvp in $BoundParameters.GetEnumerator()) {
        if ($kvp.Key -ne 'SkipSelfUpdate') {
            $relaunchParams[$kvp.Key] = $kvp.Value
        }
    }
    $relaunchParams['SkipSelfUpdate'] = $true

    Write-Host 'Re-launching updated script.'
    & $targetPath @relaunchParams
    $global:LASTEXITCODE = $LASTEXITCODE
    throw [System.OperationCanceledException]::new('Self-update applied; reran updated script.')
}

function Invoke-SelfUpdate {
    param(
        [Parameter(Mandatory)] [string]$Repo,
        [Parameter(Mandatory)] [string]$AssetPath,
        [Parameter(Mandatory)] [string]$FallbackBranch,
        [Parameter(Mandatory)] [string]$TargetDirectory,
        [Parameter(Mandatory)] [string]$TaskNamePrefix,
        [Parameter(Mandatory)] [DateTime]$TriggerTime,
        [Parameter(Mandatory)] [string]$ScriptPath,
        [Parameter(Mandatory)] [string]$ConfigRelativePath,
        [string]$ConfigSourceUrl,
        [pscustomobject]$Metadata,
        [Parameter()] [hashtable]$BoundParameters
    )

    if (-not $ScriptPath) {
        Write-Verbose 'Self-update skipped because script path could not be determined.'
        return
    }

    if (-not $Metadata) {
        $Metadata = Get-RemoteScriptMetadata -Repo $Repo -AssetPath $AssetPath -FallbackBranch $FallbackBranch
    }

    $localVersion = Get-LocalScriptVersion -ScriptPath $ScriptPath
    $remoteVersion = $Metadata.Version
    $comparePerformed = $false
    $shouldUpdate = $false

    try {
        $currentVersionObj = $null
        $remoteVersionObj = $null
        if ([System.Version]::TryParse($localVersion, [ref]$currentVersionObj) -and [System.Version]::TryParse($remoteVersion, [ref]$remoteVersionObj)) {
            $comparePerformed = $true
            $shouldUpdate = ($remoteVersionObj -gt $currentVersionObj)
        }
    } catch {
        $comparePerformed = $false
    }

    if (-not $comparePerformed) {
        $shouldUpdate = -not ($localVersion -eq $remoteVersion)
    }

    if ($shouldUpdate) {
        Write-Host "Self-update: local version $localVersion, remote version $remoteVersion"
        Invoke-DeploymentUpdate -Metadata $Metadata -TargetDirectory $TargetDirectory -TaskNamePrefix $TaskNamePrefix -TriggerTime $TriggerTime -Repo $Repo -FallbackBranch $FallbackBranch -ConfigRelativePath $ConfigRelativePath -ConfigSourceUrl $ConfigSourceUrl -BoundParameters $BoundParameters
    } else {
        Write-Verbose "Self-update: local version $localVersion is up to date (remote $remoteVersion)."
    }
}

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


$scriptDirectory = Get-ScriptDirectory -Path $PSCommandPath
$configPath = Join-Path -Path $scriptDirectory -ChildPath $ConfigRelativePath
$configSourceUrl = $env:CheckDiskSpaceConfig
Write-Debug "Initial configuration source URL from environment variable: $configSourceUrl"
if (-not $configSourceUrl) {
    Write-Debug "No configuration source URL provided via environment variable; attempting to derive from primary DNS domain."
    $configSourceUrl = Get-ConfigUrlFromPrimaryDomain -ConfigFileName $ConfigRelativePath
}   
$forceConfigDownload = [bool]($configSourceUrl -and $configSourceUrl -match '^https?://')
if ($forceConfigDownload) {
    Write-Host "Configuration source URL provided via environment variable: $configSourceUrl"
}
$initialMetadata = Get-RemoteScriptMetadata -Repo $GitHubRepo -AssetPath $AssetRelativePath -FallbackBranch $FallbackBranch
Ensure-ConfigFile -ConfigPath $configPath -Metadata $initialMetadata -Repo $GitHubRepo -FallbackBranch $FallbackBranch -RelativePath $ConfigRelativePath 
$config = Load-Configuration -ConfigPath $configPath

if (-not $PSBoundParameters.ContainsKey('GitHubRepo') -and $config.GitHubRepo) { $GitHubRepo = $config.GitHubRepo }
if (-not $PSBoundParameters.ContainsKey('AssetRelativePath') -and $config.AssetRelativePath) { $AssetRelativePath = $config.AssetRelativePath }
if (-not $PSBoundParameters.ContainsKey('FallbackBranch') -and $config.FallbackBranch) { $FallbackBranch = $config.FallbackBranch }
if (-not $PSBoundParameters.ContainsKey('TaskNamePrefix') -and $config.TaskNamePrefix) { $TaskNamePrefix = $config.TaskNamePrefix }
if (-not $PSBoundParameters.ContainsKey('TargetDirectory') -and $config.TargetDirectory) { $TargetDirectory = $config.TargetDirectory }
if (-not $PSBoundParameters.ContainsKey('ScheduleTime') -and $config.ScheduleTime) { $ScheduleTime = $config.ScheduleTime }

if (-not $TaskNamePrefix) { $TaskNamePrefix = 'Check Disk Space' }
if (-not $TargetDirectory) { $TargetDirectory = $scriptDirectory }
if (-not $ScheduleTime) { $ScheduleTime = '06:00' }

$thresholdsNode = $config.Thresholds
$osNode = if ($thresholdsNode -and $thresholdsNode.OS) { $thresholdsNode.OS } else { $null }
$osFreeNode = if ($osNode -and $osNode.FreeGB) { $osNode.FreeGB } else { $null }
$osPercentNode = if ($osNode -and $osNode.PercentUsed) { $osNode.PercentUsed } else { $null }
$dataNode = if ($thresholdsNode -and $thresholdsNode.Data) { $thresholdsNode.Data } else { $null }
$dataWarnNode = if ($dataNode -and $dataNode.Warn) { $dataNode.Warn } else { $null }
$dataCritNode = if ($dataNode -and $dataNode.Crit) { $dataNode.Crit } else { $null }
$dataEmergNode = if ($dataNode -and $dataNode.Emerg) { $dataNode.Emerg } else { $null }

$osDriveLetters = Normalize-StringArray -Value $(if ($osNode -and $osNode.DriveLetters) { $osNode.DriveLetters } else { @('C:') })
if (-not $osDriveLetters -or $osDriveLetters.Count -eq 0) { $osDriveLetters = @('C:') }
$osDriveLetters = $osDriveLetters | ForEach-Object { Normalize-DriveLetter -DriveLetter $_ } | Where-Object { $_ } | Select-Object -Unique
if (-not $osDriveLetters -or $osDriveLetters.Count -eq 0) { $osDriveLetters = @('C:') }

$osPercentUsedConfig = $null
if ($osPercentNode -and $osPercentNode -ne [System.DBNull]::Value) {
    $osPercentUsedConfig = @{
        Warn  = Get-ThresholdValue -Node $osPercentNode -PropertyName 'Warn' -DefaultValue 90
        Crit  = Get-ThresholdValue -Node $osPercentNode -PropertyName 'Crit' -DefaultValue 95
        Emerg = Get-ThresholdValue -Node $osPercentNode -PropertyName 'Emerg' -DefaultValue 98
    }
}

$thresholdConfig = @{
    OS = @{
        DriveLetters = $osDriveLetters
        FreeGB = @{
            Warn  = Get-ThresholdValue -Node $osFreeNode -PropertyName 'Warn' -DefaultValue 20
            Crit  = Get-ThresholdValue -Node $osFreeNode -PropertyName 'Crit' -DefaultValue 10
            Emerg = Get-ThresholdValue -Node $osFreeNode -PropertyName 'Emerg' -DefaultValue 5
        }
        PercentUsed = $osPercentUsedConfig
    }
    Data = @{
        Warn = @{
            PercentUsed = Get-ThresholdValue -Node $dataWarnNode -PropertyName 'PercentUsed' -DefaultValue 90
            MinFreeGB   = Get-ThresholdValue -Node $dataWarnNode -PropertyName 'MinFreeGB' -DefaultValue 100
            Mode        = if ($dataWarnNode -and $dataWarnNode.Mode) { ([string]$dataWarnNode.Mode).Trim().ToUpper() } else { 'AND' }
        }
        Crit = @{
            PercentUsed = Get-ThresholdValue -Node $dataCritNode -PropertyName 'PercentUsed' -DefaultValue 95
            MinFreeGB   = Get-ThresholdValue -Node $dataCritNode -PropertyName 'MinFreeGB' -DefaultValue 50
            Mode        = if ($dataCritNode -and $dataCritNode.Mode) { ([string]$dataCritNode.Mode).Trim().ToUpper() } else { 'AND' }
        }
        Emerg = @{
            PercentUsed = Get-ThresholdValue -Node $dataEmergNode -PropertyName 'PercentUsed' -DefaultValue 98
            MinFreeGB   = Get-ThresholdValue -Node $dataEmergNode -PropertyName 'MinFreeGB' -DefaultValue 20
            Mode        = if ($dataEmergNode -and $dataEmergNode.Mode) { ([string]$dataEmergNode.Mode).Trim().ToUpper() } else { 'OR' }
        }
    }
}

$alertPolicyNode = $config.AlertPolicy
$warnDayOfWeek = if ($alertPolicyNode -and $alertPolicyNode.SendWarnOnDayOfWeek) { [string]$alertPolicyNode.SendWarnOnDayOfWeek } else { 'Sunday' }
$warnDayOfWeek = if ([string]::IsNullOrWhiteSpace($warnDayOfWeek)) { 'Sunday' } else { $warnDayOfWeek.Trim() }
$includeWarnInEmailBody = if ($alertPolicyNode -and $alertPolicyNode.PSObject.Properties['IncludeWarnInEmailBody']) { [bool]$alertPolicyNode.IncludeWarnInEmailBody } else { $true }
$isSunday = (Get-Date).DayOfWeek -eq 'Sunday'
$isWarnTriggerDay = (Get-Date).DayOfWeek.ToString().Equals($warnDayOfWeek, [System.StringComparison]::OrdinalIgnoreCase)
$warnGateDayMatch = if ($warnDayOfWeek -eq 'Sunday') { $isSunday } else { $isWarnTriggerDay }

$metadata = Get-RemoteScriptMetadata -Repo $GitHubRepo -AssetPath $AssetRelativePath -FallbackBranch $FallbackBranch
$triggerTime = Get-DailyTriggerTime -TimeString $ScheduleTime

if (-not $DisableSelfUpdate -and -not $DisableSelfUpdate -and -not $debug) {
    try {
        Invoke-SelfUpdate -Repo $GitHubRepo -AssetPath $AssetRelativePath -FallbackBranch $FallbackBranch -TargetDirectory $TargetDirectory -TaskNamePrefix $TaskNamePrefix -TriggerTime $triggerTime -ScriptPath $PSCommandPath -ConfigRelativePath $ConfigRelativePath -ConfigSourceUrl $configSourceUrl -Metadata $metadata -BoundParameters $PSBoundParameters
    } catch [System.OperationCanceledException] {
        return
    }
} elseif ($debug) {
    Write-Verbose 'Self-update skipped because debug mode is enabled.'
}

$SmtpServer = if ($config.SmtpServer) { [string]$config.SmtpServer } else { 'nettek-com.mail.protection.outlook.com' }
$SmtpPort = if ($config.SmtpPort) { [int]$config.SmtpPort } else { 25 }
$UseSsl = if ($config.UseSsl -ne $null) { [bool]$config.UseSsl } else { $true }
$From = if ($config.From) { [string]$config.From } else { 'noreply@nettek.com' }
$defaultTo = Normalize-StringArray -Value $config.To
$debugTo = Normalize-StringArray -Value $config.DebugTo

$primaryTo = if ($defaultTo.Count -gt 0) { $defaultTo } else { @('help@nettek.com') }
$debugRecipients = if ($debugTo.Count -gt 0) { $debugTo } elseif ($defaultTo.Count -gt 0) { $defaultTo } else { @('dabrames@nettek.com') }

$SubjectPrefix = if ($config.SubjectPrefix) { [string]$config.SubjectPrefix } else { $null }
$ComputerName = $env:COMPUTERNAME

try {
    $disks = Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType=3'
} catch {
    Write-Error "Failed to query disks on $ComputerName. $_"
    return
}

if (-not $disks) {
    Write-Warning "No fixed disks found on $ComputerName."
}

$report = @()

foreach ($disk in $disks) {
    if (-not $disk.Size) { continue }

    $driveLetter = Normalize-DriveLetter -DriveLetter $disk.DeviceID
    $isOSVolume = $thresholdConfig.OS.DriveLetters -contains $driveLetter
    $sizeBytes = [double]$disk.Size
    $freeBytes = [double]$disk.FreeSpace
    $freeGBExact = $freeBytes / 1GB
    $usedPercentExact = if ($sizeBytes -eq 0) { 0 } else { (($sizeBytes - $freeBytes) / $sizeBytes) * 100 }
    $severity = Get-DiskSeverity -DriveLetter $driveLetter -SizeBytes $sizeBytes -FreeBytes $freeBytes -UsedPercentExact $usedPercentExact -IsOSVolume $isOSVolume -Config $thresholdConfig
    $alertTriggered = ($severity -in @('CRIT', 'EMERG')) -or (($severity -eq 'WARN') -and $warnGateDayMatch)

    $report += [pscustomobject]@{
        ComputerName   = $disk.SystemName
        Drive          = $disk.DeviceID
        SizeGB         = [Math]::Round($sizeBytes / 1GB, 2)
        FreeGB         = [Math]::Round($freeGBExact, 2)
        UsedPercent    = [Math]::Round($usedPercentExact, 1)
        Severity       = $severity
        IsOSVolume     = $isOSVolume
        AlertTriggered = $alertTriggered
        Source         = 'LogicalDisk'
    }
}

$clusterDetected = $false
$csvVolumes = @()
$allVolumes = @()

if (Get-Command -Name Get-ClusterSharedVolume -ErrorAction SilentlyContinue) {
    try {
        $csvVolumes = Get-ClusterSharedVolume -ErrorAction Stop
        $allVolumes = Get-CimInstance -ClassName Win32_Volume -ErrorAction Stop
    } catch {
        Write-Warning "Detected Failover Clustering cmdlets but failed to enumerate Cluster Shared Volumes: $($_.Exception.Message)"
    }
}

if ($csvVolumes) {
    $clusterDetected = $true

    foreach ($csvVolume in $csvVolumes) {
        $info = $csvVolume.SharedVolumeInfo
        if (-not $info) { continue }

        $friendlyName = if ($info.FriendlyVolumeName) { $info.FriendlyVolumeName.Trim() } else { $csvVolume.Name }
        $displayName = if ($friendlyName) { "CSV: $friendlyName" } else { "CSV: $($csvVolume.Name)" }

        $sizeBytes = 0.0
        $freeBytes = 0.0

        if ($info.PSObject.Properties['TotalSize'] -and $info.TotalSize) {
            $sizeBytes = [double]$info.TotalSize
        }
        if ($info.PSObject.Properties['FreeSpace'] -and $info.FreeSpace -ne $null) {
            $freeBytes = [double]$info.FreeSpace
        }

        $partition = $info.Partition
        if ($partition) {
            if ($partition.PSObject.Properties['Size'] -and $partition.Size) {
                $sizeBytes = [double]$partition.Size
            }
            if ($partition.PSObject.Properties['FreeSpace'] -and $partition.FreeSpace -ne $null) {
                $freeBytes = [double]$partition.FreeSpace
            }
        }

        $needsVolumeLookup = ($sizeBytes -le 0)

        if ($needsVolumeLookup) {
            $volume = $null

            if ($info.PSObject.Properties['VolumeName'] -and $info.VolumeName) {
                $volume = $allVolumes | Where-Object { $_.DeviceID -eq $info.VolumeName }
            }

            if (-not $volume -and $partition -and $partition.PSObject.Properties['Name'] -and $partition.Name) {
                $volume = $allVolumes | Where-Object { $_.DeviceID -eq $partition.Name }
            }

            if (-not $volume -and $friendlyName) {
                $mountPath = "C:\\ClusterStorage\\$friendlyName"
                $alternates = @(
                    $mountPath,
                    "$mountPath\\"
                )
                foreach ($candidate in $alternates) {
                    $volume = $allVolumes | Where-Object { $_.Name -eq $candidate }
                    if ($volume) { break }
                }
            }

            if (-not $volume -and $info.PSObject.Properties['MountPoints'] -and $info.MountPoints) {
                foreach ($mountPoint in $info.MountPoints) {
                    $normalized = if ($mountPoint.EndsWith('\\')) { $mountPoint } else { "$mountPoint\\" }
                    $volume = $allVolumes | Where-Object { $_.Name -eq $normalized }
                    if ($volume) { break }
                }
            }

            if ($volume) {
                if ($volume.Capacity) { $sizeBytes = [double]$volume.Capacity }
                if ($volume.FreeSpace -ne $null) { $freeBytes = [double]$volume.FreeSpace }
            } else {
                Write-Warning "Unable to determine size for CSV '$displayName' using available metadata."
                continue
            }
        }

        $freeGBExact = $freeBytes / 1GB
        $usedPercentExact = if ($sizeBytes -eq 0) { 0 } else { (($sizeBytes - $freeBytes) / $sizeBytes) * 100 }
        $severity = Get-DiskSeverity -DriveLetter $displayName -SizeBytes $sizeBytes -FreeBytes $freeBytes -UsedPercentExact $usedPercentExact -IsOSVolume $false -Config $thresholdConfig
        $alertTriggered = ($severity -in @('CRIT', 'EMERG')) -or (($severity -eq 'WARN') -and $warnGateDayMatch)

        $report += [pscustomobject]@{
            ComputerName   = $ComputerName
            Drive          = $displayName
            SizeGB         = [Math]::Round($sizeBytes / 1GB, 2)
            FreeGB         = [Math]::Round($freeGBExact, 2)
            UsedPercent    = [Math]::Round($usedPercentExact, 1)
            Severity       = $severity
            IsOSVolume     = $false
            AlertTriggered = $alertTriggered
            Source         = 'CSV'
        }
    }
}

if (-not $report) {
    Write-Warning "No disk data collected on $ComputerName."
    return
}

$sortedReport = $report | Sort-Object Source, Drive
$reportTable = $sortedReport | Format-Table Drive, SizeGB, FreeGB, UsedPercent, Severity, IsOSVolume, AlertTriggered, Source -AutoSize | Out-String
Write-Host $reportTable

$triggeredAlerts = $sortedReport | Where-Object { $_.AlertTriggered }
$shouldSendPrimary = ($triggeredAlerts.Count -gt 0)
$shouldSendDebug = [bool]$debug

if ($shouldSendPrimary -or $shouldSendDebug) {
    $timestampBody = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $timestampSubject = Get-Date -Format 'MM-dd-yy HH:mm'
    $warnGateDescription = if ($warnDayOfWeek) { $warnDayOfWeek } else { 'Sunday' }
    $osPolicy = "OS policy (drives: {0}) FreeGB tiers: WARN < {1}, CRIT < {2}, EMERG < {3}. OS PercentUsed: {4}." -f ($thresholdConfig.OS.DriveLetters -join ', '), $thresholdConfig.OS.FreeGB.Warn, $thresholdConfig.OS.FreeGB.Crit, $thresholdConfig.OS.FreeGB.Emerg, $(if ($thresholdConfig.OS.PercentUsed) { "enabled (Warn>=$($thresholdConfig.OS.PercentUsed.Warn), Crit>=$($thresholdConfig.OS.PercentUsed.Crit), Emerg>=$($thresholdConfig.OS.PercentUsed.Emerg)) and applied only when FreeGB is already below WARN" } else { 'not set' })
    $dataPolicy = "Data policy: WARN (UsedPercent >= {0} AND FreeGB < {1}), CRIT (UsedPercent >= {2} AND FreeGB < {3}), EMERG (UsedPercent >= {4} {5} FreeGB < {6})." -f $thresholdConfig.Data.Warn.PercentUsed, $thresholdConfig.Data.Warn.MinFreeGB, $thresholdConfig.Data.Crit.PercentUsed, $thresholdConfig.Data.Crit.MinFreeGB, $thresholdConfig.Data.Emerg.PercentUsed, $thresholdConfig.Data.Emerg.Mode.ToUpper(), $thresholdConfig.Data.Emerg.MinFreeGB
    $gatingPolicy = "Alert policy: CRIT/EMERG always trigger; WARN triggers on $warnGateDescription."

    $emailReportItems = if ($includeWarnInEmailBody) { $sortedReport } else { $sortedReport | Where-Object { $_.Severity -ne 'WARN' } }
    if (-not $emailReportItems -or $emailReportItems.Count -eq 0) { $emailReportItems = $sortedReport }

    $driveBlocks = foreach ($item in $emailReportItems) {
        $status = switch ($item.Severity) {
            'EMERG' { '***EMERGENCY***' }
            'CRIT' { '***ALERT***' }
            'WARN' { 'WARN' }
            default { 'OK' }
        }
        $sizeDisplay = ('{0:N2} GB' -f $item.SizeGB)
        $freeDisplay = ('{0:N2} GB' -f $item.FreeGB)
        $usedDisplay = ('{0:N1}%' -f $item.UsedPercent)
        $volumePolicy = if ($item.IsOSVolume) { 'OS' } else { 'Data' }

        $lines = @(
            "Drive $($item.Drive):    $status",
            "Severity:    $($item.Severity)",
            "Policy:    $volumePolicy",
            "Triggered:    $($item.AlertTriggered)",
            "Size:    $sizeDisplay",
            "Free:    $freeDisplay",
            "Used Percent:    $usedDisplay"
        )
        $lines -join [Environment]::NewLine
    }

    $bodySections = $driveBlocks -join ([Environment]::NewLine + [Environment]::NewLine)
    $bodyLines = @(
        "Disk space report for $ComputerName generated at $timestampBody.",
        "Script version: $($metadata.Version) from release '$($metadata.ReleaseTag)'.",
        $gatingPolicy,
        $osPolicy,
        $dataPolicy,
        $(if ($clusterDetected) { 'Cluster Shared Volumes detected.' } else { $null }),
        '',
        $bodySections
    ) | Where-Object { $_ -ne $null }
    $body = ($bodyLines -join [Environment]::NewLine).TrimEnd()

    $highestTriggeredSeverity = Get-HighestSeverity -Items $triggeredAlerts
    $highestSeenSeverity = Get-HighestSeverity -Items $sortedReport
    $subjectSeverity = if ($shouldSendPrimary) { $highestTriggeredSeverity } else { $highestSeenSeverity }
    $subjectTags = @("[{0}]" -f $subjectSeverity)
    if (-not $shouldSendPrimary -and $shouldSendDebug) { $subjectTags += '[DEBUG]' }
    $subjectCore = "Disk space alert on $ComputerName ($timestampSubject)"
    $subjectWithTags = "{0} {1}" -f ($subjectTags -join ''), $subjectCore
    $Subject = if ($SubjectPrefix) { "{0} {1}" -f $SubjectPrefix.Trim(), $subjectWithTags } else { $subjectWithTags }

    $recipientSets = @()
    if ($shouldSendPrimary) { $recipientSets += ,@{ Name = 'primary'; Recipients = $primaryTo } }
    if ($shouldSendDebug) { $recipientSets += ,@{ Name = 'debug'; Recipients = $debugRecipients } }

    foreach ($recipientSet in $recipientSets) {
        $To = $recipientSet.Recipients
        if ($SmtpServer -and $To -and $From) {
            try {
                $mailParams = @{
                    SmtpServer = $SmtpServer
                    Port       = $SmtpPort
                    To         = $To
                    From       = $From
                    Subject    = $Subject
                    Body       = $body
                }

                if ($UseSsl) { $mailParams.UseSsl = $true }
                if ($Credential) { $mailParams.Credential = $Credential }

                Send-MailMessage @mailParams
                Write-Host "Alert email sent to $($To -join ', ') [$($recipientSet.Name)]."
            } catch {
                Write-Error "Failed to send alert email to $($recipientSet.Name) recipients. $_"
            }
        } else {
            Write-Warning 'Alert conditions met but email parameters were not fully specified.'
        }
    }
} else {
    Write-Host "All disks are within thresholds for $ComputerName."
}

return $sortedReport
