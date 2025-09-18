param(
    [ValidateRange(1,100)]
    [int]$PercentUsedThreshold = 90,

    [ValidateRange(0,[int]::MaxValue)]
    [int]$MinFreeGB = 50,

    [switch]$debug,
    [switch]$DisableSelfUpdate,
    [switch]$SkipSelfUpdate,
    [string]$GitHubRepo = 'DuaneAbrames/SpaceChecker',
    [string]$AssetRelativePath = 'check-diskspace.ps1',
    [string]$FallbackBranch = 'main',
    [string]$TargetDirectory = $(if ($PSCommandPath) { Split-Path -Path $PSCommandPath -Parent } else { '.' }),
    [string]$TaskNamePrefix = 'Check Disk Space'
)

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

function Invoke-DeploymentUpdate {
    param(
        [Parameter(Mandatory)] [pscustomobject]$Metadata,
        [Parameter(Mandatory)] [string]$TargetDirectory,
        [Parameter(Mandatory)] [string]$TaskNamePrefix,
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
    $trigger = New-ScheduledTaskTrigger -Daily -At 6:00AM
    $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    $settings = New-ScheduledTaskSettingsSet -Compatibility Win8 -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

    $taskName = "{0} (v{1})" -f $TaskNamePrefix, $remoteVersion
    Write-Verbose "Registering scheduled task '$taskName'"
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description "Runs check-diskspace.ps1 v$remoteVersion at 6 AM daily" -Force | Out-Null
    Write-Host "Scheduled task '$taskName' created to run $targetPath at 6 AM daily."

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
        [Parameter(Mandatory)] [string]$ScriptPath,
        [Parameter()] [hashtable]$BoundParameters
    )

    $localVersion = Get-LocalScriptVersion -ScriptPath $ScriptPath
    $metadata = Get-RemoteScriptMetadata -Repo $Repo -AssetPath $AssetPath -FallbackBranch $FallbackBranch
    if (-not $metadata) { return }

    $remoteVersion = $metadata.Version
    $comparePerformed = $false
    $shouldUpdate = $false

    try {
        $currentVersionObj = $null
        $remoteVersionObj = $null
        if ([System.Version]::TryParse($localVersion, [ref]$currentVersionObj) -and
            [System.Version]::TryParse($remoteVersion, [ref]$remoteVersionObj)) {
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
        Invoke-DeploymentUpdate -Metadata $metadata -TargetDirectory $TargetDirectory -TaskNamePrefix $TaskNamePrefix -BoundParameters $BoundParameters
    } else {
        Write-Verbose "Self-update: local version $localVersion is up to date (remote $remoteVersion)."
    }
}

if (-not $DisableSelfUpdate -and -not $SkipSelfUpdate) {
    try {
        Invoke-SelfUpdate -Repo $GitHubRepo -AssetPath $AssetRelativePath -FallbackBranch $FallbackBranch -TargetDirectory $TargetDirectory -TaskNamePrefix $TaskNamePrefix -ScriptPath $PSCommandPath -BoundParameters $PSBoundParameters
    } catch [System.OperationCanceledException] {
        return
    }
}

$SmtpServer = 'nettek-com.mail.protection.outlook.com'
$From = 'noreply@nettek.com'

if ($debug) {
    $To = 'dabrames@nettek.com'
} else {
    $To = 'help@nettek.com'
}

$SmtpPort = 25
$UseSsl = $true
$ComputerName = $env:COMPUTERNAME
$timestampSubject = Get-Date -Format 'MM-dd-yy HH:mm'
$Subject = "Disk space alert on $ComputerName ($timestampSubject)"

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

    $sizeBytes = [double]$disk.Size
    $freeBytes = [double]$disk.FreeSpace
    $freeGBExact = $freeBytes / 1GB
    $usedPercentExact = if ($sizeBytes -eq 0) { 0 } else { (($sizeBytes - $freeBytes) / $sizeBytes) * 100 }

    $report += [pscustomobject]@{
        ComputerName   = $disk.SystemName
        Drive          = $disk.DeviceID
        SizeGB         = [Math]::Round($sizeBytes / 1GB, 2)
        FreeGB         = [Math]::Round($freeGBExact, 2)
        UsedPercent    = [Math]::Round($usedPercentExact, 1)
        AlertTriggered = ($usedPercentExact -ge $PercentUsedThreshold) -or ($freeGBExact -lt $MinFreeGB)
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
                    $normalized = if ($mountPoint.EndsWith('\')) { $mountPoint } else { "$mountPoint\" }
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

        $report += [pscustomobject]@{
            ComputerName   = $ComputerName
            Drive          = $displayName
            SizeGB         = [Math]::Round($sizeBytes / 1GB, 2)
            FreeGB         = [Math]::Round($freeGBExact, 2)
            UsedPercent    = [Math]::Round($usedPercentExact, 1)
            AlertTriggered = ($usedPercentExact -ge $PercentUsedThreshold) -or ($freeGBExact -lt $MinFreeGB)
            Source         = 'CSV'
        }
    }
}

if (-not $report) {
    Write-Warning "No disk data collected on $ComputerName."
    return
}

$sortedReport = $report | Sort-Object Source, Drive
$reportTable = $sortedReport | Format-Table Drive, SizeGB, FreeGB, UsedPercent, AlertTriggered, Source -AutoSize | Out-String
Write-Host $reportTable

$alertDisks = $sortedReport | Where-Object { $_.AlertTriggered }
Write-Host "debug: $debug"

if ($alertDisks -or $debug) {
    $thresholdDescription = "Thresholds: used >= $PercentUsedThreshold% or free < $MinFreeGB GB."
    if ($clusterDetected) {
        $thresholdDescription += ' Cluster Shared Volumes detected.'
    }
    $timestampBody = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    $driveBlocks = foreach ($item in $sortedReport) {
        $status = if ($item.AlertTriggered) { '***ALERT***' } else { 'OK' }
        $sizeDisplay = ('{0:N2} GB' -f $item.SizeGB)
        $freeDisplay = ('{0:N2} GB' -f $item.FreeGB)
        $usedDisplay = ('{0:N1}%' -f $item.UsedPercent)

        $lines = @(
            "Drive $($item.Drive):    $status",
            "Size:    $sizeDisplay",
            "Free:    $freeDisplay",
            "Used Percent:    $usedDisplay"
        )

        $lines -join [Environment]::NewLine
    }

    $bodySections = $driveBlocks -join ([Environment]::NewLine + [Environment]::NewLine)

    $bodyLines = @(
        "Disk space report for $ComputerName generated at $timestampBody.",
        $thresholdDescription,
        '',
        $bodySections
    )

    $body = ($bodyLines -join [Environment]::NewLine).TrimEnd()

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
            Write-Host "Alert email sent to $To."
        } catch {
            Write-Error "Failed to send alert email. $_"
        }
    } else {
        Write-Warning 'Alert thresholds exceeded but email parameters were not fully specified.'
    }
} else {
    Write-Host "All disks are within thresholds for $ComputerName."
}

return $sortedReport
