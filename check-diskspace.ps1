param(
    [ValidateRange(1,100)]
    [int]$PercentUsedThreshold = 90,

    [ValidateRange(0,[int]::MaxValue)]
    [int]$MinFreeGB = 50,

    [switch]$debug
)

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
$Subject = "Disk space alert on $ComputerName"

# Acquire logical disks and return current usage metrics.
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

        $hasTotal = $info.PSObject.Properties['TotalSize'] -and $info.TotalSize -ne $null -and $info.TotalSize -gt 0
        $hasFree = $info.PSObject.Properties['FreeSpace'] -and $info.FreeSpace -ne $null -and $info.FreeSpace -ge 0

        $sizeBytes = if ($hasTotal) { [double]$info.TotalSize } else { 0 }
        $freeBytes = if ($hasFree) { [double]$info.FreeSpace } else { 0 }

        if (-not $hasTotal -or -not $hasFree -or $sizeBytes -eq 0) {
            $volume = $null

            if ($info.VolumeName) {
                $volume = $allVolumes | Where-Object { $_.DeviceID -eq $info.VolumeName }
            }

            if (-not $volume -and $friendlyName) {
                $mountPath = "C:\\ClusterStorage\\$friendlyName\\"
                $volume = $allVolumes | Where-Object { $_.Name -eq $mountPath }
            }

            if (-not $volume -and $info.MountPoints) {
                foreach ($mountPoint in $info.MountPoints) {
                    $normalized = if ($mountPoint.EndsWith('\')) { $mountPoint } else { "$mountPoint\" }
                    $volume = $allVolumes | Where-Object { $_.Name -eq $normalized }
                    if ($volume) { break }
                }
            }

            if ($volume) {
                $sizeBytes = [double]$volume.Capacity
                $freeBytes = [double]$volume.FreeSpace
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
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

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
        "Disk space report for $ComputerName generated at $timestamp.",
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
