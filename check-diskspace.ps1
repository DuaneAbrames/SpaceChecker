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
$timestamp = Get-Date -Format 'MM-dd-yy HH:mm'
$Subject = "Disk space alert on $ComputerName ($timestamp)"

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
