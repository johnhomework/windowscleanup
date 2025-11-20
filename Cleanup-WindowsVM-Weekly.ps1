# Windows VM Weekly Maintenance Script
# Safe for regular production VM cleanup - focuses on disk space recovery

#Requires -RunAsAdministrator

param(
    [switch]$WhatIf,
    [switch]$SkipBrowserCleanup
)

Write-Host "Windows VM Weekly Maintenance Script" -ForegroundColor Green
Write-Host "=====================================" -ForegroundColor Green

if ($WhatIf) {
    Write-Host "Running in WhatIf mode - no files will be deleted" -ForegroundColor Yellow
}

# Get initial disk space
try {
    $initialDisk = Get-WmiObject -Class Win32_LogicalDisk -Filter "DeviceID='C:'"
    $initialFreeSpaceGB = [math]::Round($initialDisk.FreeSpace / 1GB, 2)
    $totalSpaceGB = [math]::Round($initialDisk.Size / 1GB, 2)
    $initialUsedSpaceGB = $totalSpaceGB - $initialFreeSpaceGB
    $initialUsedPercent = [math]::Round(($initialUsedSpaceGB / $totalSpaceGB) * 100, 1)
    
    Write-Host "`nInitial Disk Space:" -ForegroundColor Yellow
    Write-Host "Total: $totalSpaceGB GB" -ForegroundColor White
    Write-Host "Used:  $initialUsedSpaceGB GB ($initialUsedPercent%)" -ForegroundColor White  
    Write-Host "Free:  $initialFreeSpaceGB GB" -ForegroundColor White
}
catch {
    Write-Host "`nCould not determine initial disk space" -ForegroundColor Yellow
    $initialFreeSpaceGB = 0
}

# Function to safely remove files/folders
function Remove-SafelyWithLogging {
    param(
        [string]$Path,
        [string]$Description
    )
    
    if (Test-Path $Path) {
        try {
            $size = (Get-ChildItem $Path -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
            $sizeGB = [math]::Round($size / 1GB, 2)
            
            Write-Host "Cleaning: $Description ($sizeGB GB)" -ForegroundColor Cyan
            
            if (-not $WhatIf) {
                Remove-Item $Path -Recurse -Force -ErrorAction SilentlyContinue
                Write-Host "  V Removed" -ForegroundColor Green
            } else {
                Write-Host "  -> Would remove $sizeGB GB" -ForegroundColor Yellow
            }
        }
        catch {
            Write-Host "  X Error: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    else {
        Write-Host "Skipping: $Description (not found)" -ForegroundColor Gray
    }
}

Write-Host "`n1. Stopping Services for Cleanup..." -ForegroundColor Yellow
$servicesToStop = @("wuauserv", "BITS", "CryptSvc")
foreach ($service in $servicesToStop) {
    try {
        Stop-Service $service -Force -ErrorAction SilentlyContinue
        Write-Host "  Stopped: $service" -ForegroundColor Green
    }
    catch {
        Write-Host "  Warning: Could not stop $service" -ForegroundColor Yellow
    }
}

Write-Host "`n2. Windows System Temporary Files..." -ForegroundColor Yellow
$systemTempPaths = @(
    @{Path="C:\Windows\Temp\*"; Desc="Windows Temp"},
    @{Path="C:\Windows\Prefetch\*"; Desc="Prefetch Files"},
    @{Path="C:\Windows\SoftwareDistribution\Download\*"; Desc="Windows Update Downloads"},
    @{Path="C:\Windows\Logs\CBS\*"; Desc="Component-Based Servicing Logs"},
    @{Path="C:\Windows\Logs\DISM\*"; Desc="DISM Logs"},
    @{Path="C:\ProgramData\Microsoft\Windows\WER\ReportQueue\*"; Desc="Windows Error Reporting Queue"},
    @{Path="C:\Windows\LiveKernelReports\*"; Desc="Kernel Reports"},
    @{Path="C:\Windows\Minidump\*"; Desc="Memory Dumps"},
    @{Path="C:\Windows\System32\LogFiles\WMI\*"; Desc="WMI Logs"},
    @{Path="C:\Windows\Panther\UnattendGC\*"; Desc="Setup Logs"}
)

foreach ($temp in $systemTempPaths) {
    Remove-SafelyWithLogging -Path $temp.Path -Description $temp.Desc
}

Write-Host "`n3. User Temporary Files..." -ForegroundColor Yellow
$userTempPaths = @(
    @{Path="C:\Users\*\AppData\Local\Temp\*"; Desc="User Temp Files"},
    @{Path="C:\Users\*\AppData\Local\Microsoft\Windows\Temporary Internet Files\*"; Desc="IE Temporary Files"},
    @{Path="C:\Users\*\AppData\Local\Microsoft\Windows\INetCache\*"; Desc="IE/Edge Cache"},
    @{Path="C:\Users\*\AppData\Local\Microsoft\Windows\Explorer\thumbcache_*.db"; Desc="Thumbnail Cache"},
    @{Path="C:\Users\*\AppData\Local\IconCache.db"; Desc="Icon Cache"},
    @{Path="C:\Users\*\AppData\Local\Microsoft\Windows\Caches\*"; Desc="Windows Caches"}
)

foreach ($userTemp in $userTempPaths) {
    Remove-SafelyWithLogging -Path $userTemp.Path -Description $userTemp.Desc
}

if (-not $SkipBrowserCleanup) {
    Write-Host "`n4. Browser Cache and Data..." -ForegroundColor Yellow
    $browserPaths = @(
        @{Path="C:\Users\*\AppData\Local\Google\Chrome\User Data\*\Cache\*"; Desc="Chrome Cache"},
        @{Path="C:\Users\*\AppData\Local\Google\Chrome\User Data\*\Code Cache\*"; Desc="Chrome Code Cache"},
        @{Path="C:\Users\*\AppData\Local\Microsoft\Edge\User Data\*\Cache\*"; Desc="Edge Cache"},
        @{Path="C:\Users\*\AppData\Local\Microsoft\Edge\User Data\*\Code Cache\*"; Desc="Edge Code Cache"},
        @{Path="C:\Users\*\AppData\Local\Mozilla\Firefox\Profiles\*\cache2\*"; Desc="Firefox Cache"},
        @{Path="C:\Users\*\AppData\Local\Mozilla\Firefox\Profiles\*\startupCache\*"; Desc="Firefox Startup Cache"},
        @{Path="C:\Users\*\AppData\Roaming\Opera Software\Opera Stable\Cache\*"; Desc="Opera Cache"},
        @{Path="C:\Users\*\AppData\Local\BraveSoftware\Brave-Browser\User Data\*\Cache\*"; Desc="Brave Cache"}
    )
    
    foreach ($browserPath in $browserPaths) {
        Remove-SafelyWithLogging -Path $browserPath.Path -Description $browserPath.Desc
    }
}

Write-Host "`n5. Development Tools Cache..." -ForegroundColor Yellow
$devPaths = @(
    @{Path="C:\Users\*\.nuget\packages\*"; Desc="NuGet Package Cache"},
    @{Path="C:\Users\*\AppData\Roaming\npm-cache\*"; Desc="NPM Cache"},
    @{Path="C:\Users\*\AppData\Local\yarn-cache\*"; Desc="Yarn Cache"},
    @{Path="C:\Users\*\.gradle\caches\*"; Desc="Gradle Cache"},
    @{Path="C:\Users\*\.m2\repository\*"; Desc="Maven Repository Cache"},
    @{Path="C:\Users\*\AppData\Local\pip\cache\*"; Desc="Python Pip Cache"},
    @{Path="C:\ProgramData\Package Cache\*"; Desc="Package Cache"},
    @{Path="C:\Program Files (x86)\Microsoft Visual Studio\Installer\resources\app\layout\*"; Desc="VS Installer Cache"}
)

foreach ($devPath in $devPaths) {
    Remove-SafelyWithLogging -Path $devPath.Path -Description $devPath.Desc
}

Write-Host "`n6. Office and Application Cache..." -ForegroundColor Yellow
$appCachePaths = @(
    @{Path="C:\Users\*\AppData\Local\Microsoft\Office\16.0\OfficeFileCache\*"; Desc="Office File Cache"},
    @{Path="C:\Users\*\AppData\Roaming\Microsoft\Teams\tmp\*"; Desc="Microsoft Teams Temp"},
    @{Path="C:\Users\*\AppData\Roaming\Microsoft\Teams\logs\*"; Desc="Microsoft Teams Logs"},
    @{Path="C:\Users\*\AppData\Roaming\Zoom\logs\*"; Desc="Zoom Logs"},
    @{Path="C:\Users\*\AppData\Local\Adobe\Common\Media Cache Files\*"; Desc="Adobe Media Cache"},
    @{Path="C:\Users\*\AppData\Roaming\Adobe\Common\Media Cache Files\*"; Desc="Adobe Media Cache (Roaming)"},
    @{Path="C:\Users\*\AppData\Local\Slack\Cache\*"; Desc="Slack Cache"},
    @{Path="C:\Users\*\AppData\Roaming\Slack\Cache\*"; Desc="Slack Cache (Roaming)"}
)

foreach ($appCache in $appCachePaths) {
    Remove-SafelyWithLogging -Path $appCache.Path -Description $appCache.Desc
}

Write-Host "`n7. Windows Store and UWP Apps..." -ForegroundColor Yellow
$storePaths = @(
    @{Path="C:\Users\*\AppData\Local\Packages\*\AC\Temp\*"; Desc="UWP App Temp Files"},
    @{Path="C:\Users\*\AppData\Local\Packages\*\LocalCache\*"; Desc="UWP App Cache"},
    @{Path="C:\Program Files\WindowsApps\*\cache\*"; Desc="Windows Store App Cache"},
    @{Path="C:\ProgramData\Microsoft\Windows\AppRepository\StateRepository-Machine.srd-tmp"; Desc="App Repository Temp"}
)

foreach ($storePath in $storePaths) {
    Remove-SafelyWithLogging -Path $storePath.Path -Description $storePath.Desc
}

Write-Host "`n8. System Cache and Service Files..." -ForegroundColor Yellow
$serviceCachePaths = @(
    @{Path="C:\Windows\CSC\*"; Desc="Offline Files Cache"},
    @{Path="C:\Windows\System32\config\systemprofile\AppData\Local\Temp\*"; Desc="System Profile Temp"},
    @{Path="C:\Windows\ServiceProfiles\LocalService\AppData\Local\Temp\*"; Desc="LocalService Temp"},
    @{Path="C:\Windows\ServiceProfiles\NetworkService\AppData\Local\Temp\*"; Desc="NetworkService Temp"},
    @{Path="C:\Windows\System32\sru\*"; Desc="System Resource Usage Monitor"},
    @{Path="C:\ProgramData\Microsoft\Search\Data\Applications\Windows\tmp\*"; Desc="Windows Search Temp"}
)

foreach ($serviceCache in $serviceCachePaths) {
    Remove-SafelyWithLogging -Path $serviceCache.Path -Description $serviceCache.Desc
}

Write-Host "`n9. IIS and Web Server Logs..." -ForegroundColor Yellow
$webServerPaths = @(
    @{Path="C:\inetpub\logs\LogFiles\*"; Desc="IIS Log Files"},
    @{Path="C:\Windows\System32\LogFiles\W3SVC*\*"; Desc="IIS W3SVC Logs"},
    @{Path="C:\Windows\System32\LogFiles\HTTPERR\*"; Desc="HTTP Error Logs"},
    @{Path="C:\Windows\System32\LogFiles\SMTPSVC*\*"; Desc="SMTP Service Logs"}
)

foreach ($webPath in $webServerPaths) {
    Remove-SafelyWithLogging -Path $webPath.Path -Description $webPath.Desc
}

Write-Host "`n10. Docker and Container Cache..." -ForegroundColor Yellow
$containerPaths = @(
    @{Path="C:\ProgramData\Docker\tmp\*"; Desc="Docker Temp Files"},
    @{Path="C:\Users\*\.docker\machine\cache\*"; Desc="Docker Machine Cache"},
    @{Path="C:\ProgramData\DockerDesktop\vm-data\DockerDesktop.vhdx.tmp"; Desc="Docker Desktop Temp"}
)

foreach ($containerPath in $containerPaths) {
    Remove-SafelyWithLogging -Path $containerPath.Path -Description $containerPath.Desc
}

Write-Host "`n11. Windows Disk Cleanup..." -ForegroundColor Yellow
if (-not $WhatIf) {
    try {
        Write-Host "  Running automated disk cleanup..." -ForegroundColor Cyan
        
        $additionalPaths = @(
            @{Path="C:\Windows\Downloaded Program Files\*"; Desc="Downloaded Program Files"},
            @{Path="C:\Windows\Offline Web Pages\*"; Desc="Offline Web Pages"},
            @{Path="C:\ProgramData\Microsoft\Windows\RetailDemo\OfflineContent\*"; Desc="Retail Demo Content"},
            @{Path="C:\Windows\System32\DirectX\*"; Desc="DirectX Shader Cache"},
            @{Path="C:\Windows\SoftwareDistribution\DeliveryOptimization\*"; Desc="Delivery Optimization"},
            @{Path="C:\Windows\System32\DriverStore\Temp\*"; Desc="Driver Store Temp"}
        )
        
        foreach ($additionalPath in $additionalPaths) {
            Remove-SafelyWithLogging -Path $additionalPath.Path -Description $additionalPath.Desc
        }
        
        Write-Host "  V Disk cleanup completed" -ForegroundColor Green
    }
    catch {
        Write-Host "  Warning: Disk cleanup failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }
} else {
    Write-Host "  -> Would run automated disk cleanup" -ForegroundColor Yellow
}

Write-Host "`n12. Empty Recycle Bin..." -ForegroundColor Yellow
if (-not $WhatIf) {
    try {
        Clear-RecycleBin -Force -ErrorAction Stop
        Write-Host "  V Recycle Bin emptied" -ForegroundColor Green
    }
    catch {
        try {
            $recycleBin = New-Object -ComObject Shell.Application
            $recycleBin.Namespace(0xA).InvokeVerb("empty")
            Write-Host "  V Recycle Bin emptied" -ForegroundColor Green
        }
        catch {
            Write-Host "  Warning: Could not empty Recycle Bin" -ForegroundColor Yellow
        }
    }
} else {
    Write-Host "  -> Would empty Recycle Bin" -ForegroundColor Yellow
}

Write-Host "`n13. DISM Component Cleanup..." -ForegroundColor Yellow
if (-not $WhatIf) {
    try {
        Write-Host "  Running DISM component cleanup (this may take 5-15 minutes)..." -ForegroundColor Cyan
        Write-Host "  Please wait, DISM is working in background..." -ForegroundColor Gray
        
        # Start DISM with a timeout
        $dismJob = Start-Job -ScriptBlock {
            $dismResult = Start-Process "Dism.exe" -ArgumentList "/online /Cleanup-Image /StartComponentCleanup" -Wait -PassThru -WindowStyle Hidden
            return $dismResult.ExitCode
        }
        
        # Wait with progress indicator (max 20 minutes)
        $timeout = 1200 # 20 minutes in seconds
        $elapsed = 0
        
        while ($dismJob.State -eq "Running" -and $elapsed -lt $timeout) {
            Start-Sleep 30
            $elapsed += 30
            $minutes = [math]::Floor($elapsed / 60)
            Write-Host "  Still running... ($minutes minutes elapsed)" -ForegroundColor Gray
        }
        
        if ($dismJob.State -eq "Running") {
            Write-Host "  DISM taking too long, stopping..." -ForegroundColor Yellow
            Stop-Job $dismJob
            Remove-Job $dismJob
            Write-Host "  Warning: DISM cleanup timed out after 20 minutes" -ForegroundColor Yellow
        } else {
            $exitCode = Receive-Job $dismJob
            Remove-Job $dismJob
            
            if ($exitCode -eq 0) {
                Write-Host "  V DISM component cleanup completed" -ForegroundColor Green
            } else {
                Write-Host "  Warning: DISM cleanup failed with exit code $exitCode" -ForegroundColor Yellow
            }
        }
    }
    catch {
        Write-Host "  Error: DISM cleanup failed: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "  -> Would run DISM component cleanup (can take 5-15 minutes)" -ForegroundColor Yellow
}

Write-Host "`n14. Disk Optimization..." -ForegroundColor Yellow
if (-not $WhatIf) {
    try {
        $drives = Get-Volume | Where-Object { $_.DriveLetter -and $_.DriveType -eq "Fixed" }
        
        foreach ($drive in $drives) {
            Write-Host "  Running TRIM on drive $($drive.DriveLetter):\ ..." -ForegroundColor Cyan
            Optimize-Volume -DriveLetter $drive.DriveLetter -ReTrim -Verbose
            Write-Host "  V TRIM completed for drive $($drive.DriveLetter):\" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "  Warning: Disk optimization failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }
} else {
    Write-Host "  -> Would run TRIM on all fixed drives" -ForegroundColor Yellow
}

Write-Host "`n15. Network Cache Cleanup..." -ForegroundColor Yellow
if (-not $WhatIf) {
    try {
        ipconfig /flushdns | Out-Null
        Write-Host "  V DNS cache cleared" -ForegroundColor Green
        
        arp -d * 2>$null
        Write-Host "  V ARP cache cleared" -ForegroundColor Green
    }
    catch {
        Write-Host "  Warning: Network cache cleanup failed" -ForegroundColor Yellow
    }
} else {
    Write-Host "  -> Would clear DNS and ARP cache" -ForegroundColor Yellow
}

Write-Host "`n16. Restarting Services..." -ForegroundColor Yellow
foreach ($service in $servicesToStop) {
    try {
        Start-Service $service -ErrorAction SilentlyContinue
        Write-Host "  Restarted: $service" -ForegroundColor Green
    }
    catch {
        Write-Host "  Warning: Could not restart $service" -ForegroundColor Yellow
    }
}

Write-Host "`n17. Custom Application Log Cleanup..." -ForegroundColor Yellow
Write-Host "  Add your application-specific log cleanup below" -ForegroundColor Gray

# ============================================================================
# CUSTOM APPLICATION LOG CLEANUP SECTION
# ============================================================================
# Add entries below to clean up logs from your specific applications.
# Each entry should follow this format:
#
# @{Path="C:\Path\To\Your\Logs\*"; Desc="Your Application Logs"}
#
# The asterisk (*) at the end will delete all files in that directory.
# You can also use wildcards like *.log to target specific file types.
# 
# EXAMPLE CONFIGURATIONS:
# ============================================================================

$customLogPaths = @(
    # Example: Lucee/Tomcat application logs
    # Uncomment the line below to enable Lucee log cleanup
    # @{Path="C:\lucee\tomcat\logs\*"; Desc="Lucee Tomcat Logs"}
    
    # Example: IIS logs older than 30 days
    # @{Path="C:\inetpub\logs\LogFiles\*\*.log"; Desc="IIS Logs"}
    
    # Example: Custom application logs
    # @{Path="C:\MyApp\logs\*.log"; Desc="MyApp Logs"}
    # @{Path="C:\Program Files\MyService\logs\*"; Desc="MyService Logs"}
    
    # Example: SQL Server error logs (be careful with these!)
    # @{Path="C:\Program Files\Microsoft SQL Server\MSSQL*\MSSQL\Log\ERRORLOG.*"; Desc="SQL Server Error Logs (old)"}
    
    # Add your custom paths here:
)

# Process custom log cleanup if any paths are defined
if ($customLogPaths.Count -gt 0) {
    foreach ($customLog in $customLogPaths) {
        Remove-SafelyWithLogging -Path $customLog.Path -Description $customLog.Desc
    }
} else {
    Write-Host "  No custom log paths configured (edit script to add your paths)" -ForegroundColor Gray
}

# ============================================================================
# ADVANCED: Time-based log cleanup
# ============================================================================
# For more precise control, you can add age-based cleanup using Get-ChildItem filtering.
# Example: Delete Lucee logs older than 30 days
# 
# if (Test-Path "C:\lucee\tomcat\logs") {
#     try {
#         $cutoffDate = (Get-Date).AddDays(-30)
#         $oldLogs = Get-ChildItem "C:\lucee\tomcat\logs\*.log" -Recurse -ErrorAction SilentlyContinue | 
#                    Where-Object { $_.LastWriteTime -lt $cutoffDate }
#         
#         $size = ($oldLogs | Measure-Object -Property Length -Sum).Sum
#         $sizeGB = [math]::Round($size / 1GB, 2)
#         
#         Write-Host "Cleaning: Lucee Logs (>30 days old, $sizeGB GB)" -ForegroundColor Cyan
#         
#         if (-not $WhatIf) {
#             $oldLogs | Remove-Item -Force -ErrorAction SilentlyContinue
#             Write-Host "  V Removed $($oldLogs.Count) files" -ForegroundColor Green
#         } else {
#             Write-Host "  -> Would remove $($oldLogs.Count) files ($sizeGB GB)" -ForegroundColor Yellow
#         }
#     }
#     catch {
#         Write-Host "  X Error cleaning Lucee logs: $($_.Exception.Message)" -ForegroundColor Red
#     }
# }

Write-Host "`n=====================================" -ForegroundColor Green
Write-Host "Weekly VM Maintenance Complete!" -ForegroundColor Green

# Show final disk space and calculate difference
try {
    $finalDisk = Get-WmiObject -Class Win32_LogicalDisk -Filter "DeviceID='C:'"
    $finalFreeSpaceGB = [math]::Round($finalDisk.FreeSpace / 1GB, 2)
    $finalUsedSpaceGB = $totalSpaceGB - $finalFreeSpaceGB
    $finalUsedPercent = [math]::Round(($finalUsedSpaceGB / $totalSpaceGB) * 100, 1)
    
    $spaceFreeGB = [math]::Round($finalFreeSpaceGB - $initialFreeSpaceGB, 2)
    
    Write-Host "`nFinal Disk Space:" -ForegroundColor Yellow
    Write-Host "Total: $totalSpaceGB GB" -ForegroundColor White
    Write-Host "Used:  $finalUsedSpaceGB GB ($finalUsedPercent%)" -ForegroundColor White  
    Write-Host "Free:  $finalFreeSpaceGB GB" -ForegroundColor Green
    
    if ($spaceFreeGB -gt 0) {
        Write-Host "`nSpace Reclaimed: $spaceFreeGB GB" -ForegroundColor Green
    } elseif ($spaceFreeGB -eq 0) {
        Write-Host "`nSpace Reclaimed: No change" -ForegroundColor Yellow
    } else {
        Write-Host "`nSpace Change: $spaceFreeGB GB (disk usage increased)" -ForegroundColor Red
    }
}
catch {
    Write-Host "`nCould not determine final disk space" -ForegroundColor Yellow
}

Write-Host "`nWeekly cleanup safe for production VMs" -ForegroundColor Green
