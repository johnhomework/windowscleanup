# Windows VM Weekly Maintenance Script
# Safe for regular production VM cleanup - focuses on disk space recovery
# Avoids system-critical items like event logs and user profile data

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
                Write-Host "  ✓ Removed" -ForegroundColor Green
            } else {
                Write-Host "  → Would remove $sizeGB GB" -ForegroundColor Yellow
            }
        }
        catch {
            Write-Host "  ✗ Error: $($_.Exception.Message)" -ForegroundColor Red
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

Write-Host "`n10. Docker and Container Cache (if present)..." -ForegroundColor Yellow
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
        # Use built-in PowerShell cmdlets instead of cleanmgr.exe to avoid GUI issues
        Write-Host "  Running automated disk cleanup..." -ForegroundColor Cyan
        
        # Clear additional temp locations that cleanmgr would handle
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
        
        Write-Host "  ✓ Disk cleanup completed (automated method)" -ForegroundColor Green
    }
    catch {
        Write-Host "  Warning: Disk cleanup failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }
} else {
    Write-Host "  → Would run automated disk cleanup (no GUI)" -ForegroundColor Yellow
}

Write-Host "`n12. Empty Recycle Bin..." -ForegroundColor Yellow
if (-not $WhatIf) {
    try {
        Clear-RecycleBin -Force -ErrorAction Stop
        Write-Host "  ✓ Recycle Bin emptied" -ForegroundColor Green
    }
    catch {
        try {
            $recycleBin = New-Object -ComObject Shell.Application
            $recycleBin.Namespace(0xA).InvokeVerb("empty")
            Write-Host "  ✓ Recycle Bin emptied" -ForegroundColor Green
        }
        catch {
            Write-Host "  Warning: Could not empty Recycle Bin" -ForegroundColor Yellow
        }
    }
} else {
    Write-Host "  → Would empty Recycle Bin" -ForegroundColor Yellow
}

Write-Host "`n13. DISM Component Cleanup (Safe Mode)..." -ForegroundColor Yellow
if (-not $WhatIf) {
    try {
        Write-Host "  Running DISM component cleanup (safe mode - no ResetBase)..." -ForegroundColor Cyan
        
        # Run DISM cleanup WITHOUT ResetBase (keeps ability to uninstall updates)
        $dismResult = Start-Process "Dism.exe" -ArgumentList "/online /Cleanup-Image /StartComponentCleanup" -Wait -PassThru -WindowStyle Hidden
        
        if ($dismResult.ExitCode -eq 0) {
            Write-Host "  ✓ DISM component cleanup completed" -ForegroundColor Green
        } else {
            Write-Host "  Warning: DISM cleanup failed with exit code $($dismResult.ExitCode)" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "  Error: DISM cleanup failed: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "  → Would run DISM component cleanup (safe mode)" -ForegroundColor Yellow
}

Write-Host "`n14. Disk Optimization..." -ForegroundColor Yellow
if (-not $WhatIf) {
    try {
        # Get all drives for optimization
        $drives = Get-Volume | Where-Object { $_.DriveLetter -and $_.DriveType -eq "Fixed" }
        
        foreach ($drive in $drives) {
            Write-Host "  Running TRIM on drive $($drive.DriveLetter):\ ..." -ForegroundColor Cyan
            Optimize-Volume -DriveLetter $drive.DriveLetter -ReTrim -Verbose
            Write-Host "  ✓ TRIM completed for drive $($drive.DriveLetter):\" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "  Warning: Disk optimization failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }
} else {
    Write-Host "  → Would run TRIM on all fixed drives" -ForegroundColor Yellow
}

Write-Host "`n15. Network Cache Cleanup..." -ForegroundColor Yellow
if (-not $WhatIf) {
    try {
        # Clear DNS cache
        ipconfig /flushdns | Out-Null
        Write-Host "  ✓ DNS cache cleared" -ForegroundColor Green
        
        # Clear ARP cache  
        arp -d * 2>$null
        Write-Host "  ✓ ARP cache cleared" -ForegroundColor Green
    }
    catch {
        Write-Host "  Warning: Network cache cleanup failed" -ForegroundColor Yellow
    }
} else {
    Write-Host "  → Would clear DNS and ARP cache" -ForegroundColor Yellow
}

# Restart stopped services
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

Write-Host "`n=====================================" -ForegroundColor Green
Write-Host "Weekly VM Maintenance Complete!" -ForegroundColor Green

# Show disk space summary
try {
    $disk = Get-WmiObject -Class Win32_LogicalDisk -Filter "DeviceID='C:'"
    $freeSpaceGB = [math]::Round($disk.FreeSpace / 1GB, 2)
    $totalSpaceGB = [math]::Round($disk.Size / 1GB, 2)
    $usedSpaceGB = $totalSpaceGB - $freeSpaceGB
    $usedPercent = [math]::Round(($usedSpaceGB / $totalSpaceGB) * 100, 1)
    
    Write-Host "`nDisk Space Summary:" -ForegroundColor Yellow
    Write-Host "Total: $totalSpaceGB GB" -ForegroundColor White
    Write-Host "Used:  $usedSpaceGB GB ($usedPercent%)" -ForegroundColor White  
    Write-Host "Free:  $freeSpaceGB GB" -ForegroundColor Green
}
catch {
    Write-Host "`nCould not determine disk space" -ForegroundColor Yellow
}

Write-Host "`nWeekly cleanup safe for production VMs ✓" -ForegroundColor Green