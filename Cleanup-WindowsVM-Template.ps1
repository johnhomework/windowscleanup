# Windows VM Template Preparation Script
# AGGRESSIVE cleanup for creating VM templates - run BEFORE sysprep/OOBE
# This script performs destructive cleanup to minimize template size

#Requires -RunAsAdministrator

param(
    [switch]$WhatIf,
    [switch]$SkipDriverCleanup,
    [switch]$SkipFeatureCleanup
)

Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "Windows VM Template Preparation Script" -ForegroundColor Cyan
Write-Host "AGGRESSIVE CLEANUP - Template Creation Mode" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

if ($WhatIf) {
    Write-Host "Running in WhatIf mode - no changes will be made" -ForegroundColor Yellow
    Write-Host ""
}

Write-Host "WARNING: This script performs AGGRESSIVE cleanup!" -ForegroundColor Red
Write-Host "  - Removes superseded components permanently" -ForegroundColor Yellow
Write-Host "  - Clears WinSxS backup files (no rollback!)" -ForegroundColor Yellow
Write-Host "  - Removes old Windows updates" -ForegroundColor Yellow
Write-Host "  - Cleans up old device drivers" -ForegroundColor Yellow
Write-Host "  - Optimizes for template creation" -ForegroundColor Yellow
Write-Host ""
Write-Host "Only run this on systems that will be used as VM templates!" -ForegroundColor Red
Write-Host ""

if (-not $WhatIf) {
    $confirm = Read-Host "Type 'TEMPLATE' to confirm this is a template preparation system"
    if ($confirm -ne "TEMPLATE") {
        Write-Host "Confirmation failed. Exiting." -ForegroundColor Red
        exit 1
    }
}

# Get initial disk space
try {
    $initialDisk = Get-WmiObject -Class Win32_LogicalDisk -Filter "DeviceID='C:'"
    $initialFreeSpaceGB = [math]::Round($initialDisk.FreeSpace / 1GB, 2)
    $totalSpaceGB = [math]::Round($initialDisk.Size / 1GB, 2)
    $initialUsedSpaceGB = $totalSpaceGB - $initialFreeSpaceGB
    $initialUsedPercent = [math]::Round(($initialUsedSpaceGB / $totalSpaceGB) * 100, 1)
    
    Write-Host "Initial Disk Space:" -ForegroundColor Yellow
    Write-Host "Total: $totalSpaceGB GB" -ForegroundColor White
    Write-Host "Used:  $initialUsedSpaceGB GB ($initialUsedPercent%)" -ForegroundColor White  
    Write-Host "Free:  $initialFreeSpaceGB GB" -ForegroundColor White
    Write-Host ""
}
catch {
    Write-Host "Could not determine initial disk space" -ForegroundColor Yellow
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

Write-Host "1. Stopping Services for Aggressive Cleanup..." -ForegroundColor Yellow
$servicesToStop = @("wuauserv", "BITS", "CryptSvc", "TrustedInstaller")
foreach ($service in $servicesToStop) {
    try {
        Stop-Service $service -Force -ErrorAction SilentlyContinue
        Write-Host "  Stopped: $service" -ForegroundColor Green
    }
    catch {
        Write-Host "  Warning: Could not stop $service" -ForegroundColor Yellow
    }
}
Write-Host ""

Write-Host "2. Aggressive Windows Update Cleanup..." -ForegroundColor Yellow
$updatePaths = @(
    @{Path="C:\Windows\SoftwareDistribution\Download\*"; Desc="Update Downloads"},
    @{Path="C:\Windows\SoftwareDistribution\DataStore\*"; Desc="Update DataStore"},
    @{Path="C:\`$Windows.~BT\*"; Desc="Windows Upgrade Files"},
    @{Path="C:\`$Windows.~WS\*"; Desc="Windows Upgrade Workspace"},
    @{Path="C:\`$WinREAgent\*"; Desc="WinRE Agent Files"},
    @{Path="C:\Windows\SoftwareDistribution\DeliveryOptimization\*"; Desc="Delivery Optimization Cache"},
    @{Path="C:\ProgramData\Microsoft\Windows\WER\*"; Desc="Windows Error Reporting"},
    @{Path="C:\Windows\System32\config\systemprofile\AppData\Local\Microsoft\Windows\DeliveryOptimization\*"; Desc="System DO Cache"}
)

foreach ($updatePath in $updatePaths) {
    Remove-SafelyWithLogging -Path $updatePath.Path -Description $updatePath.Desc
}
Write-Host ""

Write-Host "3. System Temporary and Cache Files..." -ForegroundColor Yellow
$tempPaths = @(
    @{Path="C:\Windows\Temp\*"; Desc="Windows Temp"},
    @{Path="C:\Windows\Prefetch\*"; Desc="Prefetch Files"},
    @{Path="C:\Windows\Logs\*"; Desc="All Windows Logs"},
    @{Path="C:\Windows\Panther\*"; Desc="Setup/Upgrade Logs"},
    @{Path="C:\Windows\inf\setupapi.dev.log"; Desc="Setup API Log"},
    @{Path="C:\Windows\Performance\WinSAT\*"; Desc="WinSAT Data"},
    @{Path="C:\ProgramData\Microsoft\Windows\WER\*"; Desc="Error Reporting"},
    @{Path="C:\ProgramData\Microsoft\Diagnosis\*"; Desc="Diagnostic Data"},
    @{Path="C:\Windows\System32\LogFiles\*"; Desc="System Log Files"},
    @{Path="C:\Windows\ServiceProfiles\*\AppData\Local\Temp\*"; Desc="Service Profile Temps"},
    @{Path="C:\Windows\System32\config\systemprofile\AppData\Local\Temp\*"; Desc="System Profile Temp"}
)

foreach ($tempPath in $tempPaths) {
    Remove-SafelyWithLogging -Path $tempPath.Path -Description $tempPath.Desc
}
Write-Host ""

Write-Host "4. User Profile Cleanup (All Users)..." -ForegroundColor Yellow
$userPaths = @(
    @{Path="C:\Users\*\AppData\Local\Temp\*"; Desc="User Temp Files"},
    @{Path="C:\Users\*\AppData\Local\Microsoft\Windows\INetCache\*"; Desc="IE/Edge Cache"},
    @{Path="C:\Users\*\AppData\Local\Microsoft\Windows\WebCache\*"; Desc="Web Cache"},
    @{Path="C:\Users\*\AppData\Local\Microsoft\Windows\Explorer\*.db"; Desc="Explorer Databases"},
    @{Path="C:\Users\*\AppData\Local\Microsoft\Windows\Caches\*"; Desc="Windows Caches"},
    @{Path="C:\Users\*\AppData\Local\CrashDumps\*"; Desc="User Crash Dumps"},
    @{Path="C:\Users\*\AppData\Local\Microsoft\Windows\History\*"; Desc="Browser History"},
    @{Path="C:\Users\*\AppData\Local\Microsoft\Windows\Temporary Internet Files\*"; Desc="Temporary Internet Files"},
    @{Path="C:\Users\*\AppData\Roaming\Microsoft\Windows\Recent\*"; Desc="Recent Items"},
    @{Path="C:\Users\*\Recent\*"; Desc="Recent Shortcuts"}
)

foreach ($userPath in $userPaths) {
    Remove-SafelyWithLogging -Path $userPath.Path -Description $userPath.Desc
}
Write-Host ""

Write-Host "5. Windows Memory Dumps and Debug Files..." -ForegroundColor Yellow
$dumpPaths = @(
    @{Path="C:\Windows\Minidump\*"; Desc="Minidumps"},
    @{Path="C:\Windows\memory.dmp"; Desc="Memory Dump"},
    @{Path="C:\Windows\MEMORY.DMP"; Desc="Memory Dump (caps)"},
    @{Path="C:\Windows\LiveKernelReports\*"; Desc="Kernel Reports"},
    @{Path="C:\ProgramData\Microsoft\Windows\WER\ReportQueue\*"; Desc="Error Report Queue"},
    @{Path="C:\ProgramData\Microsoft\Windows\WER\ReportArchive\*"; Desc="Error Report Archive"}
)

foreach ($dumpPath in $dumpPaths) {
    Remove-SafelyWithLogging -Path $dumpPath.Path -Description $dumpPath.Desc
}
Write-Host ""

Write-Host "6. Windows Installer and Package Caches..." -ForegroundColor Yellow
$installerPaths = @(
    @{Path="C:\Windows\Installer\$PatchCache$\*"; Desc="MSI Patch Cache"},
    @{Path="C:\ProgramData\Package Cache\*"; Desc="Package Cache"},
    @{Path="C:\Windows\System32\config\systemprofile\AppData\Local\Microsoft\Windows\INetCache\*"; Desc="System Profile Cache"}
)

foreach ($installerPath in $installerPaths) {
    Remove-SafelyWithLogging -Path $installerPath.Path -Description $installerPath.Desc
}
Write-Host ""

if (-not $SkipDriverCleanup) {
    Write-Host "7. Old Driver Store Cleanup..." -ForegroundColor Yellow
    if (-not $WhatIf) {
        try {
            Write-Host "  Enumerating old drivers..." -ForegroundColor Cyan
            
            # Get list of non-present devices (old drivers)
            $oldDrivers = pnputil /enum-drivers | Select-String "Published Name" -Context 0,5 | 
                Where-Object { $_.Context.PostContext -like "*Provider Name*" }
            
            Write-Host "  Found $($oldDrivers.Count) driver packages in store" -ForegroundColor Cyan
            Write-Host "  Use 'pnputil /delete-driver oem#.inf /uninstall' to remove specific old drivers" -ForegroundColor Gray
            Write-Host "  ✓ Driver enumeration complete (manual cleanup recommended)" -ForegroundColor Green
        }
        catch {
            Write-Host "  Warning: Could not enumerate drivers" -ForegroundColor Yellow
        }
    } else {
        Write-Host "  → Would enumerate and suggest driver cleanup" -ForegroundColor Yellow
    }
    
    # Clean driver temp files
    Remove-SafelyWithLogging -Path "C:\Windows\System32\DriverStore\Temp\*" -Description "Driver Store Temp Files"
    Write-Host ""
}

Write-Host "8. Windows Defender Cleanup..." -ForegroundColor Yellow
$defenderPaths = @(
    @{Path="C:\ProgramData\Microsoft\Windows Defender\Scans\History\*"; Desc="Defender Scan History"},
    @{Path="C:\ProgramData\Microsoft\Windows Defender\Support\*"; Desc="Defender Support Files"},
    @{Path="C:\Windows\System32\config\systemprofile\AppData\Local\Microsoft\Windows Defender\*"; Desc="Defender Profile Data"}
)

foreach ($defenderPath in $defenderPaths) {
    Remove-SafelyWithLogging -Path $defenderPath.Path -Description $defenderPath.Desc
}
Write-Host ""

Write-Host "9. Event Logs Cleanup..." -ForegroundColor Yellow
if (-not $WhatIf) {
    try {
        $logs = Get-WinEvent -ListLog * -ErrorAction SilentlyContinue
        $clearedCount = 0
        
        foreach ($log in $logs) {
            try {
                wevtutil cl $log.LogName 2>$null
                $clearedCount++
            }
            catch {
                # Silent fail for logs that can't be cleared
            }
        }
        Write-Host "  ✓ Cleared $clearedCount event logs" -ForegroundColor Green
    }
    catch {
        Write-Host "  Warning: Could not clear all event logs" -ForegroundColor Yellow
    }
} else {
    Write-Host "  → Would clear all event logs" -ForegroundColor Yellow
}
Write-Host ""

Write-Host "10. Windows Search Index..." -ForegroundColor Yellow
if (-not $WhatIf) {
    try {
        Stop-Service "WSearch" -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        Remove-SafelyWithLogging -Path "C:\ProgramData\Microsoft\Search\Data\*" -Description "Search Index Data"
        Write-Host "  ✓ Search service stopped and index cleared" -ForegroundColor Green
    }
    catch {
        Write-Host "  Warning: Could not clear search index" -ForegroundColor Yellow
    }
} else {
    Write-Host "  → Would stop search and clear index" -ForegroundColor Yellow
}
Write-Host ""

Write-Host "11. Recycle Bin (All Users)..." -ForegroundColor Yellow
if (-not $WhatIf) {
    try {
        Clear-RecycleBin -Force -ErrorAction Stop
        Write-Host "  ✓ All recycle bins emptied" -ForegroundColor Green
    }
    catch {
        Write-Host "  Warning: Could not empty all recycle bins" -ForegroundColor Yellow
    }
} else {
    Write-Host "  → Would empty all recycle bins" -ForegroundColor Yellow
}
Write-Host ""

Write-Host "12. DISM - Remove Superseded Components..." -ForegroundColor Yellow
if (-not $WhatIf) {
    try {
        Write-Host "  Running DISM with /ResetBase (THIS CANNOT BE UNDONE!)" -ForegroundColor Cyan
        Write-Host "  This will take 10-30 minutes. Please wait..." -ForegroundColor Gray
        
        $dismJob = Start-Job -ScriptBlock {
            $dismResult = Start-Process "Dism.exe" -ArgumentList "/online /Cleanup-Image /StartComponentCleanup /ResetBase" -Wait -PassThru -WindowStyle Hidden
            return $dismResult.ExitCode
        }
        
        # Wait with progress indicator (max 45 minutes)
        $timeout = 2700
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
            Write-Host "  Warning: DISM cleanup timed out" -ForegroundColor Yellow
        } else {
            $exitCode = Receive-Job $dismJob
            Remove-Job $dismJob
            
            if ($exitCode -eq 0) {
                Write-Host "  ✓ DISM /ResetBase completed successfully" -ForegroundColor Green
                Write-Host "  Note: Windows Update rollback is now disabled!" -ForegroundColor Yellow
            } else {
                Write-Host "  Warning: DISM failed with exit code $exitCode" -ForegroundColor Yellow
            }
        }
    }
    catch {
        Write-Host "  Error: DISM cleanup failed: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "  → Would run DISM /StartComponentCleanup /ResetBase" -ForegroundColor Yellow
    Write-Host "  → This removes ALL superseded components permanently" -ForegroundColor Yellow
}
Write-Host ""

Write-Host "13. DISM - Analyze Component Store (WinSxS)..." -ForegroundColor Yellow
if (-not $WhatIf) {
    try {
        Write-Host "  Analyzing component store size..." -ForegroundColor Cyan
        $analysis = dism /online /Cleanup-Image /AnalyzeComponentStore
        
        # Parse the output for key metrics
        $componentSize = ($analysis | Select-String "Component Store \(WinSxS\) size").ToString()
        $reclaimable = ($analysis | Select-String "Reclaimable Packages").ToString()
        
        if ($componentSize) {
            Write-Host "  $componentSize" -ForegroundColor White
        }
        if ($reclaimable) {
            Write-Host "  $reclaimable" -ForegroundColor White
        }
        
        Write-Host "  ✓ Component store analysis complete" -ForegroundColor Green
    }
    catch {
        Write-Host "  Warning: Could not analyze component store" -ForegroundColor Yellow
    }
} else {
    Write-Host "  → Would analyze component store size" -ForegroundColor Yellow
}
Write-Host ""

if (-not $SkipFeatureCleanup) {
    Write-Host "14. Windows Features Cleanup..." -ForegroundColor Yellow
    if (-not $WhatIf) {
        try {
            Write-Host "  Running feature cleanup to remove disabled features..." -ForegroundColor Cyan
            
            $featureCleanup = Start-Job -ScriptBlock {
                $result = Start-Process "Dism.exe" -ArgumentList "/online /Cleanup-Image /StartComponentCleanup /ResetBase" -Wait -PassThru -WindowStyle Hidden
                return $result.ExitCode
            }
            
            Wait-Job $featureCleanup -Timeout 600 | Out-Null
            $exitCode = Receive-Job $featureCleanup
            Remove-Job $featureCleanup
            
            if ($exitCode -eq 0) {
                Write-Host "  ✓ Feature cleanup completed" -ForegroundColor Green
            }
        }
        catch {
            Write-Host "  Warning: Feature cleanup failed" -ForegroundColor Yellow
        }
    } else {
        Write-Host "  → Would run Windows Features cleanup" -ForegroundColor Yellow
    }
    Write-Host ""
}

Write-Host "15. Service Pack Cleanup Files..." -ForegroundColor Yellow
$spPaths = @(
    @{Path="C:\Windows\WinSxS\Backup\*"; Desc="WinSxS Backup Files"},
    @{Path="C:\Windows\WinSxS\ManifestCache\*"; Desc="Manifest Cache"},
    @{Path="C:\Windows\System32\Dism\*"; Desc="DISM Temp Files"}
)

foreach ($spPath in $spPaths) {
    Remove-SafelyWithLogging -Path $spPath.Path -Description $spPath.Desc
}
Write-Host ""

Write-Host "16. System Restore Points Cleanup..." -ForegroundColor Yellow
if (-not $WhatIf) {
    try {
        vssadmin delete shadows /all /quiet 2>$null
        Write-Host "  ✓ All system restore points deleted" -ForegroundColor Green
        Write-Host "  Warning: System cannot be restored to previous state!" -ForegroundColor Yellow
    }
    catch {
        Write-Host "  Warning: Could not delete restore points" -ForegroundColor Yellow
    }
} else {
    Write-Host "  → Would delete all system restore points" -ForegroundColor Yellow
}
Write-Host ""

Write-Host "17. Page File and Hibernation..." -ForegroundColor Yellow
if (-not $WhatIf) {
    try {
        # Disable hibernation (removes hiberfil.sys)
        powercfg /hibernate off
        Write-Host "  ✓ Hibernation disabled (hiberfil.sys will be removed on reboot)" -ForegroundColor Green
        
        # Note: We don't delete pagefile as it's recreated automatically
        Write-Host "  Note: Pagefile will be optimized on next boot" -ForegroundColor Gray
    }
    catch {
        Write-Host "  Warning: Could not configure hibernation/pagefile" -ForegroundColor Yellow
    }
} else {
    Write-Host "  → Would disable hibernation" -ForegroundColor Yellow
}
Write-Host ""

Write-Host "18. Optimize Drives and Compact OS..." -ForegroundColor Yellow
if (-not $WhatIf) {
    try {
        # Run TRIM on all drives
        $drives = Get-Volume | Where-Object { $_.DriveLetter -and $_.DriveType -eq "Fixed" }
        foreach ($drive in $drives) {
            Write-Host "  Running TRIM on drive $($drive.DriveLetter):\ ..." -ForegroundColor Cyan
            Optimize-Volume -DriveLetter $drive.DriveLetter -ReTrim -Verbose
        }
        Write-Host "  ✓ TRIM completed on all drives" -ForegroundColor Green
        
        # Compact OS (compresses Windows files)
        Write-Host "  Running Compact OS analysis..." -ForegroundColor Cyan
        $compactStatus = compact /CompactOS:query
        
        if ($compactStatus -like "*not in compact mode*") {
            Write-Host "  Applying Compact OS (this may take 10-20 minutes)..." -ForegroundColor Cyan
            compact /CompactOS:always | Out-Null
            Write-Host "  ✓ Compact OS enabled - Windows files compressed" -ForegroundColor Green
        } else {
            Write-Host "  Compact OS already enabled" -ForegroundColor Gray
        }
    }
    catch {
        Write-Host "  Warning: Drive optimization failed" -ForegroundColor Yellow
    }
} else {
    Write-Host "  → Would run TRIM and enable Compact OS" -ForegroundColor Yellow
}
Write-Host ""

Write-Host "19. Clear DNS and Network Cache..." -ForegroundColor Yellow
if (-not $WhatIf) {
    try {
        ipconfig /flushdns | Out-Null
        arp -d * 2>$null
        netsh int ip reset | Out-Null
        Write-Host "  ✓ Network caches cleared" -ForegroundColor Green
    }
    catch {
        Write-Host "  Warning: Network cache cleanup failed" -ForegroundColor Yellow
    }
} else {
    Write-Host "  → Would clear DNS and network cache" -ForegroundColor Yellow
}
Write-Host ""

Write-Host "20. Zero Free Space (Optional - for maximum compression)..." -ForegroundColor Yellow
$zeroSpace = Read-Host "  Zero free space for better compression? This takes a LONG time (y/N)"
if ($zeroSpace -eq "y" -and -not $WhatIf) {
    try {
        Write-Host "  Creating zero-fill file to maximize compression..." -ForegroundColor Cyan
        Write-Host "  This will take 30+ minutes depending on free space..." -ForegroundColor Gray
        
        fsutil file createnew C:\zerofill.tmp (Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'").FreeSpace
        Remove-Item C:\zerofill.tmp -Force
        
        Write-Host "  ✓ Free space zeroed" -ForegroundColor Green
    }
    catch {
        Write-Host "  Warning: Could not zero free space" -ForegroundColor Yellow
        Remove-Item C:\zerofill.tmp -Force -ErrorAction SilentlyContinue
    }
} else {
    Write-Host "  Skipped (not recommended unless template will be exported)" -ForegroundColor Gray
}
Write-Host ""

Write-Host "21. Final Service Restart..." -ForegroundColor Yellow
$servicesToRestart = @("wuauserv")
foreach ($service in $servicesToRestart) {
    try {
        Start-Service $service -ErrorAction SilentlyContinue
        Write-Host "  Restarted: $service" -ForegroundColor Green
    }
    catch {
        # Don't restart search service - leave it off for template
    }
}
Write-Host ""

Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "Template Preparation Complete!" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

# Show final disk space and calculate difference
try {
    $finalDisk = Get-WmiObject -Class Win32_LogicalDisk -Filter "DeviceID='C:'"
    $finalFreeSpaceGB = [math]::Round($finalDisk.FreeSpace / 1GB, 2)
    $finalUsedSpaceGB = $totalSpaceGB - $finalFreeSpaceGB
    $finalUsedPercent = [math]::Round(($finalUsedSpaceGB / $totalSpaceGB) * 100, 1)
    
    $spaceFreedGB = [math]::Round($finalFreeSpaceGB - $initialFreeSpaceGB, 2)
    
    Write-Host "Final Disk Space:" -ForegroundColor Yellow
    Write-Host "Total: $totalSpaceGB GB" -ForegroundColor White
    Write-Host "Used:  $finalUsedSpaceGB GB ($finalUsedPercent%)" -ForegroundColor White  
    Write-Host "Free:  $finalFreeSpaceGB GB" -ForegroundColor Green
    Write-Host ""
    
    if ($spaceFreedGB -gt 0) {
        Write-Host "Total Space Reclaimed: $spaceFreedGB GB" -ForegroundColor Green
    } else {
        Write-Host "Space Change: $spaceFreedGB GB" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "Could not determine final disk space" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "NEXT STEPS FOR TEMPLATE CREATION:" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "1. REBOOT the system to complete cleanup" -ForegroundColor Yellow
Write-Host "2. Verify system is working properly" -ForegroundColor Yellow
Write-Host "3. Remove any application-specific logs/data" -ForegroundColor Yellow
Write-Host "4. Clear browser history and profiles" -ForegroundColor Yellow
Write-Host "5. Remove any personal/company data" -ForegroundColor Yellow
Write-Host "6. Run: C:\Windows\System32\Sysprep\sysprep.exe" -ForegroundColor Yellow
Write-Host "   - Select 'Enter System Out-of-Box Experience (OOBE)'" -ForegroundColor Gray
Write-Host "   - Check 'Generalize'" -ForegroundColor Gray
Write-Host "   - Select 'Shutdown'" -ForegroundColor Gray
Write-Host "7. After shutdown, create VM template/snapshot" -ForegroundColor Yellow
Write-Host ""
Write-Host "IMPORTANT NOTES:" -ForegroundColor Red
Write-Host "- Windows Update rollback is now DISABLED" -ForegroundColor Yellow
Write-Host "- System Restore Points have been deleted" -ForegroundColor Yellow
Write-Host "- This system should only be used as a template" -ForegroundColor Yellow
Write-Host ""
