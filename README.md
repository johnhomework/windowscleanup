# Windows VM Cleanup Scripts

Two PowerShell scripts I use for Windows VMs — one to prep a template (destructive, one-time), and one safe weekly cleanup for running VMs.

## Cleanup-WindowsVM-Template.ps1 — Template prep (DO NOT run on production)

Use this once before sysprepping a VM. It nukes user-specific junk so the template is clean.

### What it does
- Clears all event logs
- Removes user profiles, Documents and user data caches
- Removes saved network profiles and cached credentials
- Runs DISM `/ResetBase` (permanent — you can't uninstall old updates afterward)
- Disables hibernation and sets sane VM power settings
- Zeros free space for much better template compression

### Run
```powershell
.\Cleanup-WindowsVM-Template.ps1
# Dry-run:
.\Cleanup-WindowsVM-Template.ps1 -WhatIf
```

**Warning:** destructive. Verify before running. Run once, then sysprep and capture the template.

## Cleanup-WindowsVM-Weekly.ps1 — Safe weekly maintenance

Run this weekly on production VMs. It reclaims disk space and removes caches without touching user data or event logs.

### What it does
- System and user temp files cleanup
- Dev caches: npm, nuget, gradle, pip wheel caches, etc.
- Common app caches: Office, Teams, Zoom, Adobe (where safe)
- Windows Store app cache (safe clear)
- IIS logs, webserver temp files, application logs rotation/cleanup
- Docker / container image & volume cleanup where applicable
- Dism `/Online /Cleanup-Image` (no `/ResetBase`)
- TRIM / Optimize-Volume where supported
- Reports what it cleaned and total space reclaimed

### Run
```powershell
.\Cleanup-WindowsVM-Weekly.ps1
# Dry-run:
.\Cleanup-WindowsVM-Weekly.ps1 -WhatIf
# Skip browser cleanup:
.\Cleanup-WindowsVM-Weekly.ps1 -SkipBrowserCleanup
```

## Why these replace Disk Cleanup

The weekly script covers everything Disk Cleanup does and adds dev/container/app cache cleanup that Disk Cleanup misses. The template script goes further — it prepares a VM for sysprep and templating (including DISM `/ResetBase` and zeroing free space).

## Requirements & notes

- Run as Administrator
- PowerShell 5.0+ (Windows 10 / Server 2016 or newer)
- Template script is destructive — double-check and test on a throwaway VM first
- Both scripts print what they do and how much space they reclaimed
- Always snapshot/backup before host-side image operations (repacking/shrinking)
