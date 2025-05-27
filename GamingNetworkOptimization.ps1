# Script for System-Level Network Optimizations (Complementing Manual NIC Settings)
# Version: Leaves Congestion Control Provider as Default (CUBIC)
#          Includes Pre-Execution Checks, Improved Error Handling, GUID Validation, 
#          Accepts TargetGuidsCsv, Persists after execution, and separates 'rsc=disabled' netsh.
#          Refined null handling for current value display.
# Run as Administrator
# Designed to be used WITH manual configuration of NIC advanced properties.

param (
    [string]$TargetGuidsCsv = "" # Expecting a comma-separated string of GUIDs from the launcher
)

$ScriptEncounteredErrors = $false # Flag to track if any warnings/errors occur
$ChangesMade = $false # Flag to track if any actual changes were made to settings

Write-Host "Applying system-level network optimizations with pre-execution checks..." -ForegroundColor Yellow
Write-Host "------------------------------------------------------------"
Write-Host "NOTE: This script assumes you have manually configured your"
Write-Host "Network Adapter's Advanced Properties for settings like:"
Write-Host "- Large Send Offload (LSO)"
Write-Host "- Checksum Offloads (TCP/UDP/IP)"
Write-Host "- Receive Side Scaling (RSS) enablement and queue numbers"
Write-Host "- Energy Efficient Ethernet (EEE), Flow Control, Interrupt Moderation, etc."
Write-Host "This script will NOT change the TCP Congestion Control Provider."
Write-Host "------------------------------------------------------------"

Function Test-AndSetRegistryValue {
    param (
        [string]$Path,
        [string]$Name,
        [object]$DesiredValue,
        [Microsoft.Win32.RegistryValueKind]$Type,
        [string]$SuccessMessage
    )
    try {
        $CurrentValueObject = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
        $CurrentValue = if ($null -ne $CurrentValueObject) { $CurrentValueObject.$Name } else { $null }

        if ($CurrentValue -eq $DesiredValue) {
            Write-Host "    - INFO: $SuccessMessage already set to '$DesiredValue'. No change needed." -ForegroundColor Gray
        } else {
            if (-not (Test-Path $Path)) { New-Item -Path $Path -Force -ErrorAction Stop | Out-Null }
            Set-ItemProperty -Path $Path -Name $Name -Value $DesiredValue -Type $Type -Force -ErrorAction Stop | Out-Null
            $PreviousValueDisplay = if ($null -eq $CurrentValue) { "[Does Not Exist/Null]" } else { "'$CurrentValue'" }
            Write-Host "    - SUCCESS: $SuccessMessage set to '$DesiredValue' (was $PreviousValueDisplay)." -ForegroundColor Green
            $script:ChangesMade = $true
        }
    } catch {
        Write-Warning "    * FAILURE: Could not set registry value '$Name' at path '$Path'. Error: $($_.Exception.Message)"
        $script:ScriptEncounteredErrors = $true
    }
}

# --- Registry Tweaks ---
Write-Host "[REGISTRY] Applying registry tweaks..."

# MSMQ TCPNoDelay
Write-Host "  [REGISTRY] MSMQ Parameters:"
Test-AndSetRegistryValue -Path "HKLM:\SOFTWARE\Microsoft\MSMQ\Parameters" -Name "TCPNoDelay" -DesiredValue 1 -Type DWord -SuccessMessage "TCPNoDelay (Disable Nagle for MSMQ)"

# --- Interface-Specific Tweaks (Registry) based on $TargetGuidsCsv ---
$InterfaceGuidsToProcess = $TargetGuidsCsv -split ',' | ForEach-Object {$_.Trim()} | Where-Object {$_ -ne ""}

if ($InterfaceGuidsToProcess.Count -eq 0) {
    Write-Warning "  [REGISTRY] No Target GUIDs were provided or selected for interface-specific tweaks. Skipping this section."
} else {
    Write-Host "  [REGISTRY] Applying Interface-Specific TCP Settings for selected GUID(s)..."
    foreach ($CurrentInterfaceGuidInLoop in $InterfaceGuidsToProcess) {
        if ($CurrentInterfaceGuidInLoop -notmatch "^{?[A-Fa-f0-9]{8}-([A-Fa-f0-9]{4}-){3}[A-Fa-f0-9]{12}}?$") {
            Write-Error "    * INVALID GUID FORMAT: '$CurrentInterfaceGuidInLoop'. Skipping this entry."
            $ScriptEncounteredErrors = $true
            continue
        }
        $InterfaceRegPath = "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\$CurrentInterfaceGuidInLoop"
        Write-Host "    Processing Interface GUID: $CurrentInterfaceGuidInLoop (Path: $InterfaceRegPath)"
        if (-not (Test-Path $InterfaceRegPath)) {
            Write-Error "      * CRITICAL ERROR: The registry path for Interface GUID '$CurrentInterfaceGuidInLoop' does NOT exist. Settings SKIPPED."
            $ScriptEncounteredErrors = $true
            continue
        }
        Test-AndSetRegistryValue -Path $InterfaceRegPath -Name "TcpAckFrequency" -DesiredValue 1 -Type DWord -SuccessMessage "TcpAckFrequency (Immediate ACKs) for $CurrentInterfaceGuidInLoop"
        Test-AndSetRegistryValue -Path $InterfaceRegPath -Name "TCPNoDelay" -DesiredValue 1 -Type DWord -SuccessMessage "TCPNoDelay (Disable Nagle) for $CurrentInterfaceGuidInLoop"
        Test-AndSetRegistryValue -Path $InterfaceRegPath -Name "TcpDelAckTicks" -DesiredValue 0 -Type DWord -SuccessMessage "TcpDelAckTicks (ACK Delay Timeout) for $CurrentInterfaceGuidInLoop"
    }
}

# Multimedia System Profile
Write-Host "  [REGISTRY] Multimedia System Profile:"
Test-AndSetRegistryValue -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile" -Name "NetworkThrottlingIndex" -DesiredValue 0xFFFFFFFF -Type DWord -SuccessMessage "NetworkThrottlingIndex (Disable Network Throttling)"
Test-AndSetRegistryValue -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile" -Name "SystemResponsiveness" -DesiredValue 0 -Type DWord -SuccessMessage "SystemResponsiveness (Prioritize Foreground Apps)"

# ServiceProvider Host Resolution Priorities
Write-Host "  [REGISTRY] ServiceProvider Host Resolution Priorities:"
try {
    $ServiceProviderPath = "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\ServiceProvider"
    if (-not (Test-Path $ServiceProviderPath)) { New-Item -Path $ServiceProviderPath -Force -ErrorAction Stop | Out-Null }
    # Check current values to see if changes will be made for the $ChangesMade flag
    $LP = (Get-ItemProperty -Path $ServiceProviderPath -Name LocalPriority -ErrorAction SilentlyContinue).LocalPriority
    $HP = (Get-ItemProperty -Path $ServiceProviderPath -Name HostsPriority -ErrorAction SilentlyContinue).HostsPriority
    $DP = (Get-ItemProperty -Path $ServiceProviderPath -Name DnsPriority -ErrorAction SilentlyContinue).DnsPriority
    $NP = (Get-ItemProperty -Path $ServiceProviderPath -Name NetbtPriority -ErrorAction SilentlyContinue).NetbtPriority

    $PrioritiesChanged = $false
    if ($LP -ne 4) { Set-ItemProperty -Path $ServiceProviderPath -Name "LocalPriority" -Value 4 -Type DWord -Force -ErrorAction Stop | Out-Null; $PrioritiesChanged = $true }
    if ($HP -ne 5) { Set-ItemProperty -Path $ServiceProviderPath -Name "HostsPriority" -Value 5 -Type DWord -Force -ErrorAction Stop | Out-Null; $PrioritiesChanged = $true }
    if ($DP -ne 6) { Set-ItemProperty -Path $ServiceProviderPath -Name "DnsPriority" -Value 6 -Type DWord -Force -ErrorAction Stop | Out-Null; $PrioritiesChanged = $true }
    if ($NP -ne 7) { Set-ItemProperty -Path $ServiceProviderPath -Name "NetbtPriority" -Value 7 -Type DWord -Force -ErrorAction Stop | Out-Null; $PrioritiesChanged = $true }

    if ($PrioritiesChanged) {
        Write-Host "    - SUCCESS: Set Priorities (Local:4, Hosts:5, DNS:6, NetBT:7)." -ForegroundColor Green
        $script:ChangesMade = $true
    } else {
        Write-Host "    - INFO: ServiceProvider priorities already optimally set. No change needed." -ForegroundColor Gray
    }
} catch { 
    Write-Warning "    * FAILURE: Could not set ServiceProvider priorities at path '$ServiceProviderPath'. Error: $($_.Exception.Message)"
    $script:ScriptEncounteredErrors = $true
}

# Global TCP/IP Parameters (Registry)
Write-Host "  [REGISTRY] Global TCP/IP Parameters:"
Test-AndSetRegistryValue -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters" -Name "MaxUserPort" -DesiredValue 65534 -Type DWord -SuccessMessage "MaxUserPort (Ephemeral Port Range)"
Test-AndSetRegistryValue -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters" -Name "TcpTimedWaitDelay" -DesiredValue 30 -Type DWord -SuccessMessage "TcpTimedWaitDelay (TIME_WAIT Interval)"
Test-AndSetRegistryValue -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters" -Name "DefaultTTL" -DesiredValue 64 -Type DWord -SuccessMessage "DefaultTTL (Packet Time-To-Live)"

# QoS Tweaks (Registry)
Write-Host "  [REGISTRY] Quality of Service (QoS) Tweaks:"
Test-AndSetRegistryValue -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Psched" -Name "NonBestEffortLimit" -DesiredValue 0 -Type DWord -SuccessMessage "NonBestEffortLimit (Disable Bandwidth Limit)"
Test-AndSetRegistryValue -Path "HKLM:\System\CurrentControlSet\Services\Tcpip\QoS" -Name "Do not use NLA" -DesiredValue "1" -Type String -SuccessMessage "'Do not use NLA' (QoS Ignores NLA)"

# Memory Management (Registry)
Write-Host "  [REGISTRY] Memory Management Tweak:"
Test-AndSetRegistryValue -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management" -Name "LargeSystemCache" -DesiredValue 0 -Type DWord -SuccessMessage "LargeSystemCache (Prioritize Application Memory)"

Write-Host "------------------------------------------------------------"
Write-Host "[NETSH/CMDLET] Applying system-level command-line tweaks..."

# TCP Global Parameters (Netsh) - With Pre-Checks
Write-Host "  [NETSH] TCP Global Parameters:"
try {
    $GlobalTcpSettingsOutput = Invoke-Expression "netsh interface tcp show global" # Capture output once using Invoke-Expression
    $ChangesMadeThisNetshBlock = $false 

    Function Test-AndSetNetshGlobal {
        param(
            [string]$SettingNamePattern, 
            [string]$NetshParamName, 
            [string]$DesiredValueString, 
            [string]$SuccessMessagePart
        )
        
        $CurrentValueString = $null # Initialize
        # Precise pattern: Start of line, exact setting name, spaces, colon, spaces, capture rest of line
        $Pattern = "^$SettingNamePattern\s+:\s+(.+)$" 
        
        foreach ($line in ($script:GlobalTcpSettingsOutput -split [System.Environment]::NewLine)) {
            if ($line -match $Pattern) {
                $CurrentValueString = $Matches[1].Trim()
                break # Found the line
            }
        }

        if ($CurrentValueString -ne $null -and $CurrentValueString -eq $DesiredValueString) {
            Write-Host "    - INFO: Global TCP $SuccessMessagePart already '$DesiredValueString'. No change needed." -ForegroundColor Gray
        } else {
            $CommandToExecute = "netsh.exe interface tcp set global $($NetshParamName)='$($DesiredValueString)'"
            try {
                Invoke-Expression $CommandToExecute | Out-Null # Use Invoke-Expression to run netsh
                $PreviousValueDisplayNetsh = if ($null -eq $CurrentValueString) { "[Not Found in Prior Output/Null]" } else { "'$CurrentValueString'" }
                Write-Host "    - SUCCESS: Global TCP $SuccessMessagePart set to '$DesiredValueString' (was $PreviousValueDisplayNetsh). Command: $CommandToExecute" -ForegroundColor Green
                $script:ChangesMade = $true
                $script:ChangesMadeThisNetshBlock = $true
            } catch {
                Write-Warning "    * FAILURE executing command '$CommandToExecute'. Error: $($_.Exception.Message)"
                $script:ScriptEncounteredErrors = $true
            }
        }
    }

    # DCA - Attempt enable.
    $NetshCommandDca = "netsh.exe interface tcp set global dca=enabled"
    try {
        Invoke-Expression $NetshCommandDca | Out-Null
        Write-Host "    - ACTION: Command '$NetshCommandDca' executed (Attempted to enable Direct Cache Access)." -ForegroundColor Cyan
    } catch {
        Write-Warning "    * FAILURE executing command '$NetshCommandDca'. Error: $($_.Exception.Message)"
        $script:ScriptEncounteredErrors = $true
    }
    
    # RSS Global State - target: enabled
    Test-AndSetNetshGlobal -SettingNamePattern "Receive-Side Scaling State" -NetshParamName "rss" -DesiredValueString "enabled" -SuccessMessagePart "Receive-Side Scaling State"
    # RSC Global State - target: disabled
    Test-AndSetNetshGlobal -SettingNamePattern "Receive Segment Coalescing State" -NetshParamName "rsc" -DesiredValueString "disabled" -SuccessMessagePart "Receive Segment Coalescing State (Key for low latency)"
    # Timestamps - target: disabled
    Test-AndSetNetshGlobal -SettingNamePattern "RFC 1323 Timestamps" -NetshParamName "timestamps" -DesiredValueString "disabled" -SuccessMessagePart "RFC 1323 Timestamps"
    # InitialRTO - target: 2000
    Test-AndSetNetshGlobal -SettingNamePattern "Initial RTO" -NetshParamName "initialrto" -DesiredValueString "2000" -SuccessMessagePart "Initial RTO"
    # NonSackRttResiliency - target: disabled
    Test-AndSetNetshGlobal -SettingNamePattern "Non Sack Rtt Resiliency" -NetshParamName "nonsackrttresiliency" -DesiredValueString "disabled" -SuccessMessagePart "Non SACK RTT Resiliency"
    # MaxSYNRetransmissions - target: 2
    Test-AndSetNetshGlobal -SettingNamePattern "Max SYN Retransmissions" -NetshParamName "maxsynretransmissions" -DesiredValueString "2" -SuccessMessagePart "Max SYN Retransmissions"

    if (-not $ChangesMadeThisNetshBlock) { 
         Write-Host "    - INFO: All checked global TCP parameters (RSC, Timestamps, RTOs etc.) appear to be optimally set or DCA was attempted." -ForegroundColor Gray
    }
    
} catch { 
    Write-Warning "    * FAILURE: Could not apply one or more 'netsh int tcp set global' commands or parse current settings. Error: $($_.Exception.Message)"
    $script:ScriptEncounteredErrors = $true
}

# Global Offload Settings (PowerShell Cmdlet)
Write-Host "  [CMDLET] Global Offload Settings:"
try {
    $CurrentOffloadSettings = Get-NetOffloadGlobalSetting -ErrorAction SilentlyContinue
    if ($CurrentOffloadSettings) {
        if ($CurrentOffloadSettings.Chimney -ne "Disabled") {
            Set-NetOffloadGlobalSetting -Chimney Disabled | Out-Null
            Write-Host "    - SUCCESS: TCP Chimney Offload set to Disabled (was '$($CurrentOffloadSettings.Chimney)')." -ForegroundColor Green
            $script:ChangesMade = $true
        } else {
            Write-Host "    - INFO: TCP Chimney Offload already Disabled. No change needed." -ForegroundColor Gray
        }
        if ($CurrentOffloadSettings.PacketCoalescingFilter -ne "Disabled") {
            Set-NetOffloadGlobalSetting -PacketCoalescingFilter Disabled | Out-Null
            Write-Host "    - SUCCESS: Packet Coalescing Filter set to Disabled (was '$($CurrentOffloadSettings.PacketCoalescingFilter)')." -ForegroundColor Green
            $script:ChangesMade = $true
        } else {
            Write-Host "    - INFO: Packet Coalescing Filter already Disabled. No change needed." -ForegroundColor Gray
        }
    } else {
        Write-Warning "    * INFO/FAILURE: Could not retrieve current NetOffloadGlobalSettings. Skipping checks and attempts to set."
        # $script:ScriptEncounteredErrors = $true # Uncomment if this should be treated as a hard error
    }
} catch { 
    Write-Warning "    * FAILURE: Could not execute or query Set-NetOffloadGlobalSetting cmdlets. Error: $($_.Exception.Message)"
    $script:ScriptEncounteredErrors = $true
}

# Per-Adapter Settings (PowerShell Cmdlet) - RSC Only
Write-Host "  [CMDLET] Per-Adapter Receive Segment Coalescing (RSC) Setting:"
$RscAdapterChangeMadeOrAttempted = $false
try {
    $Adapters = Get-NetAdapter -ErrorAction SilentlyContinue
    if ($null -eq $Adapters -or $Adapters.Count -eq 0) {
        Write-Warning "    * INFO: No network adapters found to apply per-adapter RSC settings."
    } else {
        foreach ($Adapter in $Adapters) {
            $AdapterName = $Adapter.Name
            if ($Adapter.Status -eq 'Up') {
                try {
                    $CurrentAdapterRsc = Get-NetAdapterRsc -Name $AdapterName -ErrorAction SilentlyContinue
                    if ($CurrentAdapterRsc) {
                        if ($CurrentAdapterRsc.IPv4Enabled -eq $true -or $CurrentAdapterRsc.IPv6Enabled -eq $true) {
                            Disable-NetAdapterRsc -Name $AdapterName -Confirm:$false -ErrorAction Stop | Out-Null
                            Write-Host "    - SUCCESS: Attempted to disable RSC on active adapter: '$AdapterName' (was IPv4:$($CurrentAdapterRsc.IPv4Enabled)/IPv6:$($CurrentAdapterRsc.IPv6Enabled))." -ForegroundColor Green
                            $script:ChangesMade = $true
                            $RscAdapterChangeMadeOrAttempted = $true
                        } else {
                             Write-Host "    - INFO: RSC already disabled (IPv4:$($CurrentAdapterRsc.IPv4Enabled)/IPv6:$($CurrentAdapterRsc.IPv6Enabled)) for active adapter: '$AdapterName'. No change needed." -ForegroundColor Gray
                        }
                    } else { 
                        Write-Host "    - INFO: Adapter '$AdapterName' does not appear to support RSC control via Get/Set-NetAdapterRsc. Skipping." -ForegroundColor Gray
                    }
                } catch { # Catches errors from Disable-NetAdapterRsc or issues within the inner try
                    Write-Warning "    * INFO/FAILURE: Could not query or disable RSC on adapter '$AdapterName'. Error: $($_.Exception.Message)"
                    # Decide if this should set $script:ScriptEncounteredErrors = $true
                }
            } else {
                Write-Host "    - INFO: Skipping RSC check/disable for adapter '$AdapterName' as it is not 'Up'." -ForegroundColor Gray
            }
        }
         if (-not $RscAdapterChangeMadeOrAttempted -and $Adapters.Count -gt 0) {
            Write-Host "    - INFO: Per-adapter RSC settings appear optimal or adapters do not support this control method." -ForegroundColor Gray
        }
    }
} catch { # Catches errors from Get-NetAdapter itself
    Write-Warning "    * FAILURE: General error during per-adapter RSC processing (e.g., Get-NetAdapter failed). Error: $($_.Exception.Message)"
    $script:ScriptEncounteredErrors = $true 
}

# TCP Settings Profile (PowerShell Cmdlet - InternetCustom)
Write-Host "  [CMDLET] TCP Settings Profile Tweaks (for 'InternetCustom' profile):"
$TcpTemplate = "InternetCustom"
try {
    $CurrentTcpSettings = Get-NetTCPSetting -SettingName $TcpTemplate -ErrorAction SilentlyContinue
    if ($CurrentTcpSettings) {
        $ChangesMadeThisTcpTemplateBlock = $false
        if ($CurrentTcpSettings.MinRto -ne 300) {
            Set-NetTCPSetting -SettingName $TcpTemplate -MinRto 300 | Out-Null
            Write-Host "    - SUCCESS: MinRto for '$TcpTemplate' set to 300 (was '$($CurrentTcpSettings.MinRto)')." -ForegroundColor Green
            $script:ChangesMade = $true; $ChangesMadeThisTcpTemplateBlock = $true
        } else { Write-Host "    - INFO: MinRto for '$TcpTemplate' already 300. No change needed." -ForegroundColor Gray }

        if ($CurrentTcpSettings.InitialCongestionWindow -ne 10) {
            Set-NetTCPSetting -SettingName $TcpTemplate -InitialCongestionWindow 10 | Out-Null
            Write-Host "    - SUCCESS: InitialCongestionWindow for '$TcpTemplate' set to 10 (was '$($CurrentTcpSettings.InitialCongestionWindow)')." -ForegroundColor Green
            $script:ChangesMade = $true; $ChangesMadeThisTcpTemplateBlock = $true
        } else { Write-Host "    - INFO: InitialCongestionWindow for '$TcpTemplate' already 10. No change needed." -ForegroundColor Gray }

        if ($CurrentTcpSettings.AutoTuningLevelLocal -ne "Normal" -or $CurrentTcpSettings.ScalingHeuristics -ne "Disabled") {
            Set-NetTCPSetting -SettingName $TcpTemplate -AutoTuningLevelLocal Normal -ScalingHeuristics Disabled | Out-Null
            Write-Host "    - SUCCESS: AutoTuningLevelLocal:Normal, ScalingHeuristics:Disabled for '$TcpTemplate' applied." -ForegroundColor Green
            Write-Host "               (was AutoTune:'$($CurrentTcpSettings.AutoTuningLevelLocal)', Heuristics:'$($CurrentTcpSettings.ScalingHeuristics)')"
            $script:ChangesMade = $true; $ChangesMadeThisTcpTemplateBlock = $true
        } else { Write-Host "    - INFO: AutoTuningLevelLocal (Normal) & ScalingHeuristics (Disabled) for '$TcpTemplate' already set. No change needed." -ForegroundColor Gray }
        
        if (-not $ChangesMadeThisTcpTemplateBlock) {
             Write-Host "    - INFO: All '$TcpTemplate' profile settings were already optimal." -ForegroundColor Gray
        }
    } else {
        Write-Warning "    * INFO: TCP template '$TcpTemplate' not found. Skipping these settings. (This is okay if you haven't used tools like TCP Optimizer that create this profile)."
    }
} catch { 
    Write-Warning "    * FAILURE: Could not apply or query NetTCPSetting for profile '$TcpTemplate'. Error: $($_.Exception.Message)"
    $script:ScriptEncounteredErrors = $true
}

Write-Host "------------------------------------------------------------"
Write-Host "System-level network optimizations script finished." -ForegroundColor Green

if ($ScriptEncounteredErrors) {
    Write-Host ""
    Write-Host "ATTENTION: One or more settings could not be applied or an error/warning occurred. Please review messages above." -ForegroundColor Yellow
} elseif (-not $ChangesMade) {
    Write-Host ""
    Write-Host "All checked settings were already optimally configured. No changes were made by this script." -ForegroundColor Cyan
} else { # $ChangesMade is true and $ScriptEncounteredErrors is false
    Write-Host ""
    Write-Host "All applicable settings were checked and necessary changes applied successfully." -ForegroundColor Cyan
    Write-Host "A system RESTART is recommended for all changes to take full effect." -ForegroundColor Yellow
}

# If there were errors, a restart is still likely a good idea if any changes *were* made before an error.
# The message for $ScriptEncounteredErrors already implies to check, and then the restart message appears.
# If changes were made AND errors occurred, the yellow ATTENTION message will be shown, followed by the restart recommendation.

if ($ChangesMade -and -not $ScriptEncounteredErrors) {
    # Covered by the else block above
} elseif (-not $ChangesMade -and -not $ScriptEncounteredErrors) {
    # Covered by the elseif block above, restart message is still shown but context is different
}

Write-Host "------------------------------------------------------------"
Read-Host "Press ENTER to exit..."