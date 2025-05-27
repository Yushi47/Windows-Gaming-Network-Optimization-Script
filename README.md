# Windows Gaming Network Optimization Scripts

This repository contains a Python launcher and a PowerShell script designed to apply various system-level network optimizations on Windows, aimed at improving performance for gaming and reducing latency.

**Philosophy:**
* **Targeted Tweaks:** Focuses on well-known optimizations that can benefit latency-sensitive applications.
* **User Control:** The Python launcher allows interactive selection of network adapters for interface-specific tweaks.
* **No "Bloat":** Avoids obscure or unnecessary changes. Does not modify the TCP Congestion Control Provider (uses system default, typically CUBIC).
* **Complements Manual NIC Settings:** This script handles system-level and TCP/IP stack settings. It assumes the user will manually configure detailed advanced properties of their Network Interface Card (NIC) driver (e.g., LSO, Checksum Offloads, RSS queues, EEE, Flow Control) via Device Manager. This README provides guidelines for those manual NIC settings.
* **Transparency:** The PowerShell script provides detailed output on checks performed and actions taken.

## Features

* **Python Launcher (`Run_GamingNetworkOptimization_Admin.py`):**
    * Automatically detects available network adapters and their GUIDs.
    * Provides an interactive command-line interface to select one or more network adapters for targeted interface-specific TCP tweaks.
    * Launches the PowerShell optimization script with the necessary administrator privileges.
* **PowerShell Optimization Script (`GamingNetworkOptimization.ps1`):**
    * Applies a range of system-level network configuration changes.
    * Includes pre-execution checks: verifies if settings are already optimal before applying changes.
    * Provides detailed console output for each action (success, failure, or already set).
    * Keeps the console window open after execution for review, prompting the user to press Enter to exit.
    * Recommends a system restart if changes were made.

## Prerequisites

* Windows Operating System (tested on Windows 10/11, should work on others).
* Python 3.x installed (for the `Run_GamingNetworkOptimization_Admin.py` launcher).
* PowerShell (available by default on modern Windows).

## How to Use

1.  **Download:** Download both `Run_GamingNetworkOptimization_Admin.py` and `GamingNetworkOptimization.ps1` into the **same directory**.
2.  **Configure Manual NIC Settings (Recommended First):** Before running the script for the first time, or after any NIC driver update, it's advisable to configure your NIC's Advanced Properties in Device Manager. See the "Recommended Manual NIC Advanced Property Settings" section below for guidelines.
3.  **Run Python Launcher:** Execute the Python script (`Run_GamingNetworkOptimization_Admin.py`). You can usually do this by double-clicking it or running `python Run_GamingNetworkOptimization_Admin.py` from a command prompt in that directory.
4.  **Select Network Adapters:** The Python script will list your network adapters. Follow the prompts to enter the number(s) corresponding to the physical network adapter(s) (e.g., your main Ethernet or Wi-Fi) you want the PowerShell script to apply its interface-specific TCP tweaks to. You can select multiple adapters or choose to skip this step.
    * *Note: It is generally recommended to apply the script's interface-specific tweaks only to your primary physical gaming adapters and usually NOT to virtual adapters (like VPNs) unless you are sure.*
5.  **Administrator Privileges (UAC):** The Python script will then attempt to launch the `GamingNetworkOptimization.ps1` script as an administrator. If User Account Control (UAC) is enabled, you will see a prompt asking for permission. Click "Yes" to allow it.
6.  **Review PowerShell Output:** A new PowerShell window will open and execute the optimization script. Pay attention to the output messages. It will tell you what settings were checked, if they were already optimal, or if they were changed. It will also report any errors.
7.  **Restart (Recommended):** After the PowerShell script finishes, it will recommend a system restart if any settings were changed by the script. It's highly advised to restart your computer for all tweaks (script-based and manual NIC settings) to take full effect.
8.  **Press Enter:** Both the Python script window (if run from a console) and the PowerShell script window will wait for you to press Enter before closing.

## Optimizations Applied by `GamingNetworkOptimization.ps1`

The PowerShell script targets the following areas (among others), only applying changes if the current settings are not already optimal:

* **Interface-Specific TCP Settings (for selected NICs via Python launcher):**
    * `TcpAckFrequency` (Immediate ACKs): Enabled
    * `TCPNoDelay` (Nagle's Algorithm): Disabled for the interface
    * `TcpDelAckTicks` (Delayed ACK Timeout): Disabled
* **Global TCP/IP Stack Parameters (via `netsh`):**
    * Direct Cache Access (DCA): Attempted enable
    * Receive-Side Scaling (RSS) Global State: Enabled
    * Receive Segment Coalescing (RSC) Global State: Disabled (Key for low latency)
    * RFC 1323 Timestamps: Disabled
    * Initial Retransmission Timeout (InitialRTO): Set to 2000ms
    * Non SACK RTT Resiliency: Disabled
    * Max SYN Retransmissions: Set to 2
* **System-Wide Registry Tweaks:**
    * MSMQ `TCPNoDelay`: Enabled
    * Network Throttling Index: Disabled (`ffffffff`)
    * System Responsiveness: Prioritized for foreground applications (`0`)
    * Host Resolution Priorities (Local, Hosts, DNS, NetBT)
    * Max User Port: Set to 65534
    * TCP Timed Wait Delay: Set to 30 seconds
    * Default TTL: Set to 64
    * QoS `NonBestEffortLimit`: Disabled (`0`)
    * QoS `Do not use NLA`: Enabled (`1`)
    * LargeSystemCache (Memory Management): Disabled (`0` - favors application memory)
* **Global Offload Settings (via PowerShell cmdlets):**
    * TCP Chimney Offload: Disabled
    * Packet Coalescing Filter: Disabled
* **Per-Adapter RSC Setting (supplementary):**
    * Attempts to disable RSC on active network adapters if supported by the driver via PowerShell cmdlets.
* **TCP Settings Profile ("InternetCustom"):**
    * Applies specific parameters like `MinRto`, `InitialCongestionWindow`, `AutoTuningLevelLocal Normal`, `ScalingHeuristics Disabled`. (If the "InternetCustom" profile exists, often created by tools like TCP Optimizer).

## Recommended Manual NIC Advanced Property Settings (via Device Manager)

The PowerShell script handles system-level settings. For optimal performance, it's crucial to also configure your primary Network Interface Card's (NIC) advanced driver properties directly through **Device Manager**.
To access these: Right-click Start -> Device Manager -> Expand "Network adapters" -> Right-click your main network adapter (e.g., "Realtek PCIe GbE Family Controller") -> Properties -> Advanced tab.

The exact names and availability of these settings can vary significantly between network card manufacturers and driver versions. Apply the following recommended settings **if they are available for your adapter**:

* **Energy Efficiency Related:**
    * **Advanced EEE / Energy Efficient Ethernet / EEE Max Support Speed:** `Disabled`
        * *Benefit:* These power-saving features can introduce latency. Disabling them ensures the NIC is always at full performance.
    * **Enable Green Ethernet / Green Ethernet:** `Disabled`
        * *Benefit:* Similar to EEE, aims to save power but can impact performance and latency.
    * **Auto Disable Gigabit / Power Saving Mode:** `Disabled`
        * *Benefit:* Prevents the NIC from downshifting speed or performance to save power.
    * **GigaLite:** `Disabled`
        * *Benefit:* Prevents a potential speed limit (e.g., to 500Mbps) on a gigabit link for power saving.

* **Flow Control & Interrupts:**
    * **Flow Control:** `Disabled` (or `Rx & Tx Disabled`)
        * *Benefit:* Prevents potential pauses in transmission, which can cause lag spikes in gaming.
    * **Interrupt Moderation / Interrupt Moderation Rate:** `Disabled`
        * *Benefit:* Forces the NIC to interrupt the CPU more frequently for new packets, reducing packet processing latency. Increases CPU usage but is often best for the lowest latency.

* **Offloads (Configure based on your preference - Disabled for lowest potential latency at cost of CPU):**
    * **IPv4 Checksum Offload / TCP Checksum Offload (IPv4/IPv6) / UDP Checksum Offload (IPv4/IPv6):** `Disabled`
        * *Benefit of Disabling (Your Preference):* CPU handles checksums, potentially reducing NIC processing pipeline latency. Increases CPU load. (Default is often `Rx & Tx Enabled` to save CPU).
    * **Large Send Offload (LSO) (V1/V2, IPv4/IPv6):** `Disabled`
        * *Benefit of Disabling (Your Preference):* CPU handles TCP segmentation, potentially reducing NIC processing latency. Increases CPU load. (Default is often `Enabled` to save CPU).

* **Performance & Throughput:**
    * **Receive Side Scaling (RSS):** `Enabled`
        * *Benefit:* Distributes network receive processing across multiple CPU cores. (The script enables the OS global state; ensure it's enabled here in the driver too).
    * **Number of RSS Queues / \*NumRssQueues:** Typically `2` or `4`, or matching available CPU cores (e.g., `4 Queues`).
        * *Benefit:* Optimizes processing distribution for RSS.
    * **Receive Buffers / \*ReceiveBuffers:** (e.g., `512` or as tested)
        * *Benefit:* Memory for incoming packets. Balance between preventing drops (higher) and reducing latency (not excessively high).
    * **Transmit Buffers / \*TransmitBuffers:** (e.g., `128` or as tested)
        * *Benefit:* Memory for outgoing packets. Similar balance to receive buffers.
    * **Jumbo Packet / Jumbo Frame:** `Disabled` (or standard MTU like 1500 Bytes)
        * *Benefit:* Not beneficial for internet gaming; ensures compatibility.

* **Connectivity & Wake Features:**
    * **Speed & Duplex:** `Auto Negotiation` (usually best) or manually set to your network's highest (e.g., `1.0 Gbps Full Duplex`).
        * *Benefit:* Ensures maximum link performance. Manual setting can sometimes resolve negotiation issues.
    * **Wake on Magic Packet / Wake on Pattern Match / S5WakeOnLan / \*ModernStandbyWoLMagicPacket:** `Disabled` (unless you specifically use Wake-on-LAN).
        * *Benefit:* Prevents unexpected system wake-ups.
    * **WolShutdownLinkSpeed:** Setting is less relevant if WoL is disabled, but `10 Mbps First` is a common default.

* **Miscellaneous:**
    * **\*PriorityVLANTag / Priority & VLAN:** `Priority & VLAN Disabled`
        * *Benefit:* Not needed for typical home gaming networks.
    * **\*PMARPOffload / \*PMNSOffload (ARP/NS Offload for power management):** `Disabled`
        * *Benefit:* Ensures OS handles these actively, avoiding potential power-save related delays.
    * **Network Address:** Leave blank/default unless you need to spoof your MAC address.
    * **RegVlanid:** `0` or default (no VLAN).

**Important Note on Manual NIC Settings:** The names of these properties can differ greatly. If a setting listed above is not present for your adapter, you cannot configure it. Always refer to your NIC manufacturer's documentation if unsure.

## Important Notes & Disclaimer (General)

* **Administrator Privileges:** The `GamingNetworkOptimization.ps1` script **must** be run as an administrator to modify system settings. The Python launcher handles this elevation request.
* **Understand the Changes:** It is recommended to understand what each tweak does. While these are generally considered safe and beneficial for gaming, every system is different.
* **Backup/System Restore Point:** Before running any script that modifies system settings, it's always a good practice to create a system restore point.
* **No Revert Script:** This package does not currently include an automatic "revert to defaults" script. Most settings can be reverted using tools like TCP Optimizer (which has a "Windows Defaults" option) or by manually changing registry values back.
* **Use at Your Own Risk:** While these settings are commonly recommended for performance, the author(s) of these scripts are not responsible for any issues that may arise from their use.

## Troubleshooting

* **PowerShell script window closes immediately:** Ensure you are running it via the Python launcher, which includes a pause at the end of the PowerShell script's execution. If running the `.ps1` file directly in an admin PowerShell console, the console will remain open.
* **"Script Not Found":** Make sure `GamingNetworkOptimization.ps1` is in the same folder as `Run_GamingNetworkOptimization_Admin.py`.
* **Settings not sticking (especially RSC globally):** If a setting like the global RSC state reverts after a reboot despite the script indicating success, an external factor (another application, a driver behavior on boot, or a system policy) might be overriding it.
* **"Access Denied" or UAC prompt denied:** The PowerShell script needs admin rights. Ensure you approve the UAC prompt.

---