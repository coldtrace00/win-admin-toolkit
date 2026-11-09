# Win Admin Toolkit

[Українська](README.md) · **English**

A PowerShell toolkit for Windows admins in small offices: PC inventory, PC and network diagnostics, new PC setup, network security audit and Active Directory reports.

- **One menu** for every tool — `Admin-Menu.bat`.
- **The result is an HTML report** on the desktop: search, sorting, CSV export, light and dark theme, print-friendly.
- **Nothing to install**: runs on the built-in PowerShell 5.1 in Windows 10/11.
- **Plain language**: every issue found comes with an explanation and what to do about it.

> The tools' interface and reports are in Ukrainian.

## Tools

| Tool | What it does | Admin rights |
|---|---|---|
| [PC Inventory](#pc-inventory) | Hardware, Windows, serial numbers, network, installed software | Not needed |
| [PC Health](#pc-health) | Why a PC is slow or unstable and what to do about it | Recommended, the script asks itself |
| [New PC Setup](#new-pc-setup) | Silent app install via winget and baseline Windows settings | Required, the script asks itself |
| [Network Diagnostics](#network-diagnostics) | Finds step by step where exactly the network breaks | Not needed |
| [Network Scanner](#network-scanner) | Devices on the network, open ports, risks and how to close them | Not needed |
| [Active Directory Reports](#active-directory-reports) | Stale accounts, passwords, lockouts, admins, OS versions | Regular domain user |

## Quick start

1. Click **Code → Download ZIP** and extract it to any folder.
2. Unblock the files — Windows blocks scripts downloaded from the internet. Open PowerShell in the extracted folder and run:
   ```powershell
   Get-ChildItem -Recurse | Unblock-File
   ```
3. Run **`Admin-Menu.bat`** and pick a tool by number.

Each tool can also be run on its own — via the `Run-*.bat` in its folder or directly:

```powershell
powershell -ExecutionPolicy Bypass -File .\Network-Diag\Test-Network.ps1 -Quick
```

## PC Inventory

`PC-Inventory\Get-PCInventory.ps1`

Collects a full PC profile into one HTML file:

- model, serial number (from BIOS or motherboard), BIOS, domain or workgroup, user;
- CPU, memory modules per slot, physical disks with type (NVMe / SSD / HDD) and health, volume usage;
- video adapters with driver versions, network adapters with IP, MAC, gateway and DNS;
- installed software with version, publisher and install date — searchable.

Handy for asset tracking and before replacing or repairing a PC.

## PC Health

`PC-Health\Test-PCHealth.ps1`

Shows in about a minute why a PC is slow or unstable and what to do about it. Checks:

- uptime and updates waiting for a reboot;
- disk space, disk health (SMART), system drive on an HDD;
- what is eating memory and CPU;
- blue screens, unexpected shutdowns, application crashes, disk errors in event logs;
- how old Windows updates are, antivirus, devices without drivers, key services;
- startup items, temp files, network, laptop battery, power plan.

Check this PC or another one by name over WinRM (requires a domain and admin rights on that PC):

```powershell
.\PC-Health\Test-PCHealth.ps1 -ComputerName BUH-03
```

On the local PC it offers quick fixes afterwards: clean temp files older than a day and start a stopped print spooler.

## New PC Setup

`PC-Setup\Setup-NewPC.ps1`

Sets up a workstation without manual clicking:

- checks for winget and installs it if missing (common on Windows 10);
- shows an app menu: Chrome, AnyDesk, 7-Zip, Acrobat Reader, Notepad++, VLC, Telegram, Zoom and more. Skips apps that are already installed;
- installs per-user apps as the person sitting at the PC, not as the admin;
- creates a restore point before making changes;
- sets the time zone and power plan, shows file extensions, configures Explorer, turns off ads and "suggested" apps, optionally enables RDP with NLA;
- applies Explorer settings to the default profile too — i.e. to every future user;
- renames the PC and keeps a log in `C:\ProgramData\PC-Setup`.

For bulk deployment without the menu:

```powershell
.\PC-Setup\Setup-NewPC.ps1 -Auto -ComputerName BUH-03
```

The app list and settings live at the top of the file (the `НАЛАШТУВАННЯ` section).

## Network Diagnostics

`Network-Diag\Test-Network.ps1`

Walks the chain step by step and tells you where exactly the problem is:

**adapter → IP address (DHCP) → gateway → internet → DNS → websites (HTTP / HTTPS)**

Also:

- ping with packet loss and jitter, quality for VoIP calls;
- traceroute and a download speed test;
- detects a 169.254 address (no DHCP response), a guest Wi-Fi login page, a wrong PC clock;
- proxy, hosts file, "Public" network profile;
- reachability of the domain controller and your own servers by port.

Quick mode without traceroute and speed test (~20 s): `-Quick`. Add your own servers to check in the `НАЛАШТУВАННЯ` section at the top of the file.

## Network Scanner

`Network-Scan\Scan-Network.ps1`

A security-focused scanner: what is on the network, what is open and how to close it.

- finds devices via ping, ARP, TCP and NetBIOS — including those that block ping;
- identifies name, MAC, vendor and device type: PC, printer, camera, router, NAS;
- checks ~35 important ports (RDP, SMB, Telnet, VNC, databases, cameras, MikroTik Winbox and more) and explains the risk and the fix for each;
- shows what is visible from the internet (Shodan InternetDB) and what the router has forwarded via UPnP;
- checks this PC itself: RDP, NLA, firewall, SMBv1, listening ports;
- cross-checks with Active Directory: devices outside the domain and domain PCs that don't respond;
- remembers previous scans and shows new and missing devices, plus free IP addresses.

```powershell
.\Network-Scan\Scan-Network.ps1 -Subnet 192.168.10.0/24 -NoExternal
```

Without `-Subnet` it scans the networks of all active adapters, VPN included. `-NoExternal` turns off the check of what is visible from the internet.

> Only scan networks you manage or have permission to scan.

## Active Directory Reports

`AD-Reports\Get-ADReports.ps1` · **beta**

10 reports from a menu, one at a time or all in one file:

1. Inactive users
2. Inactive computers
3. Passwords expiring soon
4. Locked-out accounts
5. Groups and their members
6. Users and their groups
7. Privileged groups (nested groups expanded)
8. Passwords that never expire
9. Disabled accounts
10. All computers and Windows versions (outdated OS flagged)

Plus the domain password policy and an account unlock menu with confirmation.

```powershell
.\AD-Reports\Get-ADReports.ps1 -Report All -InactiveDays 60
```

Requires the ActiveDirectory module (RSAT) — if it's missing, the script offers to install it. Unlocking accounts requires admin or helpdesk rights.

## Requirements

- Windows 10 or 11 with the built-in Windows PowerShell 5.1.
- AD reports: a domain-joined PC and RSAT.
- Remote checks in PC Health: a domain and WinRM enabled on the target PC.

## Where data is stored

| What | Where |
|---|---|
| HTML reports | Current user's desktop |
| New PC setup log | `C:\ProgramData\PC-Setup` |
| Scan history and vendor database | `C:\ProgramData\NetScan` |

External services the tools contact:

- **Network Diagnostics**: `msftconnecttest.com`, `google.com`, `api.ipify.org`, `speed.cloudflare.com`.
- **Network Scanner**: `api.ipify.org` and `internetdb.shodan.io` — the internet exposure check, turned off with `-NoExternal`; `standards-oui.ieee.org` — the NIC vendor database, downloaded once every 90 days.
- **New PC Setup**: downloads apps via winget from their official sources.

Apart from these requests, the tools send nothing anywhere: reports and history stay on your PC.

## License

[MIT](LICENSE)
