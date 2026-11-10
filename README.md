# DNSPingSweep: NetTools for Excel

DNS resolution, ping, port check, traceroute, subnet scan, subnet calculator
MAC address and SSH log collection tool. An Excel add-in that adds
a **NetTools** tab to the ribbon:

| Button | What it does |
|---|---|
| **Check DNS** | For each IP address or hostname: forward lookup, reverse (PTR) lookup, a forward/reverse match check, and up to 4 pings |
| **DNS Only** | The lookups without the ping (faster) |
| **Ping Only** | Up to 4 pings per host: status, replies, average latency, TTL |
| **Port Check** | Is TCP 22 / 443 / 3389 / ... open on each host? |
| **Traceroute** | Hop-by-hop path to each host (3 probes per hop, like `tracert`) |
| **Subnet Scan** | Every address in a subnet: ping, MAC address + vendor, reverse DNS, and optionally TCP ports. Hundreds at once |
| **Subnet Calc** | Network, mask, wildcard, first/last usable, broadcast and host count for each subnet |
| **MAC Format** | Reformats MAC addresses as `xxxx.xxxx.xxxx` or `XX:XX:XX:XX:XX:XX` (you choose) and shows the vendor |
| **SSH Logs** | SSH into Cisco IOS / IOS-XE / 9800 or AireOS devices with your username and password, run commands such as `show logging`, one sheet per device |
| **About** | Short help |

**Results always go on a new sheet as a formatted Excel table with headers.
The cells you select are never changed or overwritten.** Each run adds its
own sheet (`DNS Check`, `DNS Lookup`, `Ping`, `Subnets`, `Trace server1`,
`Scan 10.1.4.0_24`, `Port Check`, `MAC Addresses`, `SSH Summary`, `SSH <device>`)
with a title, the time it ran, and a table you can sort and filter.
Running the same button again **overwrites** its sheet with the new
results. NetTools only overwrites sheets it created: if one of your own
sheets already has that name, the results go on `Subnets 2` (etc.) instead.

## Using it

1. Put IP addresses, hostnames or subnets in a column. A header such as
   `Hostname`, `IP Address` or `Subnet` is fine; NetTools spots it and skips it.
2. Select the column, a range in it, or just its first cell (NetTools then
   runs down to the last value in the column).
   Or select an **empty cell** and NetTools asks you to type them in
   (comma separated).
3. Click a button. The new results sheet opens and fills in as it goes.
   Press **Esc** to stop a long run; what's done so far is kept.

### Check DNS results

| Column | Example |
|---|---|
| Host / IP | what was checked |
| Input Type | `Hostname` or `IP` |
| Resolved IP(s) | `10.1.2.3, fe80::1` (IPv4 first) or `(not found)` |
| Reverse DNS (PTR) | `sw-core-01.corp.example.com` or `(no PTR record)` |
| Fwd/Rev Match | `Yes`, `No PTR`, or `No - ...` with the reason |
| Ping | `Up`, `Up - packet loss`, `Timed out`, `Destination host unreachable`, ... |
| Replies | `4/4`, or `0/2` for a host that missed 2 pings in a row and was marked down |
| Avg Latency (ms) | `3` |
| TTL | `254` |

Cells are coloured green (all good), yellow (packet loss / mismatch) or red
(down / not found).

**Fwd/Rev Match** means:
- for a **hostname**: the PTR record of its first address is the same name
  (`server1` matches `server1.corp.example.com`);
- for an **IP**: its PTR name resolves back to that same IP.

### Traceroute results

One sheet per host, with a table of: Hop, Address, Hostname (PTR),
Probe 1-3 (ms, `*` = no reply) and a Note. The row where the destination
answers is green. The trace stops at the destination, on an "unreachable"
reply, after 30 hops, or after 5 hops in a row with no reply (often a
firewall blocking ICMP). IPv4 only. Up to 20 hosts per run.

### Port Check results

Select hosts (or type them in), then enter the TCP ports, e.g. `22,80,443,3389`
or `8000-8010` (up to 64 ports; NetTools remembers your last list). The
`Port Check` sheet has one column per port:

- **Open**: the port accepted the connection (green).
- **Closed**: the host refused it. The host is up, but nothing is listening.
- **Filtered**: no answer within 2.5 seconds, usually a firewall dropping it
  (or nothing at that address).
- **Unreachable**: a router said the host or network can't be reached.

Plus an **Open Ports** column listing the open ones. Many connections are
tried at once, so a long list is still quick.

### Subnet Scan results

Type or select a subnet as `10.1.4.0/24` or `10.1.4.0 255.255.255.0` (any
address inside it works too). NetTools then asks for TCP ports to check on
every address. Leave it blank to skip the port check.

NetTools pings up to 256 addresses at once (1 second timeout) and pings
anything that didn't answer a second time. It then reads MAC addresses from
the PC's ARP table, looks up each PTR record straight from DNS, and runs the
port check if you asked for one. A /24 takes roughly 5-15 seconds without
ports.

Each subnet gets its own sheet, e.g. `Scan 10.1.4.0_24`:

| Column | Example |
|---|---|
| IP Address | `10.1.4.17`: green if it answered ping, yellow if it's there (MAC or open port) but blocks ping |
| Ping | `Up`, `No reply`, `Destination host unreachable` |
| Latency (ms) | `2` |
| TTL | `255` |
| MAC Address | `00:1b:54:aa:bb:cc`, in your chosen MAC format |
| Vendor | `Cisco Systems, Inc` / `(private / random MAC)` |
| Reverse DNS (PTR) | `ap-floor2-03.corp.example.com` |
| Fwd/Rev Match | `Yes` if the PTR name resolves back to this IP |
| TCP 22 SSH, TCP 443 HTTPS, ... | `Open` / `Closed` / `Filtered` (only if you entered ports) |
| Open Ports | `22, 443` |

**MAC addresses only appear for subnets your PC is directly connected to**
(same VLAN). That's how ARP works: for remote subnets the router answers
on the devices' behalf. They do appear for devices that block ping, which
helps you find devices hiding behind a firewall.

The summary line above the table and the pop-up at the end count the
devices found, the ones answering ping, the ones with a MAC, open ports and
PTR records. Use the filter arrows to show only what you need. Up to 4,096
addresses (a /20) per subnet; NetTools asks before anything bigger than a
/24. Esc stops the scan.

### MAC Format

Select a column of MAC addresses in any notation (`aabb.ccdd.eeff`,
`AA:BB:CC:DD:EE:FF`, `AA-BB-CC-DD-EE-FF`, `AABBCCDDEEFF`), click
**MAC Format**, and choose:

- `1`: `xxxx.xxxx.xxxx` (Cisco style, lower case), e.g. `aabb.ccdd.eeff`
- `2`: `XX:XX:XX:XX:XX:XX` (upper case), e.g. `AA:BB:CC:DD:EE:FF`

The `MAC Addresses` sheet lists each original, the formatted MAC and the
vendor. Your choice is remembered and also used for the MAC column in
Subnet Scan.

### SSH Logs

1. Select a column of device hostnames or IPs (or an empty cell to type them).
2. Click **SSH Logs** and choose the device type:
   - `1`: Cisco IOS / IOS-XE (switches, routers, Catalyst 9800 WLC)
   - `2`: Cisco AireOS WLC
3. Enter your username and password in the NetTools login box. The password
   is masked, used only for this run, and **never saved**.
4. Enter the commands, separated by semicolons, e.g. `show logging; show version`.
   NetTools remembers the list for each device type. `terminal length 0`
   (IOS) or `config paging disable` (AireOS) is sent first automatically.

NetTools connects to 8 devices at a time and then adds:
- an **SSH Summary** sheet listing each device with its status (`OK`, `Login
  failed - wrong username or password`, `Connection timed out`, `SSH refused
  (port 22 closed)`, `HOST KEY CHANGED`, ...), the number of lines, and a link
  to the device's sheet;
- one **SSH <device>** sheet per device with the full output, one line per row;
- a copy of every log in `Documents\NetTools\SSH Logs\<device> <date time>.txt`.

How it works and what to know:
- It uses PuTTY's `plink.exe`, which the MSI installs to `%APPDATA%\NetTools`.
  The build downloads the official copy and checks its digital signature.
  If PuTTY 0.77 or newer is already installed, NetTools can use that instead.
- The password reaches plink through a temporary file (`-pwfile`), not the
  command line, and the file is deleted as soon as the run ends. For
  AireOS the commands file also holds the login (AireOS asks for it again
  inside the session), and it's deleted the same way.
- **Host keys:** the first time you connect to a device, its SSH key is stored
  (in PuTTY's normal place). If a known device later shows a *different*
  key, NetTools refuses to connect and reports `HOST KEY CHANGED`. Check the
  device before removing the old key in PuTTY.
- Each device gets up to 2 minutes (`SSH_TIMEOUT_S`); Esc stops the run.

### MAC vendor list

Vendors come from the IEEE OUI registry. The build downloads it once
(about 35,000 vendors), and the MSI installs it as
`%APPDATA%\NetTools\oui.txt`. Lookups happen on your PC: nothing is sent
anywhere when you use the add-in. **About** shows whether the list is
installed. If you install the `.xlam` by hand, put `oui.txt` next to it.

### Subnet Calc results

Accepts `10.1.4.0/22`, `10.1.4.7/22` (any address in the subnet),
`10.1.4.7 255.255.252.0` or `10.1.4.7/255.255.252.0`. Columns:
Subnet (as entered), Network/Prefix, Network Address, Subnet Mask,
Wildcard Mask, First Usable, Last Usable, Broadcast Address, Usable Hosts,
Total Addresses. /31 counts both addresses as usable (point-to-point links)
and /32 is a single host. IPv4 only.

### Speed

Check DNS, DNS Only, Ping Only and Subnet Scan do their DNS **for the whole
list at once**:

- **Forward lookups** (name to IP) use the Windows asynchronous DNS call
  (`GetAddrInfoExW`), up to 128 at a time, DNS only. A name that doesn't
  exist fails in milliseconds instead of waiting seconds for LLMNR/NetBIOS.
  On Windows 7 (no asynchronous lookups) NetTools falls back to one at a time.
- **Reverse lookups** (PTR) are sent straight to your PC's DNS servers as
  ordinary DNS queries, up to 200 at a time, each with a 1.5 s timeout and
  one retry to the next DNS server. If no DNS server answers at all,
  NetTools falls back to one at a time.
- **Ping** happens after DNS, for all hosts together, in up to 4 rounds of
  1 second. A host that misses **2 pings in a row** counts as down.
- Results are written to the sheet in one go.

The pop-up at the end shows how long DNS and ping each took, e.g.
`Checked 250 host(s) in 6.1 s (DNS 1.9 s, ping 4.0 s)`. If DNS is still
slow, your DNS server is slow to answer, or reverse zones are missing and
queries for them time out.

Timeouts and batch sizes are constants at the top of `src/NetTools.bas`
(`LOOKUP_*`, `DNS_*`, `PING_*`).

### Worksheet functions

Also available in formulas, if you want results next to your own data:

```
=NT_RESOLVE(A2)                 10.1.2.3
=NT_PTR(A2)                     server1.corp.example.com
=NT_PING(A2)                    Up (4/4, 3 ms)
=NT_SUBNET(A2)                  10.1.4.0/22 (10.1.4.1 - 10.1.7.254, 1022 hosts)
=NT_SUBNET(A2, "mask")          255.255.252.0
=NT_INSUBNET(B2, A2)            TRUE if the IP in B2 is inside subnet A2
=NT_MAC(A2, "cisco")            aabb.ccdd.eeff   ("colon", "dash", "bare" too)
=NT_VENDOR(A2)                  Cisco Systems, Inc
```

`NT_SUBNET` fields: `cidr`, `network`, `mask`, `wildcard`, `first`, `last`,
`broadcast`, `hosts`, `range`, `prefix`.

## Install

### MSI (recommended)

Double-click `NetTools-3.3.0.msi` and close Excel first if it's open. No admin
rights are needed. The installer puts `NetTools.xlam` in your Excel startup
folder (`%APPDATA%\Microsoft\Excel\XLSTART`). Excel loads it every time it
starts, and because that folder is one of Office's default trusted
locations, the macros run without a security prompt.

- Uninstall: **Settings > Apps > NetTools for Excel**.
- Upgrade: run a newer MSI. It replaces the old version.
- Silent / Intune / SCCM: `msiexec /i NetTools-3.3.0.msi /qn`. This is a
  **per-user** MSI, so deploy it in the *user* context.

If your company uses Group Policy to turn off trusted locations or block
all macros, the ribbon tab will appear but its buttons won't work. Ask IT
to allow the add-in, for example by signing it with the company's
code-signing certificate.

### Manually (no MSI)

Copy `NetTools.xlam` to `%APPDATA%\Microsoft\AddIns`, then in Excel go to
**File > Options > Add-ins > Manage: Excel Add-ins > Go...**, click
**Browse**, and pick the file. If you downloaded the file, right-click it
first, choose **Properties**, and tick **Unblock**. Otherwise Windows marks
it as coming from the internet and Excel blocks its macros.

## Building

Download this repo (**Code > Download ZIP**, then right-click the ZIP >
**Properties** > **Unblock** before extracting) or `git clone` it, and run
these commands from the extracted folder. The `.xlam` holds compiled VBA, so you have to build it on **Windows with
desktop Excel installed**:

```powershell
powershell -ExecutionPolicy Bypass -File .\build.ps1          # dist\NetTools.xlam
powershell -ExecutionPolicy Bypass -File .\build.ps1 -Msi     # + dist\NetTools-<version>.msi
```

- The script drives Excel in the background. It switches on *Trust access to
  the VBA project object model* for the build and then puts the setting
  back the way it was.
- `-Msi` downloads the MAC vendor list from the IEEE (falling back to
  Wireshark's copy). If the PC is offline, the MSI still builds, but
  without vendor names.
- `-Msi` also downloads PuTTY's `plink.exe` (for SSH Logs) and checks it is
  signed by Simon Tatham, PuTTY's author.
- `-Msi` uses the [WiX Toolset](https://wixtoolset.org/) v5. If WiX isn't
  installed, the script installs it as a .NET global tool, which needs the
  [.NET SDK](https://dotnet.microsoft.com/download) 6 or later.
- The version comes from `VERSION`. Bump it before you build an upgrade,
  and keep `APP_VERSION` in `src/NetTools.bas` the same.

### Files

| File | Purpose |
|---|---|
| `src/NetTools.bas` | All the VBA: ribbon callbacks, DNS (ws2_32), ping (WMI), traceroute + subnet scan (ICMP API, ARP table, DnsQuery), port check (Winsock), MAC formatting + vendor lookup, subnet maths, worksheet functions |
| `src/NetToolsLogin.vba` | Code for the SSH login box (the build adds it as a UserForm) |
| `src/customUI.xml` | Ribbon tab definition |
| `installer/NetTools.wxs` | WiX source for the per-user MSI |
| `build.ps1` | Builds the `.xlam` and, with `-Msi`, the MSI |
| `VERSION` | Version used for the add-in and the MSI |

To change the ribbon, edit `src/customUI.xml`. Each `onAction` names a
`Public Sub ...(control As IRibbonControl)` in `NetTools.bas`.
