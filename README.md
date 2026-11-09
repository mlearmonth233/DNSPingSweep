# DNSPingSweep: NetTools for Excel

DNS resolution, ping, traceroute, subnet scan and subnet calculator tool. An Excel add-in that adds
a **NetTools** tab to the ribbon:

| Button | What it does |
|---|---|
| **Check DNS** | For each IP address or hostname: forward lookup, reverse (PTR) lookup, a forward/reverse match check, and 4 pings |
| **DNS Only** | The lookups without the ping (faster) |
| **Ping Only** | 4 pings per host: status, replies, average latency, TTL |
| **Traceroute** | Hop-by-hop path to each host (3 probes per hop, like `tracert`) |
| **Subnet Scan** | Pings every address in a subnet and looks up its reverse DNS (PTR) name |
| **Subnet Calc** | Network, mask, wildcard, first/last usable, broadcast and host count for each subnet |
| **About** | Short help |

**Results always go on a new sheet as a formatted Excel table with headers.
The cells you select are never changed or overwritten.** Each run adds its
own sheet (`DNS Check`, `DNS Lookup`, `Ping`, `Subnets`, `Trace server1`,
`Scan 10.1.4.0_24`)
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
| Replies | `4/4` |
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

### Subnet Scan results

Type or select a subnet as `10.1.4.0/24` or `10.1.4.0 255.255.255.0` (any
address inside it works too). NetTools pings every usable address (500 ms timeout), pings it once more
if the first gets no reply, and looks up its PTR record straight from DNS. Each
subnet gets its own sheet, e.g. `Scan 10.1.4.0_24`, with a table of:

| Column | Example |
|---|---|
| IP Address | `10.1.4.17` (green if it answered) |
| Ping | `Up`, `No reply`, `Destination host unreachable` |
| Latency (ms) | `2` |
| TTL | `255` |
| Reverse DNS (PTR) | `ap-floor2-03.corp.example.com` |
| Fwd/Rev Match | `Yes` if the PTR name resolves back to this IP |

A summary line above the table, and a pop-up at the end, give the count of
addresses that are up and that have PTR records. Use the table's filter
arrows to show only `Up` addresses or only ones with a PTR. Up to 4,096
addresses (a /20) per subnet. NetTools asks before anything bigger than a
/24. A /24 where nothing answers takes about 4 minutes. Esc stops the scan
and keeps what's been scanned so far.

### Subnet Calc results

Accepts `10.1.4.0/22`, `10.1.4.7/22` (any address in the subnet),
`10.1.4.7 255.255.252.0` or `10.1.4.7/255.255.252.0`. Columns:
Subnet (as entered), Network/Prefix, Network Address, Subnet Mask,
Wildcard Mask, First Usable, Last Usable, Broadcast Address, Usable Hosts,
Total Addresses. /31 counts both addresses as usable (point-to-point links)
and /32 is a single host. IPv4 only.

### Speed

Lookups go through the Windows resolver, so the answers match what the PC
itself sees (DNS suffixes, hosts file and so on). Ping uses WMI, so no
command windows pop up. Each host gets 4 pings (`PING_COUNT`), and each
ping waits up to 1 second (`PING_TIMEOUT_MS`), both set at the top of
`src/NetTools.bas`. Hosts are checked one at a time: a host that is up
takes a fraction of a second, a host that is down about 4 seconds.

### Worksheet functions

Also available in formulas, if you want results next to your own data:

```
=NT_RESOLVE(A2)                 10.1.2.3
=NT_PTR(A2)                     server1.corp.example.com
=NT_PING(A2)                    Up (4/4, 3 ms)
=NT_SUBNET(A2)                  10.1.4.0/22 (10.1.4.1 - 10.1.7.254, 1022 hosts)
=NT_SUBNET(A2, "mask")          255.255.252.0
=NT_INSUBNET(B2, A2)            TRUE if the IP in B2 is inside subnet A2
```

`NT_SUBNET` fields: `cidr`, `network`, `mask`, `wildcard`, `first`, `last`,
`broadcast`, `hosts`, `range`, `prefix`.

## Install

### MSI (recommended)

Double-click `NetTools-2.2.1.msi` and close Excel first if it's open. No admin
rights are needed. The installer puts `NetTools.xlam` in your Excel startup
folder (`%APPDATA%\Microsoft\Excel\XLSTART`). Excel loads it every time it
starts, and because that folder is one of Office's default trusted
locations, the macros run without a security prompt.

- Uninstall: **Settings > Apps > NetTools for Excel**.
- Upgrade: run a newer MSI. It replaces the old version.
- Silent / Intune / SCCM: `msiexec /i NetTools-2.2.1.msi /qn`. This is a
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
- `-Msi` uses the [WiX Toolset](https://wixtoolset.org/) v5. If WiX isn't
  installed, the script installs it as a .NET global tool, which needs the
  [.NET SDK](https://dotnet.microsoft.com/download) 6 or later.
- The version comes from `VERSION`. Bump it before you build an upgrade,
  and keep `APP_VERSION` in `src/NetTools.bas` the same.

### Files

| File | Purpose |
|---|---|
| `src/NetTools.bas` | All the VBA: ribbon callbacks, DNS (ws2_32), ping (WMI), traceroute + subnet scan (ICMP API, DnsQuery), subnet maths, worksheet functions |
| `src/customUI.xml` | Ribbon tab definition |
| `installer/NetTools.wxs` | WiX source for the per-user MSI |
| `build.ps1` | Builds the `.xlam` and, with `-Msi`, the MSI |
| `VERSION` | Version used for the add-in and the MSI |

To change the ribbon, edit `src/customUI.xml`. Each `onAction` names a
`Public Sub ...(control As IRibbonControl)` in `NetTools.bas`.
