Attribute VB_Name = "NetTools"
'==============================================================================
' NetTools - Excel ribbon add-in for network engineers
'
' Check DNS : forward lookup, reverse (PTR) lookup, forward/reverse match
'             and ping for a column of IP addresses or hostnames.
' DNS Only  : the lookups without the ping.
' Ping Only : status, latency and TTL.
' Traceroute: hop-by-hop path to a host (3 probes per hop, like tracert).
' Subnet Scan: ping, MAC + vendor, reverse DNS (+ forward check) and optional
'             TCP ports for every address in a subnet - hundreds at once.
' Port Check: is TCP 22/443/3389/... open on each host?
' MAC Format: xxxx.xxxx.xxxx or XX:XX:XX:XX:XX:XX, plus the vendor.
'
' Results always go on a NEW sheet as an Excel table - the data you select
' is never changed.
' Subnet Calc: network, mask, wildcard, host range, broadcast and host
'             count for a column of subnets (10.1.4.0/22, 10.1.4.7 255.255.252.0).
'
' DNS uses the Windows resolver (ws2_32 getaddrinfo / getnameinfo), so it
' answers exactly as the PC would (hosts file, DNS suffixes, etc.).
' Ping uses WMI Win32_PingStatus - no console windows flash up.
' Traceroute uses the ICMP API (iphlpapi IcmpSendEcho) with a rising TTL.
'
' Worksheet functions are included too:
'   =NT_RESOLVE(A2)   -> IP address(es) for a hostname
'   =NT_PTR(A2)       -> reverse DNS name for an IP (or a hostname's IP)
'   =NT_PING(A2)      -> "Up (4/4, 12 ms)" / "Timed out" / ...
'   =NT_SUBNET(A2, "network")   -> "10.1.4.0"  (see NT_SUBNET for fields)
'   =NT_INSUBNET(B2, A2)        -> TRUE if the IP in B2 is inside subnet A2
'
' Requires Excel 2010 or later (VBA7), 32- or 64-bit, on Windows.
'==============================================================================
Option Explicit

' ---- Settings ----------------------------------------------------------------
Private Const PING_COUNT As Long = 4             ' echo requests per host (like Windows ping)
Private Const PING_TIMEOUT_MS As Long = 1000     ' per echo request
Private Const PING_DEAD_AFTER As Long = 2        ' stop pinging a host after this many misses in a row
Private Const APP_NAME As String = "NetTools"
Private Const APP_VERSION As String = "3.1.0"

' ---- Result column headers ---------------------------------------------------
Private Const H_HOST As String = "Host / IP"
Private Const H_TYPE As String = "Input Type"
Private Const H_IPS As String = "Resolved IP(s)"
Private Const H_PTR As String = "Reverse DNS (PTR)"
Private Const H_MATCH As String = "Fwd/Rev Match"
Private Const H_PINGIP As String = "Pinged Address"
Private Const H_PING As String = "Ping"
Private Const H_REPLIES As String = "Replies"
Private Const H_LAT As String = "Avg Latency (ms)"
Private Const H_TTL As String = "TTL"

' ---- Subnet calculator column headers ---------------------------------------
Private Const S_INPUT As String = "Subnet (as entered)"
Private Const S_CIDR As String = "Network/Prefix"
Private Const S_NET As String = "Network Address"
Private Const S_MASK As String = "Subnet Mask"
Private Const S_WILD As String = "Wildcard Mask"
Private Const S_FIRST As String = "First Usable"
Private Const S_LAST As String = "Last Usable"
Private Const S_BCAST As String = "Broadcast Address"
Private Const S_HOSTS As String = "Usable Hosts"
Private Const S_TOTAL As String = "Total Addresses"

' ---- Subnet scan column headers ------------------------------------------------
Private Const SC_IP As String = "IP Address"
Private Const SC_LAT As String = "Latency (ms)"
Private Const SC_MAC As String = "MAC Address"
Private Const SC_VENDOR As String = "Vendor"
Private Const SC_OPEN As String = "Open Ports"
Private Const RESULT_TABLE_STYLE As String = "TableStyleMedium2"
Private Const RESULT_HEADER_ROW As Long = 4

' ---- Winsock -----------------------------------------------------------------
Private Const AF_INET As Long = 2
Private Const AF_INET6 As Long = 23
Private Const SOCK_STREAM As Long = 1
Private Const AI_NUMERICHOST As Long = 4
Private Const NI_NUMERICHOST As Long = 2
Private Const NI_NAMEREQD As Long = 4
Private Const NI_MAXHOST As Long = 1025

#If Win64 Then
Private Const PTR_SIZE As Long = 8
#Else
Private Const PTR_SIZE As Long = 4
#End If

' struct addrinfo { int flags; int family; int socktype; int protocol;
'                   size_t addrlen; char *canonname; sockaddr *addr; addrinfo *next; }
Private Const AI_OFS_FAMILY As Long = 4
Private Const AI_OFS_ADDRLEN As Long = 16
Private Const AI_OFS_ADDR As Long = 16 + 2 * PTR_SIZE
Private Const AI_OFS_NEXT As Long = 16 + 3 * PTR_SIZE

Private Declare PtrSafe Function WSAStartup Lib "ws2_32.dll" ( _
    ByVal wVersionRequested As Integer, ByRef lpWSAData As Any) As Long
Private Declare PtrSafe Function WSACleanup Lib "ws2_32.dll" () As Long
Private Declare PtrSafe Function getaddrinfo Lib "ws2_32.dll" ( _
    ByVal pNodeName As String, ByVal pServiceName As LongPtr, _
    ByVal pHints As LongPtr, ByRef ppResult As LongPtr) As Long
Private Declare PtrSafe Sub freeaddrinfo Lib "ws2_32.dll" (ByVal pAddrInfo As LongPtr)
Private Declare PtrSafe Function getnameinfo Lib "ws2_32.dll" ( _
    ByVal pSockaddr As LongPtr, ByVal SockaddrLength As Long, _
    ByVal pNodeBuffer As String, ByVal NodeBufferSize As Long, _
    ByVal pServiceBuffer As LongPtr, ByVal ServiceBufferSize As Long, _
    ByVal Flags As Long) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, ByVal Source As LongPtr, ByVal Length As LongPtr)

' ---- ICMP (traceroute) -------------------------------------------------------
Private Const TRACE_MAX_HOPS As Long = 30
Private Const TRACE_PROBES As Long = 3
Private Const TRACE_TIMEOUT_MS As Long = 1000
Private Const TRACE_GIVE_UP_AFTER As Long = 5    ' stop after this many silent hops in a row
Private Const TRACE_MAX_HOSTS As Long = 20
Private Const SCAN_TIMEOUT_MS As Long = 1000     ' per ping when scanning (a 2nd ping goes to non-repliers)
Private Const SCAN_PARALLEL As Long = 256        ' pings in flight at once
Private Const ICMP_REPLY_BUF As Long = 512       ' reply buffer per ping
Private Const PORT_TIMEOUT_MS As Long = 2500     ' TCP connect timeout
Private Const PORT_PARALLEL As Long = 200        ' TCP connects in flight at once
Private Const MAX_PORTS As Long = 64
Private Const DEFAULT_PORTS As String = "22,80,443,3389"
Private Const MAC_CISCO As Long = 1              ' aabb.ccdd.eeff
Private Const MAC_COLON As Long = 2              ' AA:BB:CC:DD:EE:FF
Private Const SCAN_MAX_ADDRESSES As Long = 4096  ' biggest scan allowed (a /20)
Private Const SCAN_CONFIRM_ABOVE As Long = 256   ' ask before scanning more than this
Private Const DNS_TYPE_PTR As Integer = 12
Private Const IP_SUCCESS As Long = 0
Private Const IP_REQ_TIMED_OUT As Long = 11010
Private Const IP_TTL_EXPIRED_TRANSIT As Long = 11013

Private Declare PtrSafe Function IcmpCreateFile Lib "iphlpapi.dll" () As LongPtr
Private Declare PtrSafe Function IcmpCloseHandle Lib "iphlpapi.dll" (ByVal IcmpHandle As LongPtr) As Long
Private Declare PtrSafe Function IcmpSendEcho Lib "iphlpapi.dll" ( _
    ByVal IcmpHandle As LongPtr, ByVal DestinationAddress As Long, _
    ByVal RequestData As LongPtr, ByVal RequestSize As Integer, _
    ByVal RequestOptions As LongPtr, ByVal ReplyBuffer As LongPtr, _
    ByVal ReplySize As Long, ByVal Timeout As Long) As Long

' ---- DNS API (PTR lookups for the scanner: DNS only, no NetBIOS/LLMNR waits) --
Private Declare PtrSafe Function DnsQuery_A Lib "dnsapi.dll" ( _
    ByVal pszName As String, ByVal wType As Integer, ByVal Options As Long, _
    ByVal pExtra As LongPtr, ByRef ppQueryResults As LongPtr, ByVal pReserved As LongPtr) As Long
Private Declare PtrSafe Sub DnsRecordListFree Lib "dnsapi.dll" (ByVal pRecordList As LongPtr, ByVal FreeType As Long)
Private Declare PtrSafe Function lstrlenA Lib "kernel32" (ByVal lpString As LongPtr) As Long

' ---- Parallel ping (IcmpSendEcho2 + events) -----------------------------------
Private Declare PtrSafe Function IcmpSendEcho2 Lib "iphlpapi.dll" ( _
    ByVal IcmpHandle As LongPtr, ByVal hEvent As LongPtr, ByVal ApcRoutine As LongPtr, _
    ByVal ApcContext As LongPtr, ByVal DestinationAddress As Long, ByVal RequestData As LongPtr, _
    ByVal RequestSize As Integer, ByVal RequestOptions As LongPtr, ByVal ReplyBuffer As LongPtr, _
    ByVal ReplySize As Long, ByVal Timeout As Long) As Long
Private Declare PtrSafe Function IcmpParseReplies Lib "iphlpapi.dll" ( _
    ByVal ReplyBuffer As LongPtr, ByVal ReplySize As Long) As Long
Private Declare PtrSafe Function CreateEventA Lib "kernel32" (ByVal lpEventAttributes As LongPtr, _
    ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As LongPtr) As LongPtr
Private Declare PtrSafe Function SetEvent Lib "kernel32" (ByVal hEvent As LongPtr) As Long
Private Declare PtrSafe Function ResetEvent Lib "kernel32" (ByVal hEvent As LongPtr) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function WaitForMultipleObjects Lib "kernel32" (ByVal nCount As Long, _
    ByVal lpHandles As LongPtr, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long
Private Declare PtrSafe Sub RtlZeroMemory Lib "kernel32" (ByVal Destination As LongPtr, ByVal Length As LongPtr)

' ---- ARP table (MAC addresses of devices on directly connected subnets) -------
Private Declare PtrSafe Function GetIpNetTable Lib "iphlpapi.dll" (ByVal pIpNetTable As LongPtr, _
    ByRef pdwSize As Long, ByVal bOrder As Long) As Long

' ---- TCP port check (non-blocking connect + select) ---------------------------
Private Const FIONBIO As Long = &H8004667E
Private Const SOL_SOCKET As Long = &HFFFF&
Private Const SO_ERROR As Long = &H1007
Private Const IPPROTO_TCP As Long = 6
Private Const WSAEWOULDBLOCK As Long = 10035
Private Const WSAECONNREFUSED As Long = 10061
Private Const WSAENETUNREACH As Long = 10051
Private Const WSAEHOSTUNREACH As Long = 10065
Private Declare PtrSafe Function ws_socket Lib "ws2_32.dll" Alias "socket" ( _
    ByVal af As Long, ByVal stype As Long, ByVal protocol As Long) As LongPtr
Private Declare PtrSafe Function closesocket Lib "ws2_32.dll" (ByVal s As LongPtr) As Long
Private Declare PtrSafe Function ioctlsocket Lib "ws2_32.dll" (ByVal s As LongPtr, _
    ByVal cmd As Long, ByRef argp As Long) As Long
Private Declare PtrSafe Function ws_connect Lib "ws2_32.dll" Alias "connect" (ByVal s As LongPtr, _
    ByVal name As LongPtr, ByVal namelen As Long) As Long
Private Declare PtrSafe Function ws_select Lib "ws2_32.dll" Alias "select" (ByVal nfds As Long, _
    ByVal readfds As LongPtr, ByVal writefds As LongPtr, ByVal exceptfds As LongPtr, _
    ByVal timeout As LongPtr) As Long
Private Declare PtrSafe Function getsockopt Lib "ws2_32.dll" (ByVal s As LongPtr, ByVal level As Long, _
    ByVal optname As Long, ByRef optval As Long, ByRef optlen As Long) As Long
Private Declare PtrSafe Function WSAGetLastError Lib "ws2_32.dll" () As Long

Private mWmi As Object
Private mOui As Object          ' MAC vendor list (OUI -> vendor), loaded on first use
Private mOuiTried As Boolean

' Result of parsing/calculating one IPv4 subnet (addresses as 0..2^32-1)
Private Type SubnetInfo
    Ok As Boolean
    ErrorText As String
    Address As Double       ' the address as typed (may be a host inside the subnet)
    Prefix As Long
    Network As Double
    Broadcast As Double
    FirstHost As Double
    LastHost As Double
    Mask As Double
    Hosts As Double
End Type

' Result of checking one host
Private Type HostResult
    InputType As String
    Addresses As String     ' comma separated
    FirstAddr As String     ' address that gets pinged
    Ptr As String
    Match As String
    PingStatus As String
    Replies As String       ' e.g. "3/4"
    Latency As Variant
    Ttl As Variant
    DnsFailed As Boolean
    PingUp As Boolean       ' at least one reply
    PingLoss As Boolean     ' some, but not all, replies lost
End Type

'==============================================================================
' Ribbon callbacks
'==============================================================================
Public Sub NT_OnCheckDns(control As IRibbonControl)
    RunChecks True, True
End Sub

Public Sub NT_OnDnsOnly(control As IRibbonControl)
    RunChecks True, False
End Sub

Public Sub NT_OnPingOnly(control As IRibbonControl)
    RunChecks False, True
End Sub

Public Sub NT_OnPortCheck(control As IRibbonControl)
    RunPortCheck
End Sub

Public Sub NT_OnMacFormat(control As IRibbonControl)
    RunMacFormat
End Sub

Public Sub NT_OnTraceroute(control As IRibbonControl)
    RunTraceroute
End Sub

Public Sub NT_OnSubnetScan(control As IRibbonControl)
    RunSubnetScan
End Sub

Public Sub NT_OnSubnetCalc(control As IRibbonControl)
    RunSubnetCalc
End Sub

Public Sub NT_OnAbout(control As IRibbonControl)
    MsgBox APP_NAME & " " & APP_VERSION & vbCrLf & vbCrLf & _
        "1. Select a column of IP addresses, hostnames or subnets" & vbCrLf & _
        "   (select just the first cell to run to the end of the list)," & vbCrLf & _
        "   or select an empty cell to type them in." & vbCrLf & _
        "2. Click a button: Check DNS, DNS Only, Ping Only, Port Check," & vbCrLf & _
        "   Traceroute, Subnet Scan, Subnet Calc or MAC Format." & vbCrLf & _
        "3. Results go on their own sheet as a table (rerunning overwrites" & vbCrLf & _
        "   that sheet) - the data you selected is never changed." & vbCrLf & _
        "   Press Esc to stop a long run." & vbCrLf & vbCrLf & _
        "A header row (e.g. 'Hostname' or 'IP Address') is detected and skipped." & vbCrLf & _
        "Each traceroute and each subnet scan gets its own sheet." & vbCrLf & vbCrLf & _
        "Worksheet functions: =NT_RESOLVE(A2)  =NT_PTR(A2)  =NT_PING(A2)" & vbCrLf & _
        "   =NT_SUBNET(A2, ""network"")  =NT_INSUBNET(B2, A2)" & vbCrLf & _
        "   =NT_MAC(A2, ""cisco"")  =NT_VENDOR(A2)" & vbCrLf & vbCrLf & _
        "MAC vendor list: " & IIf(OuiList().Count > 0, Format$(OuiList().Count, "#,##0") & " vendors", "not installed") & vbCrLf & _
        "Ping: up to " & PING_COUNT & " echo requests per host (" & PING_TIMEOUT_MS & " ms timeout), " & _
        "all hosts at once; a host that misses " & PING_DEAD_AFTER & " in a row counts as down.", _
        vbInformation, "About " & APP_NAME
End Sub

' Same actions as plain macros (Alt+F8 / quick access toolbar)
Public Sub NetTools_CheckDns()
    RunChecks True, True
End Sub

Public Sub NetTools_DnsOnly()
    RunChecks True, False
End Sub

Public Sub NetTools_PingOnly()
    RunChecks False, True
End Sub

Public Sub NetTools_PortCheck()
    RunPortCheck
End Sub

Public Sub NetTools_MacFormat()
    RunMacFormat
End Sub

Public Sub NetTools_Traceroute()
    RunTraceroute
End Sub

Public Sub NetTools_SubnetScan()
    RunSubnetScan
End Sub

Public Sub NetTools_SubnetCalc()
    RunSubnetCalc
End Sub

'==============================================================================
' Main routine: Check DNS / DNS Only / Ping Only
'==============================================================================
Private Sub RunChecks(ByVal doDns As Boolean, ByVal doPing As Boolean)
    Dim items As Collection, srcDesc As String, ws As Worksheet, lo As ListObject
    Dim headers As Variant, title As String, rowRng As Range
    Dim i As Long, total As Long, done As Long, host As String, res As HostResult
    Dim nUp As Long, nDown As Long, nLoss As Long, nDnsFail As Long, nMismatch As Long
    Dim t0 As Double, wsaOk As Boolean, msg As String, results() As HostResult

    Set items = GetInputList("Hostnames or IP addresses to check, separated by commas or spaces:", False, srcDesc)
    If items Is Nothing Then Exit Sub
    total = items.Count

    If doDns And doPing Then
        title = "DNS Check"
    ElseIf doDns Then
        title = "DNS Lookup"
    Else
        title = "Ping"
    End If
    headers = BuildHeaders(doDns, doPing)

    t0 = Timer
    Application.EnableCancelKey = xlErrorHandler
    On Error GoTo Handler

    Set ws = NewResultSheet(title)
    Set lo = StartResultTable(ws, title & "  -  " & srcDesc, headers, total)

    wsaOk = WsaInit()
    If Not wsaOk Then
        MsgBox "Could not start Windows networking (WSAStartup failed).", vbCritical, APP_NAME
        GoTo Cleanup
    End If

    ' 1. Look up every host (rows appear as they're done)
    ReDim results(1 To total)
    For i = 1 To total
        host = items(i)
        Application.StatusBar = APP_NAME & ": " & IIf(doDns, "DNS ", "resolving ") & host & _
                                "  (" & i & " of " & total & ")   -   press Esc to stop"
        results(i) = CheckHost(host, doDns, False)
        Set rowRng = lo.DataBodyRange.Rows(i)
        rowRng.Value = ResultToRow(host, results(i), doDns, doPing)
        If Not doPing Then ColourRow rowRng, headers, results(i)
        done = i
        DoEvents
    Next i

    ' 2. Ping them all at once
    If doPing Then
        PingAll results, total
        Application.ScreenUpdating = False
        For i = 1 To total
            Set rowRng = lo.DataBodyRange.Rows(i)
            rowRng.Value = ResultToRow(CStr(items(i)), results(i), doDns, doPing)
            ColourRow rowRng, headers, results(i)
        Next i
        Application.ScreenUpdating = True
    End If

    For i = 1 To total
        If results(i).DnsFailed Then nDnsFail = nDnsFail + 1
        If doDns And Left$(results(i).Match, 2) = "No" Then nMismatch = nMismatch + 1
        If doPing Then
            If results(i).PingUp Then nUp = nUp + 1 Else nDown = nDown + 1
            If results(i).PingLoss Then nLoss = nLoss + 1
        End If
    Next i

    FinishTable lo, done
    msg = "Checked " & done & " host(s) in " & Format$(ElapsedSeconds(t0), "0.0") & " s." & vbCrLf & vbCrLf
    If doDns Then
        msg = msg & "DNS lookup failed: " & nDnsFail & vbCrLf & _
                    "Forward/reverse mismatch or no PTR: " & nMismatch & vbCrLf
    End If
    If doPing Then
        msg = msg & "Ping up: " & nUp & vbCrLf & _
                    "   ...with packet loss: " & nLoss & vbCrLf & _
                    "Ping down: " & nDown & vbCrLf
    End If
    msg = msg & vbCrLf & "Results are on the sheet """ & ws.Name & """."
    Application.StatusBar = False
    MsgBox msg, vbInformation, APP_NAME
    GoTo Cleanup

Handler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    If Err.Number = 18 Then
        If Not lo Is Nothing Then FinishTable lo, done
        MsgBox "Stopped after " & done & " of " & total & " host(s).", vbInformation, APP_NAME
    Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, APP_NAME
    End If

Cleanup:
    On Error Resume Next
    If wsaOk Then WSACleanup
    Application.StatusBar = False
    Application.EnableCancelKey = xlInterrupt
End Sub

'==============================================================================
' Input and result sheets
'==============================================================================

' What to check: the values in the selected column (header skipped), or - if
' an empty cell is selected - whatever the user types in. raw=True keeps the
' text as-is (subnets); otherwise each value is tidied into a host name.
' Returns Nothing if there is nothing to do.
Private Function GetInputList(ByVal prompt As String, ByVal raw As Boolean, ByRef srcDesc As String) As Collection
    Dim items As New Collection, ws As Worksheet, typed As String, part As Variant, s As String
    Dim col As Long, firstRow As Long, lastRow As Long, headerRow As Long, r As Long
    Dim askUser As Boolean

    askUser = True
    If TypeName(Selection) = "Range" Then
        If Not (Selection.Cells.CountLarge = 1 And Len(Trim$(Selection.Cells(1, 1).Text)) = 0) Then
            askUser = False
            Set ws = Selection.Worksheet
            If Not GetInputRows(ws, Selection, col, firstRow, lastRow, headerRow) Then Exit Function
            For r = firstRow To lastRow
                s = CellInput(ws.Cells(r, col), raw)
                If Len(s) > 0 Then items.Add s
            Next r
            srcDesc = "from " & ws.Name & "!" & _
                      ws.Range(ws.Cells(firstRow, col), ws.Cells(lastRow, col)).Address(False, False)
            If items.Count = 0 Then
                MsgBox "The selected cells are empty.", vbExclamation, APP_NAME
                Exit Function
            End If
        End If
    End If

    If askUser Then
        typed = InputBox(prompt & vbCrLf & vbCrLf & _
                         "(Or select a column of them first, then click the button.)", APP_NAME)
        typed = Replace(Replace(Replace(typed, ";", ","), vbCr, ","), vbLf, ",")
        If Not raw Then typed = Replace(typed, " ", ",")
        For Each part In Split(typed, ",")
            If raw Then s = Trim$(CStr(part)) Else s = CleanHost(part)
            If Len(s) > 0 Then items.Add s
        Next part
        If items.Count = 0 Then Exit Function
        srcDesc = "typed in"
    End If

    Set GetInputList = items
End Function

Private Function CellInput(ByVal cell As Range, ByVal raw As Boolean) As String
    If IsError(cell.Value) Then Exit Function
    If raw Then
        CellInput = Trim$(Replace(Replace(cell.Text, vbTab, " "), Chr$(160), " "))
    Else
        CellInput = CleanHost(cell.Value)
    End If
End Function

' Works out which column/rows to read from the selection.
Private Function GetInputRows(ByVal ws As Worksheet, ByVal sel As Range, _
        ByRef col As Long, ByRef firstRow As Long, ByRef lastRow As Long, _
        ByRef headerRow As Long) As Boolean
    Dim rng As Range

    Set rng = sel.Areas(1).Columns(1)
    col = rng.Column

    If rng.Cells.CountLarge = 1 Then
        ' Single cell: run from here to the last value in the column
        firstRow = rng.Row
        lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
        If lastRow < firstRow Then lastRow = firstRow
    Else
        ' Trim whole-column selections to the used range
        Set rng = Intersect(rng, ws.UsedRange)
        If rng Is Nothing Then
            MsgBox "The selected column is empty.", vbExclamation, APP_NAME
            Exit Function
        End If
        firstRow = rng.Row
        lastRow = rng.Row + rng.Rows.Count - 1
    End If

    If LooksLikeHeader(CStr(ws.Cells(firstRow, col).Text)) Then
        headerRow = firstRow
        firstRow = firstRow + 1
    End If
    If firstRow > lastRow Then
        MsgBox "There is nothing below the header to check.", vbExclamation, APP_NAME
        Exit Function
    End If
    GetInputRows = True
End Function

Private Function LooksLikeHeader(ByVal s As String) As Boolean
    s = LCase$(Trim$(s))
    If Len(s) = 0 Then Exit Function
    If InStr(s, " ") > 0 And InStr(s, ".") = 0 Then LooksLikeHeader = True: Exit Function
    Select Case s
        Case "host", "hosts", "hostname", "hostnames", "host_name", "ip", "ips", "ipaddress", _
             "ip_address", "ip-address", "address", "addresses", "name", "names", "device", _
             "devices", "fqdn", "server", "servers", "target", "targets", "node", "nodes", _
             "dns", "dnsname", "dns_name", "subnet", "subnets", "cidr", "network", "networks", _
             "prefix", "prefixes", "range", "ranges", "vlan", "vlans", "mac", "macs", _
             "mac_address", "macaddress", "hwaddr", "bia"
            LooksLikeHeader = True
    End Select
End Function

' Returns the results sheet called base, e.g. "Subnets". If NetTools made
' that sheet on an earlier run it is cleared and reused (overwritten). A
' sheet of yours with the same name is never touched - then a new sheet
' "Subnets 2" (etc.) is used instead.
Private Function NewResultSheet(ByVal base As String) As Worksheet
    Dim nm As String, n As Long, ch As Variant, wb As Workbook, ws As Worksheet, lo As ListObject
    If ActiveWorkbook Is Nothing Then Workbooks.Add
    Set wb = ActiveWorkbook
    For Each ch In Array(":", "\", "/", "?", "*", "[", "]")
        base = Replace(base, CStr(ch), "_")
    Next ch
    base = Left$(base, 27)
    nm = base
    n = 1
    Do While SheetExists(wb, nm)
        If IsNetToolsSheet(wb.Sheets(nm)) Then
            ' Earlier NetTools results: overwrite them
            Set ws = wb.Worksheets(nm)
            For Each lo In ws.ListObjects
                lo.Delete
            Next lo
            ws.Cells.Clear
            Set NewResultSheet = ws
            Exit Function
        End If
        n = n + 1
        nm = base & " " & n
    Loop
    If TypeName(ActiveSheet) = "Worksheet" Then
        Set ws = wb.Worksheets.Add(After:=ActiveSheet)
    Else
        Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    End If
    ws.Name = nm
    Set NewResultSheet = ws
End Function

' NetTools result sheets carry "... by NetTools x.y.z" in A2 (see WriteSheetTitle).
Private Function IsNetToolsSheet(ByVal sh As Object) As Boolean
    On Error Resume Next
    If TypeName(sh) <> "Worksheet" Then Exit Function
    IsNetToolsSheet = (InStr(1, CStr(sh.Range("A2").Value), " by " & APP_NAME & " ", vbBinaryCompare) > 0)
End Function

Private Function SheetExists(ByVal wb As Workbook, ByVal nm As String) As Boolean
    Dim sh As Object
    On Error Resume Next
    Set sh = wb.Sheets(nm)
    SheetExists = Not sh Is Nothing
End Function

' Title in A1, run details in A2, then an Excel table (header row + nRows
' empty rows) starting at row RESULT_HEADER_ROW.
Private Function StartResultTable(ByVal ws As Worksheet, ByVal title As String, _
        ByVal headers As Variant, ByVal nRows As Long) As ListObject
    Dim nCols As Long, i As Long, lo As ListObject
    nCols = UBound(headers) - LBound(headers) + 1
    If nRows < 1 Then nRows = 1

    WriteSheetTitle ws, title
    For i = 0 To nCols - 1
        With ws.Range(ws.Cells(RESULT_HEADER_ROW + 1, i + 1), ws.Cells(RESULT_HEADER_ROW + nRows, i + 1))
            If IsNumericColumn(CStr(headers(LBound(headers) + i))) Then
                .NumberFormat = "General"
                .HorizontalAlignment = xlRight
            Else
                .NumberFormat = "@"   ' stops Excel turning IPv6, "3/4" etc. into numbers/dates
            End If
        End With
    Next i
    ws.Range(ws.Cells(RESULT_HEADER_ROW, 1), ws.Cells(RESULT_HEADER_ROW, nCols)).Value = headers

    Set lo = ws.ListObjects.Add(xlSrcRange, _
        ws.Range(ws.Cells(RESULT_HEADER_ROW, 1), ws.Cells(RESULT_HEADER_ROW + nRows, nCols)), , xlYes)
    lo.TableStyle = RESULT_TABLE_STYLE
    lo.Range.Columns.AutoFit
    FreezeHeader ws
    Set StartResultTable = lo
End Function

Private Sub WriteSheetTitle(ByVal ws As Worksheet, ByVal title As String)
    With ws.Cells(1, 1)
        .Value = title
        .Font.Bold = True
        .Font.Size = 13
    End With
    ws.Cells(2, 1).Value = "Run " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " on " & Environ$("COMPUTERNAME") & _
                           " by " & APP_NAME & " " & APP_VERSION
    ws.Cells(2, 1).Font.Color = RGB(89, 89, 89)
End Sub

' Keeps the title and table header visible while scrolling.
Private Sub FreezeHeader(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.ScrollRow = 1
    ws.Cells(RESULT_HEADER_ROW + 1, 1).Select
    ActiveWindow.FreezePanes = True
End Sub

' Trims unused rows (if the run was stopped early) and sizes the columns.
Private Sub FinishTable(ByVal lo As ListObject, ByVal rowsUsed As Long)
    Dim c As Range
    On Error Resume Next
    If rowsUsed < 1 Then rowsUsed = 1
    If Not lo.DataBodyRange Is Nothing Then
        If lo.DataBodyRange.Rows.Count > rowsUsed Then
            lo.DataBodyRange.Rows(rowsUsed + 1).Resize(lo.DataBodyRange.Rows.Count - rowsUsed).Delete
        End If
    End If
    lo.Range.Columns.AutoFit
    For Each c In lo.Range.Columns
        If c.ColumnWidth > 60 Then c.ColumnWidth = 60
    Next c
End Sub

Private Function IsNumericColumn(ByVal header As String) As Boolean
    Select Case header
        Case H_LAT, H_TTL, S_HOSTS, S_TOTAL, SC_LAT, "Hop"
            IsNumericColumn = True
    End Select
End Function

Private Function BuildHeaders(ByVal doDns As Boolean, ByVal doPing As Boolean) As Variant
    If doDns And doPing Then
        BuildHeaders = Array(H_HOST, H_TYPE, H_IPS, H_PTR, H_MATCH, H_PING, H_REPLIES, H_LAT, H_TTL)
    ElseIf doDns Then
        BuildHeaders = Array(H_HOST, H_TYPE, H_IPS, H_PTR, H_MATCH)
    Else
        BuildHeaders = Array(H_HOST, H_PINGIP, H_PING, H_REPLIES, H_LAT, H_TTL)
    End If
End Function

Private Function ResultToRow(ByVal host As String, ByRef res As HostResult, _
        ByVal doDns As Boolean, ByVal doPing As Boolean) As Variant()
    Dim v() As Variant
    If doDns And doPing Then
        ReDim v(0 To 8)
        v(0) = host: v(1) = res.InputType: v(2) = res.Addresses: v(3) = res.Ptr: v(4) = res.Match
        v(5) = res.PingStatus: v(6) = res.Replies: v(7) = res.Latency: v(8) = res.Ttl
    ElseIf doDns Then
        ReDim v(0 To 4)
        v(0) = host: v(1) = res.InputType: v(2) = res.Addresses: v(3) = res.Ptr: v(4) = res.Match
    Else
        ReDim v(0 To 5)
        v(0) = host: v(1) = res.FirstAddr: v(2) = res.PingStatus: v(3) = res.Replies
        v(4) = res.Latency: v(5) = res.Ttl
    End If
    ResultToRow = v
End Function

Private Sub ColourRow(ByVal rowRng As Range, ByVal headers As Variant, ByRef res As HostResult)
    Dim i As Long, cell As Range
    For i = LBound(headers) To UBound(headers)
        Set cell = rowRng.Cells(1, i - LBound(headers) + 1)
        Select Case headers(i)
            Case H_PING, H_REPLIES
                If res.PingLoss Then
                    Warn cell
                ElseIf res.PingUp Then
                    Good cell
                Else
                    Bad cell
                End If
            Case H_MATCH
                If res.Match = "Yes" Then
                    Good cell
                ElseIf Left$(res.Match, 2) = "No" Then
                    Warn cell
                End If
            Case H_IPS, H_PINGIP
                If res.DnsFailed Then Bad cell
        End Select
    Next i
End Sub

Private Sub Good(ByVal c As Range)
    c.Interior.Color = RGB(198, 239, 206): c.Font.Color = RGB(0, 97, 0)
End Sub

Private Sub Bad(ByVal c As Range)
    c.Interior.Color = RGB(255, 199, 206): c.Font.Color = RGB(156, 0, 6)
End Sub

Private Sub Warn(ByVal c As Range)
    c.Interior.Color = RGB(255, 235, 156): c.Font.Color = RGB(156, 101, 0)
End Sub

'==============================================================================
' Checking one host
'==============================================================================
Private Function CheckHost(ByVal host As String, ByVal doDns As Boolean, ByVal doPing As Boolean) As HostResult
    Dim res As HostResult, addrs As Collection, fwd As Collection, canonical As String
    Dim lat As Variant, ttl As Variant, received As Long, sent As Long

    res.Latency = Empty
    res.Ttl = Empty
    canonical = NumericAddress(host)

    If Len(canonical) > 0 Then
        ' ---- Input is an IP address ----
        res.InputType = "IP"
        res.Addresses = canonical
        res.FirstAddr = canonical
        If doDns Then
            res.Ptr = FastPtr(canonical)
            If Len(res.Ptr) = 0 Then
                res.Ptr = "(no PTR record)"
                res.Match = "No PTR"
            Else
                ' Forward-confirm: does the PTR name resolve back to this IP?
                Set fwd = LookupAddresses(res.Ptr)
                If fwd.Count = 0 Then
                    res.Match = "No - PTR name does not resolve"
                ElseIf InCollection(fwd, canonical) Then
                    res.Match = "Yes"
                Else
                    res.Match = "No - PTR name resolves to " & JoinCollection(fwd)
                End If
            End If
        End If
    Else
        ' ---- Input is a hostname ----
        res.InputType = "Hostname"
        Set addrs = LookupAddresses(host)
        If addrs.Count = 0 Then
            res.DnsFailed = True
            res.Addresses = "(not found)"
            res.FirstAddr = "(not found)"
            res.Ptr = ""
            res.Match = "n/a"
            res.PingStatus = "Not resolved"
            CheckHost = res
            Exit Function
        End If
        res.Addresses = JoinCollection(addrs)
        res.FirstAddr = addrs(1)
        If doDns Then
            res.Ptr = FastPtr(res.FirstAddr)
            If Len(res.Ptr) = 0 Then
                res.Ptr = "(no PTR record)"
                res.Match = "No PTR"
            ElseIf NamesMatch(host, res.Ptr) Then
                res.Match = "Yes"
            Else
                res.Match = "No - PTR is a different name"
            End If
        End If
    End If

    If doPing Then
        res.PingStatus = PingHost(res.FirstAddr, lat, ttl, received, sent)
        res.Replies = received & "/" & sent
        res.PingUp = (received > 0)
        res.PingLoss = (received > 0 And received < sent)
        res.Latency = lat
        res.Ttl = ttl
    End If

    CheckHost = res
End Function

' "server1" matches "server1.corp.example.com"; full names must match exactly.
Private Function NamesMatch(ByVal a As String, ByVal b As String) As Boolean
    a = LCase$(a): b = LCase$(b)
    If Right$(a, 1) = "." Then a = Left$(a, Len(a) - 1)
    If Right$(b, 1) = "." Then b = Left$(b, Len(b) - 1)
    If a = b Then NamesMatch = True: Exit Function
    If InStr(a, ".") = 0 Then
        NamesMatch = (a = Split(b, ".")(0))
    ElseIf InStr(b, ".") = 0 Then
        NamesMatch = (b = Split(a, ".")(0))
    End If
End Function

' Tidies what's in the cell: trims, drops URL scheme, path, CIDR suffix, IPv6 brackets.
Private Function CleanHost(ByVal v As Variant) As String
    Dim s As String, p As Long
    If IsError(v) Or IsEmpty(v) Then Exit Function
    s = Trim$(CStr(v))
    s = Replace(s, vbTab, "")
    s = Replace(s, Chr$(160), "")
    s = Trim$(s)
    p = InStr(s, "://")
    If p > 0 Then s = Mid$(s, p + 3)
    p = InStr(s, "/")
    If p > 0 Then s = Left$(s, p - 1)
    If Left$(s, 1) = "[" Then
        p = InStr(s, "]")
        If p > 0 Then s = Mid$(s, 2, p - 2)
    End If
    p = InStr(s, " ")
    If p > 0 Then s = Left$(s, p - 1)
    CleanHost = s
End Function

'==============================================================================
' DNS (Windows resolver via ws2_32)
'==============================================================================
Private Function WsaInit() As Boolean
    Dim wsaData(0 To 511) As Byte
    WsaInit = (WSAStartup(&H202, wsaData(0)) = 0)
End Function

Private Function ReadLong(ByVal p As LongPtr) As Long
    Dim v As Long
    CopyMemory v, p, 4
    ReadLong = v
End Function

Private Function ReadPtr(ByVal p As LongPtr) As LongPtr
    Dim v As LongPtr
    CopyMemory v, p, PTR_SIZE
    ReadPtr = v
End Function

Private Function TrimNull(ByVal s As String) As String
    Dim p As Long
    p = InStr(s, vbNullChar)
    If p > 0 Then TrimNull = Left$(s, p - 1) Else TrimNull = s
End Function

Private Function SockaddrToString(ByVal sa As LongPtr, ByVal saLen As Long, ByVal flags As Long) As String
    Dim buf As String
    If sa = 0 Then Exit Function
    buf = String$(NI_MAXHOST, vbNullChar)
    If getnameinfo(sa, saLen, buf, NI_MAXHOST, 0, 0, flags) = 0 Then
        SockaddrToString = TrimNull(buf)
    End If
End Function

' Returns the canonical text form if s is an IP address literal, else "".
Private Function NumericAddress(ByVal s As String) As String
    Dim hints(0 To 11) As Long, res As LongPtr
    hints(0) = AI_NUMERICHOST
    hints(2) = SOCK_STREAM
    If getaddrinfo(s, 0, VarPtr(hints(0)), res) = 0 Then
        If res <> 0 Then
            NumericAddress = SockaddrToString(ReadPtr(res + AI_OFS_ADDR), ReadLong(res + AI_OFS_ADDRLEN), NI_NUMERICHOST)
            freeaddrinfo res
        End If
    End If
End Function

' Forward lookup. Returns every address, IPv4 first, then IPv6.
Private Function LookupAddresses(ByVal host As String) As Collection
    Dim hints(0 To 11) As Long, res As LongPtr, p As LongPtr
    Dim v4 As New Collection, v6 As New Collection, all As New Collection
    Dim fam As Long, ip As String, item As Variant

    hints(2) = SOCK_STREAM          ' one entry per address instead of per socket type
    If getaddrinfo(host, 0, VarPtr(hints(0)), res) = 0 Then
        p = res
        Do While p <> 0
            fam = ReadLong(p + AI_OFS_FAMILY)
            ip = SockaddrToString(ReadPtr(p + AI_OFS_ADDR), ReadLong(p + AI_OFS_ADDRLEN), NI_NUMERICHOST)
            If Len(ip) > 0 Then
                If fam = AF_INET Then
                    AddUnique v4, ip
                ElseIf fam = AF_INET6 Then
                    AddUnique v6, ip
                End If
            End If
            p = ReadPtr(p + AI_OFS_NEXT)
        Loop
        If res <> 0 Then freeaddrinfo res
    End If

    For Each item In v4: all.Add item: Next item
    For Each item In v6: all.Add item: Next item
    Set LookupAddresses = all
End Function

' Reverse (PTR) lookup for an IP address. "" if there is no record.
Private Function ReverseLookup(ByVal ip As String) As String
    Dim hints(0 To 11) As Long, res As LongPtr
    hints(0) = AI_NUMERICHOST
    hints(2) = SOCK_STREAM
    If getaddrinfo(ip, 0, VarPtr(hints(0)), res) = 0 Then
        If res <> 0 Then
            ReverseLookup = SockaddrToString(ReadPtr(res + AI_OFS_ADDR), ReadLong(res + AI_OFS_ADDRLEN), NI_NAMEREQD)
            freeaddrinfo res
        End If
    End If
End Function

Private Sub AddUnique(ByVal c As Collection, ByVal s As String)
    If Not InCollection(c, s) Then c.Add s
End Sub

Private Function InCollection(ByVal c As Collection, ByVal s As String) As Boolean
    Dim item As Variant
    For Each item In c
        If StrComp(CStr(item), s, vbTextCompare) = 0 Then InCollection = True: Exit Function
    Next item
End Function

Private Function JoinCollection(ByVal c As Collection) As String
    Dim item As Variant, s As String
    For Each item In c
        If Len(s) > 0 Then s = s & ", "
        s = s & CStr(item)
    Next item
    JoinCollection = s
End Function

'==============================================================================
' Ping
'==============================================================================
' Pings every resolved host at the same time, in rounds of one ping each
' (like Windows ping's 4, but all hosts together - so 200 hosts take about as
' long as one). A host that misses PING_DEAD_AFTER pings in a row is down and
' gets no more. IPv6-only hosts fall back to WMI, one at a time.
Private Sub PingAll(results() As HostResult, ByVal n As Long)
    Dim ips() As String, todo() As Boolean, up() As Boolean, st() As Long, rtt() As Long, ttl() As Long
    Dim v4() As Boolean, sent() As Long, recv() As Long, misses() As Long, lastSt() As Long
    Dim rttSum() As Double, lastTtl() As Long, i As Long, pass As Long, anyLeft As Boolean
    Dim lat As Variant, tl As Variant, r As Long, sn As Long

    ReDim ips(1 To n), todo(1 To n), up(1 To n), st(1 To n), rtt(1 To n), ttl(1 To n)
    ReDim v4(1 To n), sent(1 To n), recv(1 To n), misses(1 To n), lastSt(1 To n), rttSum(1 To n), lastTtl(1 To n)
    For i = 1 To n
        ips(i) = "0.0.0.0"
        If Not results(i).DnsFailed And Len(results(i).FirstAddr) > 0 And InStr(results(i).FirstAddr, ":") = 0 _
           And InStr(results(i).FirstAddr, "(") = 0 Then
            v4(i) = True
            ips(i) = results(i).FirstAddr
        End If
    Next i

    For pass = 1 To PING_COUNT
        anyLeft = False
        For i = 1 To n
            todo(i) = v4(i) And misses(i) < PING_DEAD_AFTER
            up(i) = False: st(i) = 0
            If todo(i) Then anyLeft = True
        Next i
        If Not anyLeft Then Exit For
        PingMany ips, todo, PING_TIMEOUT_MS, up, st, rtt, ttl, "ping round " & pass & " of " & PING_COUNT
        For i = 1 To n
            If todo(i) Then
                sent(i) = sent(i) + 1
                If up(i) Then
                    recv(i) = recv(i) + 1
                    rttSum(i) = rttSum(i) + rtt(i)
                    lastTtl(i) = ttl(i)
                    misses(i) = 0
                Else
                    misses(i) = misses(i) + 1
                    lastSt(i) = st(i)
                End If
            End If
        Next i
    Next pass

    For i = 1 To n
        If v4(i) Then
            results(i).Replies = recv(i) & "/" & sent(i)
            results(i).PingUp = (recv(i) > 0)
            results(i).PingLoss = (recv(i) > 0 And recv(i) < sent(i))
            If recv(i) = 0 Then
                If lastSt(i) = 0 Then lastSt(i) = IP_REQ_TIMED_OUT
                results(i).PingStatus = PingStatusText(lastSt(i))
            Else
                results(i).Latency = Round(rttSum(i) / recv(i), 0)
                results(i).Ttl = lastTtl(i)
                If recv(i) = sent(i) Then results(i).PingStatus = "Up" Else results(i).PingStatus = "Up - packet loss"
            End If
        ElseIf Not results(i).DnsFailed And InStr(results(i).FirstAddr, ":") > 0 Then
            Application.StatusBar = APP_NAME & ": ping " & results(i).FirstAddr & " (IPv6)"
            results(i).PingStatus = PingHost(results(i).FirstAddr, lat, tl, r, sn)
            results(i).Replies = r & "/" & sn
            results(i).PingUp = (r > 0)
            results(i).PingLoss = (r > 0 And r < sn)
            results(i).Latency = lat
            results(i).Ttl = tl
        End If
    Next i
End Sub

' Reverse lookup: straight from DNS for IPv4 (quick even with no record),
' the Windows resolver for IPv6.
Private Function FastPtr(ByVal ip As String) As String
    If InStr(ip, ":") = 0 Then FastPtr = PtrLookupDns(ip) Else FastPtr = ReverseLookup(ip)
End Function

' Sends PING_COUNT echo requests (one at a time, like Windows ping).
' Returns "Up", "Up - packet loss", or the failure (e.g. "Timed out").
' latency = average round trip of the replies (ms), ttl = TTL of the last reply,
' received = number of replies.
Private Function PingHost(ByVal addr As String, ByRef latency As Variant, ByRef ttl As Variant, _
        ByRef received As Long, Optional ByRef sent As Long) As String
    Dim pings As Object, reply As Object, code As Variant
    Dim i As Long, totalMs As Double, lastFailure As String, before As Long, misses As Long
    latency = Empty
    ttl = Empty
    received = 0
    sent = 0
    lastFailure = "No response"
    On Error GoTo Fail

    If mWmi Is Nothing Then
        Set mWmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    End If

    For i = 1 To PING_COUNT
        before = received
        sent = i
        Set pings = mWmi.ExecQuery( _
            "SELECT StatusCode, ResponseTime, ResponseTimeToLive FROM Win32_PingStatus " & _
            "WHERE Address='" & Replace(addr, "'", "") & "' AND Timeout=" & PING_TIMEOUT_MS)
        For Each reply In pings
            code = reply.StatusCode
            If IsNull(code) Then
                lastFailure = "No response"
            ElseIf CLng(code) = 0 Then
                received = received + 1
                If Not IsNull(reply.ResponseTime) Then totalMs = totalMs + reply.ResponseTime
                If Not IsNull(reply.ResponseTimeToLive) Then ttl = reply.ResponseTimeToLive
            Else
                lastFailure = PingStatusText(CLng(code))
            End If
        Next reply
        DoEvents        ' keeps Excel responsive and lets Esc stop the run
        If received > before Then misses = 0 Else misses = misses + 1
        If misses >= PING_DEAD_AFTER Then Exit For      ' counts as down - stop here
    Next i

    If received = 0 Then
        PingHost = lastFailure
    Else
        latency = Round(totalMs / received, 0)
        If received = sent Then PingHost = "Up" Else PingHost = "Up - packet loss"
    End If
    Exit Function

Fail:
    If Err.Number = 18 Then
        Err.Raise 18     ' Esc pressed - let RunChecks stop
    End If
    PingHost = "Error: " & Err.Description
End Function

Private Function PingStatusText(ByVal code As Long) As String
    Select Case code
        Case 11002: PingStatusText = "Destination net unreachable"
        Case 11003: PingStatusText = "Destination host unreachable"
        Case 11004: PingStatusText = "Destination protocol unreachable"
        Case 11005: PingStatusText = "Destination port unreachable"
        Case 11010: PingStatusText = "Timed out"
        Case 11013: PingStatusText = "TTL expired in transit"
        Case 11018: PingStatusText = "Bad destination"
        Case 11050: PingStatusText = "General failure"
        Case Else:  PingStatusText = "Failed (" & code & ")"
    End Select
End Function

Private Function ElapsedSeconds(ByVal t0 As Double) As Double
    Dim t As Double
    t = Timer - t0
    If t < 0 Then t = t + 86400#     ' ran past midnight
    ElapsedSeconds = t
End Function


'==============================================================================
' Traceroute
'==============================================================================
Private Sub RunTraceroute()
    Dim hosts As Collection, srcDesc As String
    Dim item As Variant, hIcmp As LongPtr, wsaOk As Boolean

    Set hosts = GetInputList("Traceroute to (hostname or IP address - several can be separated by commas):", _
                             False, srcDesc)
    If hosts Is Nothing Then Exit Sub
    If hosts.Count > TRACE_MAX_HOSTS Then
        MsgBox "Traceroute runs on up to " & TRACE_MAX_HOSTS & " hosts at a time - " & _
               "select fewer.", vbExclamation, APP_NAME
        Exit Sub
    End If
    If hosts.Count > 3 Then
        If MsgBox("Run " & hosts.Count & " traceroutes? Each one gets its own sheet and can take " & _
                  "up to a minute or two.", vbOKCancel + vbQuestion, APP_NAME) <> vbOK Then Exit Sub
    End If
    If ActiveWorkbook Is Nothing Then Workbooks.Add

    Application.EnableCancelKey = xlErrorHandler
    On Error GoTo Handler
    wsaOk = WsaInit()
    hIcmp = IcmpCreateFile()
    If hIcmp = 0 Or hIcmp = -1 Then
        MsgBox "Could not open the Windows ICMP service (IcmpCreateFile failed).", vbCritical, APP_NAME
        GoTo Cleanup
    End If

    For Each item In hosts
        TraceOne CStr(item), hIcmp
    Next item
    GoTo Cleanup

Handler:
    Application.StatusBar = False
    If Err.Number = 18 Then
        MsgBox "Traceroute stopped.", vbInformation, APP_NAME
    Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, APP_NAME
    End If

Cleanup:
    On Error Resume Next
    If hIcmp <> 0 And hIcmp <> -1 Then IcmpCloseHandle hIcmp
    If wsaOk Then WSACleanup
    Application.StatusBar = False
    Application.EnableCancelKey = xlInterrupt
End Sub

Private Sub TraceOne(ByVal host As String, ByVal hIcmp As LongPtr)
    Dim ip As String, item As Variant, ws As Worksheet, dest As Long
    Dim hop As Long, probe As Long, r As Long, firstRow As Long
    Dim fromAddr As String, hopAddr As String, status As Long, rtt As Long
    Dim gotReply As Boolean, reached As Boolean, failed As Boolean
    Dim silent As Long, note As String, ptr As String, lo As ListObject

    ' Resolve to IPv4 (IcmpSendEcho is IPv4 only)
    ip = NumericAddress(host)
    If Len(ip) = 0 Then
        For Each item In LookupAddresses(host)
            If InStr(CStr(item), ":") = 0 Then ip = CStr(item): Exit For
        Next item
    End If
    If InStr(ip, ":") > 0 Then ip = ""
    If Len(ip) = 0 Then
        MsgBox host & ": could not resolve to an IPv4 address." & vbCrLf & _
               "(Traceroute supports IPv4 only.)", vbExclamation, APP_NAME
        Exit Sub
    End If
    dest = IPv4ToNetLong(ip)

    Set ws = NewResultSheet("Trace " & host)
    WriteSheetTitle ws, "Traceroute to " & host & _
        IIf(StrComp(host, ip, vbTextCompare) <> 0, "  [" & ip & "]", "") & _
        "  -  max " & TRACE_MAX_HOPS & " hops, " & TRACE_PROBES & " probes per hop, " & TRACE_TIMEOUT_MS & " ms timeout"
    With ws
        .Range("A4:G4").Value = Array("Hop", "Address", "Hostname", "Probe 1 (ms)", "Probe 2 (ms)", "Probe 3 (ms)", "Note")
        .Range("B:C").NumberFormat = "@"
        .Range("G:G").NumberFormat = "@"
        .Range("D:F").HorizontalAlignment = xlRight
        .Columns("A").ColumnWidth = 6
        .Columns("B").ColumnWidth = 18
        .Columns("C").ColumnWidth = 40
        .Columns("D:F").ColumnWidth = 13
        .Columns("G").ColumnWidth = 30
    End With
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A4:G5"), , xlYes)
    lo.TableStyle = RESULT_TABLE_STYLE
    FreezeHeader ws

    firstRow = 5
    r = firstRow
    For hop = 1 To TRACE_MAX_HOPS
        Application.StatusBar = APP_NAME & ": traceroute to " & host & " - hop " & hop & "   -   press Esc to stop"
        If r > lo.Range.Row + lo.Range.Rows.Count - 1 Then lo.Resize ws.Range(ws.Cells(4, 1), ws.Cells(r, 7))
        ws.Cells(r, 1).Value = hop
        hopAddr = "": note = "": gotReply = False

        For probe = 1 To TRACE_PROBES
            If SendProbe(hIcmp, dest, hop, TRACE_TIMEOUT_MS, fromAddr, status, rtt) Then
                gotReply = True
                If Len(hopAddr) = 0 Then hopAddr = fromAddr
                If rtt < 1 Then ws.Cells(r, 3 + probe).Value = "<1" Else ws.Cells(r, 3 + probe).Value = rtt
                If status = IP_SUCCESS Then
                    reached = True
                ElseIf status <> IP_TTL_EXPIRED_TRANSIT Then
                    failed = True
                    note = PingStatusText(status)
                End If
            Else
                ws.Cells(r, 3 + probe).Value = "*"
            End If
            DoEvents
        Next probe

        If gotReply Then
            silent = 0
            ws.Cells(r, 2).Value = hopAddr
            ptr = ReverseLookup(hopAddr)
            ws.Cells(r, 3).Value = ptr
        Else
            silent = silent + 1
            ws.Cells(r, 2).Value = "Request timed out"
            ws.Cells(r, 2).Font.Color = RGB(128, 128, 128)
        End If
        If reached Then note = "Destination reached"
        ws.Cells(r, 7).Value = note
        If reached Then ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).Interior.Color = RGB(198, 239, 206)
        If failed And Not reached Then ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).Interior.Color = RGB(255, 199, 206)
        r = r + 1
        DoEvents

        If reached Or failed Then Exit For
        If silent >= TRACE_GIVE_UP_AFTER Then
            note = "Stopped: " & TRACE_GIVE_UP_AFTER & " hops in a row with no reply."
            Exit For
        End If
    Next hop

    r = r + 1
    If reached Then
        ws.Cells(r, 1).Value = "Trace complete: " & ip & " reached in " & (hop) & " hop(s)."
    ElseIf failed Then
        ws.Cells(r, 1).Value = "Trace ended: " & note & "."
    ElseIf silent >= TRACE_GIVE_UP_AFTER Then
        ws.Cells(r, 1).Value = note & " The destination or a firewall may be blocking ICMP."
    Else
        ws.Cells(r, 1).Value = "Destination not reached within " & TRACE_MAX_HOPS & " hops."
    End If
    ws.Cells(r, 1).Font.Bold = True
    Application.StatusBar = False
End Sub

' One ICMP echo with the given TTL. True if anything answered (the
' destination, or a router saying the TTL expired / destination unreachable).
Private Function SendProbe(ByVal hIcmp As LongPtr, ByVal dest As Long, ByVal ttl As Long, _
        ByVal timeoutMs As Long, ByRef fromAddr As String, ByRef status As Long, ByRef rtt As Long, _
        Optional ByRef replyTtl As Long) As Boolean
    Dim opt(0 To 15) As Byte          ' IP_OPTION_INFORMATION: first byte is the TTL
    Dim payload(0 To 31) As Byte
    Dim reply(0 To 1023) As Byte      ' ICMP_ECHO_REPLY: Address @0, Status @4, RoundTripTime @8,
                                      ' Options.Ttl @16+pointer size
    Dim i As Long, n As Long, pReply As LongPtr

    opt(0) = CByte(ttl)
    For i = 0 To 31
        payload(i) = 97 + (i Mod 23)  ' "abcdefghijklmnopqrstuvw..." like Windows ping
    Next i
    pReply = VarPtr(reply(0))
    n = IcmpSendEcho(hIcmp, dest, VarPtr(payload(0)), 32, VarPtr(opt(0)), pReply, 1024, timeoutMs)

    status = ReadLong(pReply + 4)
    rtt = ReadLong(pReply + 8)
    fromAddr = reply(0) & "." & reply(1) & "." & reply(2) & "." & reply(3)
    replyTtl = reply(16 + PTR_SIZE)
    ' Some Windows versions return 0 for "TTL expired" but still fill in the reply
    SendProbe = (n > 0) Or (ReadLong(pReply) <> 0 And status <> IP_REQ_TIMED_OUT And status <> IP_SUCCESS)
End Function

' "10.1.2.3" -> IPAddr (network byte order, as IcmpSendEcho wants it)
Private Function IPv4ToNetLong(ByVal ip As String) As Long
    Dim parts() As String, b(0 To 3) As Byte, i As Long, v As Long
    parts = Split(ip, ".")
    For i = 0 To 3
        b(i) = CByte(parts(i))
    Next i
    CopyMemory v, VarPtr(b(0)), 4
    IPv4ToNetLong = v
End Function

'==============================================================================
' Subnet scan
'==============================================================================
Private Sub RunSubnetScan()
    Dim items As Collection, srcDesc As String, item As Variant
    Dim subnets As New Collection, info As SubnetInfo, totalAddr As Double, bad As String
    Dim wsaOk As Boolean, i As Long, summary As String, ports() As Long, nPorts As Long
    Dim secs As Double

    Set items = GetInputList("Subnet to scan, for example:" & vbCrLf & vbCrLf & _
                             "   10.1.4.0/24" & vbCrLf & _
                             "   10.1.4.0 255.255.255.0" & vbCrLf & vbCrLf & _
                             "(Several: separate them with commas.)", True, srcDesc)
    If items Is Nothing Then Exit Sub

    ' Check every subnet first so nothing starts if one is wrong
    For Each item In items
        info = ParseSubnet(CStr(item))
        If Not info.Ok Then
            bad = bad & vbCrLf & "   " & item & ": " & info.ErrorText
        ElseIf info.Hosts > SCAN_MAX_ADDRESSES Then
            bad = bad & vbCrLf & "   " & item & ": too big (" & Format$(info.Hosts, "#,##0") & _
                  " addresses - the limit is " & Format$(SCAN_MAX_ADDRESSES, "#,##0") & ", a /20)"
        Else
            subnets.Add CStr(item)
            totalAddr = totalAddr + info.Hosts
        End If
    Next item
    If Len(bad) > 0 Then
        MsgBox "These can't be scanned:" & bad, vbExclamation, APP_NAME
        Exit Sub
    End If

    ' Optional TCP ports
    If Not AskPorts("TCP ports to check on every address (leave blank for none)." & vbCrLf & vbCrLf & _
                    "Example: 22,80,443,3389 or 8000-8010", True, ports, nPorts) Then Exit Sub

    If totalAddr > SCAN_CONFIRM_ABOVE Then
        secs = (Int((totalAddr - 1) / SCAN_PARALLEL) + 1) * 2 * (SCAN_TIMEOUT_MS / 1000#) + totalAddr * 0.02
        If nPorts > 0 Then secs = secs + (Int((totalAddr * nPorts - 1) / PORT_PARALLEL) + 1) * (PORT_TIMEOUT_MS / 1000#)
        If MsgBox("Scan " & Format$(totalAddr, "#,##0") & " addresses" & _
                  IIf(nPorts > 0, " and " & nPorts & " port(s) on each", "") & "?" & vbCrLf & vbCrLf & _
                  "This takes roughly " & Format$(secs / 60#, "0.0") & " minute(s). Esc stops it.", _
                  vbOKCancel + vbQuestion, APP_NAME) <> vbOK Then Exit Sub
    End If

    Application.EnableCancelKey = xlErrorHandler
    On Error GoTo Handler
    wsaOk = WsaInit()
    For i = 1 To subnets.Count
        summary = summary & ScanOne(subnets(i), ports, nPorts) & vbCrLf & vbCrLf
    Next i
    Application.StatusBar = False
    MsgBox summary, vbInformation, APP_NAME
    GoTo Cleanup

Handler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    If Err.Number = 18 Then
        MsgBox "Scan stopped.", vbInformation, APP_NAME
    Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, APP_NAME
    End If

Cleanup:
    On Error Resume Next
    If wsaOk Then WSACleanup
    Application.StatusBar = False
    Application.EnableCancelKey = xlInterrupt
End Sub

' Scans one subnet onto its own sheet ("Scan 10.1.4.0_24"). Returns a short summary.
Private Function ScanOne(ByVal subnetText As String, ports() As Long, ByVal nPorts As Long) As String
    Dim info As SubnetInfo, cidr As String, n As Long, i As Long, j As Long, c As Long
    Dim ips() As String, todo() As Boolean, up() As Boolean, st() As Long, rtt() As Long, ttl() As Long
    Dim arp As Object, portRes As Variant, headers() As Variant, nCols As Long, out() As Variant
    Dim ptr As String, fwd As Collection, mac As String, openList As String, macStyle As Long
    Dim ws As Worksheet, lo As ListObject, rowRng As Range, result As String
    Dim nUp As Long, nMac As Long, nPtr As Long, nOpen As Long, nMismatch As Long, nFound As Long
    Dim present As Boolean

    info = ParseSubnet(subnetText)
    cidr = IPv4ToString(info.Network) & "/" & info.Prefix
    n = CLng(info.Hosts)
    ReDim ips(1 To n), todo(1 To n), up(1 To n), st(1 To n), rtt(1 To n), ttl(1 To n)
    For i = 1 To n
        ips(i) = IPv4ToString(info.FirstHost + i - 1)
        todo(i) = True
    Next i

    ' 1. Ping everything, then once more for anything that didn't answer
    PingMany ips, todo, SCAN_TIMEOUT_MS, up, st, rtt, ttl, "scanning " & cidr & " - ping"
    For i = 1 To n
        todo(i) = Not up(i)
    Next i
    PingMany ips, todo, SCAN_TIMEOUT_MS, up, st, rtt, ttl, "scanning " & cidr & " - 2nd ping"

    ' 2. MAC addresses: the pings fill the PC's ARP table for local subnets
    Set arp = ArpTable()
    macStyle = CLng(GetSetting(APP_NAME, "Mac", "Style", CStr(MAC_COLON)))

    ' 3. TCP ports
    If nPorts > 0 Then
        For i = 1 To n
            todo(i) = True
        Next i
        portRes = TcpScan(ips, todo, ports, nPorts, "scanning " & cidr & " - ports")
    End If

    ' 4. Reverse DNS, and build the table
    ReDim headers(0 To 7 + nPorts + IIf(nPorts > 0, 1, 0))
    headers(0) = SC_IP: headers(1) = H_PING: headers(2) = SC_LAT: headers(3) = H_TTL
    headers(4) = SC_MAC: headers(5) = SC_VENDOR: headers(6) = H_PTR: headers(7) = H_MATCH
    For j = 0 To nPorts - 1
        headers(8 + j) = PortLabel(ports(j))
    Next j
    If nPorts > 0 Then headers(8 + nPorts) = SC_OPEN
    nCols = UBound(headers) + 1
    ReDim out(1 To n, 1 To nCols)

    For i = 1 To n
        If i Mod 8 = 1 Then
            Application.StatusBar = APP_NAME & ": scanning " & cidr & " - DNS " & i & " of " & n & _
                                    "   -   press Esc to stop"
            DoEvents
        End If
        present = False
        out(i, 1) = ips(i)
        If up(i) Then
            nUp = nUp + 1: present = True
            out(i, 2) = "Up"
            If rtt(i) < 1 Then out(i, 3) = "<1" Else out(i, 3) = rtt(i)
            out(i, 4) = ttl(i)
        ElseIf st(i) = 0 Or st(i) = IP_REQ_TIMED_OUT Then
            out(i, 2) = "No reply"
        Else
            out(i, 2) = PingStatusText(st(i))
        End If

        If arp.Exists(ips(i)) Then
            mac = arp(ips(i))
            nMac = nMac + 1: present = True
            out(i, 5) = FormatMac(mac, macStyle)
            out(i, 6) = MacVendor(mac)
        End If

        ptr = PtrLookupDns(ips(i))
        If Len(ptr) > 0 Then
            nPtr = nPtr + 1
            out(i, 7) = ptr
            Set fwd = LookupAddresses(ptr)
            If InCollection(fwd, ips(i)) Then
                out(i, 8) = "Yes"
            ElseIf fwd.Count = 0 Then
                out(i, 8) = "No - PTR name does not resolve"
                nMismatch = nMismatch + 1
            Else
                out(i, 8) = "No - PTR name resolves to " & JoinCollection(fwd)
                nMismatch = nMismatch + 1
            End If
        End If

        If nPorts > 0 Then
            openList = ""
            For j = 0 To nPorts - 1
                out(i, 9 + j) = portRes(i, j)
                If portRes(i, j) = "Open" Then
                    If Len(openList) > 0 Then openList = openList & ", "
                    openList = openList & ports(j)
                End If
            Next j
            out(i, 9 + nPorts) = openList
            If Len(openList) > 0 Then nOpen = nOpen + 1: present = True
        End If
        If present Then nFound = nFound + 1
    Next i

    ' 5. Write it all in one go, then colour it
    Application.StatusBar = APP_NAME & ": writing results for " & cidr
    Set ws = NewResultSheet("Scan " & cidr)
    Set lo = StartResultTable(ws, "Subnet scan of " & cidr & "  (" & IPv4ToString(info.FirstHost) & _
                              " - " & IPv4ToString(info.LastHost) & ", " & Format$(n, "#,##0") & " addresses" & _
                              IIf(nPorts > 0, ", TCP " & JoinPorts(ports, nPorts), "") & ")", headers, n)
    Application.ScreenUpdating = False
    lo.DataBodyRange.Value = out
    For i = 1 To n
        Set rowRng = lo.DataBodyRange.Rows(i)
        If up(i) Then
            Good rowRng.Cells(1, 1): Good rowRng.Cells(1, 2)
        Else
            rowRng.Cells(1, 2).Font.Color = RGB(128, 128, 128)
            ' There (has a MAC or an open port) but not answering ping: yellow
            If Len(CStr(out(i, 5))) > 0 Then
                Warn rowRng.Cells(1, 1)
            ElseIf nPorts > 0 Then
                If Len(CStr(out(i, 9 + nPorts))) > 0 Then Warn rowRng.Cells(1, 1)
            End If
        End If
        If out(i, 8) = "Yes" Then
            Good rowRng.Cells(1, 8)
        ElseIf Left$(CStr(out(i, 8)), 2) = "No" Then
            Warn rowRng.Cells(1, 8)
        End If
        For j = 0 To nPorts - 1
            ColourPortCell rowRng.Cells(1, 9 + j)
        Next j
    Next i
    FinishTable lo, n
    Application.ScreenUpdating = True

    result = cidr & ":  " & nFound & " device(s) found  (" & nUp & " answer ping, " & nMac & " with a MAC" & _
             IIf(nPorts > 0, ", " & nOpen & " with open ports", "") & "),  " & nPtr & " with a PTR record" & _
             IIf(nMismatch > 0, ",  " & nMismatch & " PTR mismatch(es)", "") & "   (sheet """ & ws.Name & """)"
    ws.Cells(3, 1).Value = result
    ScanOne = result
End Function

' Pings many IPv4 addresses at once - up to SCAN_PARALLEL in flight - using
' IcmpSendEcho2 with an event per ping. Only ips(i) with todo(i) = True are
' pinged; up/st/rtt/ttl are filled in for those.
Private Sub PingMany(ips() As String, todo() As Boolean, ByVal timeoutMs As Long, _
        up() As Boolean, st() As Long, rtt() As Long, ttl() As Long, ByVal label As String)
    Dim i As Long, k As Long, g As Long, cnt As Long, used As Long, total As Long, done As Long
    Dim idx() As Long, hIcmp() As LongPtr, ev() As LongPtr, buf() As Byte
    Dim opt(0 To 15) As Byte, payload(0 To 31) As Byte, pBuf As LongPtr
    Dim ret As Long, lastErr As Long, nReplies As Long, status As Long, t0 As Double, waitMs As Long
    Dim errNum As Long, errDesc As String

    For i = LBound(ips) To UBound(ips)
        If todo(i) Then total = total + 1
    Next i
    If total = 0 Then Exit Sub

    ReDim idx(0 To SCAN_PARALLEL - 1), hIcmp(0 To SCAN_PARALLEL - 1), ev(0 To SCAN_PARALLEL - 1)
    ReDim buf(0 To ICMP_REPLY_BUF - 1, 0 To SCAN_PARALLEL - 1)   ' one column per ping, contiguous
    opt(0) = 128                                                 ' TTL
    For k = 0 To 31
        payload(k) = 97 + (k Mod 23)
    Next k
    For k = 0 To SCAN_PARALLEL - 1
        ev(k) = CreateEventA(0, 1, 0, 0)
        hIcmp(k) = IcmpCreateFile()
    Next k

    On Error GoTo Done
    i = LBound(ips)
    Do While i <= UBound(ips)
        Application.StatusBar = APP_NAME & ": " & label & " (" & done & " of " & total & ")   -   press Esc to stop"
        DoEvents

        ' Send a batch
        used = 0
        Do While i <= UBound(ips) And used < SCAN_PARALLEL
            If todo(i) Then
                idx(used) = i
                pBuf = VarPtr(buf(0, used))
                RtlZeroMemory pBuf, ICMP_REPLY_BUF
                ResetEvent ev(used)
                ret = IcmpSendEcho2(hIcmp(used), ev(used), 0, 0, IPv4ToNetLong(ips(i)), VarPtr(payload(0)), 32, _
                                    VarPtr(opt(0)), pBuf, ICMP_REPLY_BUF, timeoutMs)
                lastErr = Err.LastDllError
                ' Anything other than "pending" won't signal the event, so do it here
                If Not (ret = 0 And lastErr = 997) Then SetEvent ev(used)
                used = used + 1
            End If
            i = i + 1
        Loop
        If used = 0 Then Exit Do

        ' Wait for the batch (WaitForMultipleObjects takes 64 handles at a time)
        t0 = Timer
        For g = 0 To used - 1 Step 64
            cnt = used - g
            If cnt > 64 Then cnt = 64
            waitMs = timeoutMs + 1000 - CLng((Timer - t0) * 1000)
            If waitMs < 1 Then waitMs = 1
            WaitForMultipleObjects cnt, VarPtr(ev(g)), 1, waitMs
        Next g

        ' Read the replies (ICMP_ECHO_REPLY: Address @0, Status @4, RTT @8, Options.Ttl @16+ptr)
        For k = 0 To used - 1
            pBuf = VarPtr(buf(0, k))
            nReplies = IcmpParseReplies(pBuf, ICMP_REPLY_BUF)
            status = ReadLong(pBuf + 4)
            If nReplies > 0 And status = IP_SUCCESS Then
                up(idx(k)) = True
                st(idx(k)) = IP_SUCCESS
                rtt(idx(k)) = ReadLong(pBuf + 8)
                ttl(idx(k)) = buf(16 + PTR_SIZE, k)
            ElseIf ReadLong(pBuf) <> 0 And status <> IP_SUCCESS And status <> IP_REQ_TIMED_OUT Then
                st(idx(k)) = status          ' e.g. a router said "destination host unreachable"
            Else
                st(idx(k)) = IP_REQ_TIMED_OUT
            End If
        Next k
        done = done + used
    Loop

Done:
    errNum = Err.Number: errDesc = Err.Description
    For k = 0 To SCAN_PARALLEL - 1
        If hIcmp(k) <> 0 And hIcmp(k) <> -1 Then IcmpCloseHandle hIcmp(k)
        If ev(k) <> 0 Then CloseHandle ev(k)
    Next k
    If errNum <> 0 Then Err.Raise errNum, , errDesc
End Sub

' Tries a TCP connection to every ips(i) (with todo(i) = True) on every port,
' up to PORT_PARALLEL at once. Returns res(i, j) = "Open" / "Closed" /
' "Filtered" / "Unreachable" for ips(i), ports(j); "" where not checked.
Private Function TcpScan(ips() As String, todo() As Boolean, ports() As Long, ByVal nPorts As Long, _
        ByVal label As String) As Variant
    Dim res() As String, total As Long, done As Long, i As Long, j As Long, k As Long, m As Long
    Dim used As Long, socks() As LongPtr, pIp() As Long, pPort() As Long, pending() As Boolean, nPending As Long
    Dim sa(0 To 15) As Byte, ipb() As String, nb As Long, ret As Long, lastErr As Long
    Dim wf() As Byte, ef() As Byte, tv(0 To 1) As Long, setBytes As Long, cnt As Long
    Dim deadline As Double, remaining As Long, s As LongPtr, soErr As Long, optLen As Long
    Dim errNum As Long, errDesc As String

    ReDim res(LBound(ips) To UBound(ips), 0 To nPorts - 1)
    For i = LBound(ips) To UBound(ips)
        If todo(i) Then total = total + nPorts
    Next i
    If total = 0 Then TcpScan = res: Exit Function

    ReDim socks(0 To PORT_PARALLEL - 1), pIp(0 To PORT_PARALLEL - 1), pPort(0 To PORT_PARALLEL - 1)
    ReDim pending(0 To PORT_PARALLEL - 1)
    setBytes = PTR_SIZE + PORT_PARALLEL * PTR_SIZE   ' fd_set: count, then the sockets
    ReDim wf(0 To setBytes - 1), ef(0 To setBytes - 1)

    On Error GoTo Done
    i = LBound(ips): j = 0
    Do
        Application.StatusBar = APP_NAME & ": " & label & " (" & done & " of " & total & ")   -   press Esc to stop"
        DoEvents

        ' Start a batch of non-blocking connects
        used = 0: nPending = 0
        Do While i <= UBound(ips) And used < PORT_PARALLEL
            If Not todo(i) Then
                i = i + 1: j = 0
            Else
                pIp(used) = i: pPort(used) = j
                s = ws_socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
                socks(used) = s
                pending(used) = False
                If s = -1 Or s = 0 Then
                    res(i, j) = "Error"
                    socks(used) = 0
                Else
                    nb = 1
                    ioctlsocket s, FIONBIO, nb
                    Erase sa
                    sa(0) = AF_INET
                    sa(2) = ports(j) \ 256: sa(3) = ports(j) Mod 256
                    ipb = Split(ips(i), ".")
                    sa(4) = CByte(ipb(0)): sa(5) = CByte(ipb(1)): sa(6) = CByte(ipb(2)): sa(7) = CByte(ipb(3))
                    ret = ws_connect(s, VarPtr(sa(0)), 16)
                    If ret = 0 Then
                        res(i, j) = "Open"
                    Else
                        lastErr = WSAGetLastError()
                        If lastErr = WSAEWOULDBLOCK Then
                            pending(used) = True: nPending = nPending + 1
                        ElseIf lastErr = WSAECONNREFUSED Then
                            res(i, j) = "Closed"
                        Else
                            res(i, j) = "Unreachable"
                        End If
                    End If
                End If
                used = used + 1
                j = j + 1
                If j >= nPorts Then j = 0: i = i + 1
            End If
        Loop
        If used = 0 Then Exit Do

        ' Wait for them: writable = connected, exception = failed
        deadline = Timer + PORT_TIMEOUT_MS / 1000#
        Do While nPending > 0
            remaining = CLng((deadline - Timer) * 1000)
            If remaining <= 0 Then Exit Do
            cnt = 0
            For k = 0 To used - 1
                If pending(k) Then
                    CopyMemory wf(PTR_SIZE + cnt * PTR_SIZE), VarPtr(socks(k)), PTR_SIZE
                    CopyMemory ef(PTR_SIZE + cnt * PTR_SIZE), VarPtr(socks(k)), PTR_SIZE
                    cnt = cnt + 1
                End If
            Next k
            CopyMemory wf(0), VarPtr(cnt), 4
            CopyMemory ef(0), VarPtr(cnt), 4
            tv(0) = remaining \ 1000
            tv(1) = (remaining Mod 1000) * 1000
            ret = ws_select(0, 0, VarPtr(wf(0)), VarPtr(ef(0)), VarPtr(tv(0)))
            If ret <= 0 Then Exit Do

            cnt = ReadLong(VarPtr(wf(0)))
            For m = 0 To cnt - 1
                s = ReadPtr(VarPtr(wf(PTR_SIZE + m * PTR_SIZE)))
                For k = 0 To used - 1
                    If pending(k) And socks(k) = s Then
                        res(pIp(k), pPort(k)) = "Open"
                        pending(k) = False: nPending = nPending - 1
                        Exit For
                    End If
                Next k
            Next m
            cnt = ReadLong(VarPtr(ef(0)))
            For m = 0 To cnt - 1
                s = ReadPtr(VarPtr(ef(PTR_SIZE + m * PTR_SIZE)))
                For k = 0 To used - 1
                    If pending(k) And socks(k) = s Then
                        soErr = 0: optLen = 4
                        getsockopt s, SOL_SOCKET, SO_ERROR, soErr, optLen
                        If soErr = WSAECONNREFUSED Then
                            res(pIp(k), pPort(k)) = "Closed"
                        ElseIf soErr = WSAEHOSTUNREACH Or soErr = WSAENETUNREACH Then
                            res(pIp(k), pPort(k)) = "Unreachable"
                        Else
                            res(pIp(k), pPort(k)) = "Filtered"
                        End If
                        pending(k) = False: nPending = nPending - 1
                        Exit For
                    End If
                Next k
            Next m
            DoEvents
        Loop

        ' No answer in time = filtered (a firewall dropped it, or nothing there)
        For k = 0 To used - 1
            If pending(k) Then res(pIp(k), pPort(k)) = "Filtered"
            If socks(k) <> 0 Then closesocket socks(k): socks(k) = 0
        Next k
        done = done + used
    Loop

Done:
    errNum = Err.Number: errDesc = Err.Description
    For k = 0 To PORT_PARALLEL - 1
        If socks(k) <> 0 Then closesocket socks(k)
    Next k
    If errNum <> 0 Then Err.Raise errNum, , errDesc
    TcpScan = res
End Function

Private Sub ColourPortCell(ByVal c As Range)
    Select Case CStr(c.Value)
        Case "Open":                   Good c
        Case "Closed":                 c.Font.Color = RGB(128, 128, 128)
        Case "Filtered":               c.Font.Color = RGB(160, 160, 160)
        Case "Unreachable", "Error":   Bad c
    End Select
End Sub

' Asks which TCP ports to check (remembers the last answer).
' Returns False if the user cancels or the list is invalid.
Private Function AskPorts(ByVal prompt As String, ByVal allowBlank As Boolean, _
        ByRef ports() As Long, ByRef nPorts As Long) As Boolean
    Dim answer As String, errText As String
    answer = InputBox(prompt, APP_NAME & " - TCP ports", GetSetting(APP_NAME, "Ports", "Last", DEFAULT_PORTS))
    If StrPtr(answer) = 0 Then Exit Function                ' Cancel
    If Len(Trim$(answer)) = 0 Then
        If Not allowBlank Then Exit Function
        nPorts = 0
        AskPorts = True
        Exit Function
    End If
    If Not ParsePorts(answer, ports, nPorts, errText) Then
        MsgBox errText, vbExclamation, APP_NAME
        Exit Function
    End If
    SaveSetting APP_NAME, "Ports", "Last", answer
    AskPorts = True
End Function

' "22, 80,443 8000-8010" -> array of unique ports
Private Function ParsePorts(ByVal text As String, ByRef ports() As Long, ByRef nPorts As Long, _
        ByRef errText As String) As Boolean
    Dim part As Variant, t As String, a As Long, b As Long, p As Long, c As New Collection, d As Object
    Set d = CreateObject("Scripting.Dictionary")
    text = Replace(Replace(text, ";", ","), " ", ",")
    For Each part In Split(text, ",")
        t = Trim$(CStr(part))
        If Len(t) > 0 Then
            If InStr(t, "-") > 0 Then
                If Not IsAllDigits(Split(t, "-")(0)) Or Not IsAllDigits(Split(t, "-")(1)) Then
                    errText = "Not a port range: " & t: Exit Function
                End If
                a = CLng(Split(t, "-")(0)): b = CLng(Split(t, "-")(1))
            ElseIf IsAllDigits(t) And Len(t) <= 5 Then
                a = CLng(t): b = a
            Else
                errText = "Not a port number: " & t: Exit Function
            End If
            If a < 1 Or b > 65535 Or a > b Then errText = "Ports must be 1-65535: " & t: Exit Function
            For p = a To b
                If Not d.Exists(p) Then
                    d.Add p, True
                    c.Add p
                    If c.Count > MAX_PORTS Then errText = "Too many ports - the limit is " & MAX_PORTS & ".": Exit Function
                End If
            Next p
        End If
    Next part
    If c.Count = 0 Then errText = "No ports given.": Exit Function
    nPorts = c.Count
    ReDim ports(0 To nPorts - 1)
    For p = 1 To nPorts
        ports(p - 1) = c(p)
    Next p
    ParsePorts = True
End Function

Private Function JoinPorts(ports() As Long, ByVal nPorts As Long) As String
    Dim j As Long, s As String
    For j = 0 To nPorts - 1
        If j > 0 Then s = s & ","
        s = s & ports(j)
    Next j
    JoinPorts = s
End Function

' Column header for a port, e.g. "22 SSH"
Private Function PortLabel(ByVal port As Long) As String
    Dim nm As String
    Select Case port
        Case 21: nm = "FTP"
        Case 22: nm = "SSH"
        Case 23: nm = "Telnet"
        Case 25: nm = "SMTP"
        Case 53: nm = "DNS"
        Case 80: nm = "HTTP"
        Case 88: nm = "Kerberos"
        Case 110: nm = "POP3"
        Case 135: nm = "RPC"
        Case 139: nm = "NetBIOS"
        Case 143: nm = "IMAP"
        Case 389: nm = "LDAP"
        Case 443: nm = "HTTPS"
        Case 445: nm = "SMB"
        Case 465, 587: nm = "SMTP"
        Case 636: nm = "LDAPS"
        Case 830: nm = "NETCONF"
        Case 993: nm = "IMAPS"
        Case 1433: nm = "SQL"
        Case 1521: nm = "Oracle"
        Case 2049: nm = "NFS"
        Case 3306: nm = "MySQL"
        Case 3389: nm = "RDP"
        Case 5060: nm = "SIP"
        Case 5432: nm = "Postgres"
        Case 5900: nm = "VNC"
        Case 5985: nm = "WinRM"
        Case 5986: nm = "WinRM-S"
        Case 8080: nm = "HTTP-alt"
        Case 8443: nm = "HTTPS-alt"
        Case 9100: nm = "Printer"
    End Select
    PortLabel = "TCP " & port & IIf(Len(nm) > 0, " " & nm, "")
End Function

' Snapshot of the PC's ARP table: IPv4 address -> 12 hex digit MAC.
' Only devices on directly connected subnets appear (that's how ARP works).
Private Function ArpTable() As Object
    Dim d As Object, size As Long, buf() As Byte, n As Long, i As Long, base As Long
    Dim mac As String, k As Long, ip As String
    Set d = CreateObject("Scripting.Dictionary")
    Set ArpTable = d
    size = 0
    GetIpNetTable 0, size, 0
    If size <= 4 Then Exit Function
    size = size + 4096                      ' room in case the table grows
    ReDim buf(0 To size - 1)
    If GetIpNetTable(VarPtr(buf(0)), size, 0) <> 0 Then Exit Function
    ' MIB_IPNETTABLE: count @0, rows from @4, 24 bytes each:
    '   index @0, physAddrLen @4, physAddr(8) @8, addr @16, type @20 (2 = invalid)
    n = ReadLong(VarPtr(buf(0)))
    For i = 0 To n - 1
        base = 4 + i * 24
        If base + 24 > size Then Exit For
        If ReadLong(VarPtr(buf(base + 4))) = 6 And ReadLong(VarPtr(buf(base + 20))) <> 2 Then
            mac = ""
            For k = 0 To 5
                mac = mac & Right$("0" & Hex$(buf(base + 8 + k)), 2)
            Next k
            If mac <> "000000000000" And mac <> "FFFFFFFFFFFF" Then
                ip = buf(base + 16) & "." & buf(base + 17) & "." & buf(base + 18) & "." & buf(base + 19)
                d(ip) = mac
            End If
        End If
    Next i
End Function

'==============================================================================
' Port Check
'==============================================================================
Private Sub RunPortCheck()
    Dim items As Collection, srcDesc As String, ports() As Long, nPorts As Long
    Dim n As Long, i As Long, j As Long, ips() As String, todo() As Boolean, res As Variant
    Dim headers() As Variant, out() As Variant, ws As Worksheet, lo As ListObject, rowRng As Range
    Dim wsaOk As Boolean, openList As String, nOpenHosts As Long, nBad As Long

    Set items = GetInputList("Hostnames or IP addresses to port-check, separated by commas or spaces:", False, srcDesc)
    If items Is Nothing Then Exit Sub
    If Not AskPorts("TCP ports to check, for example:" & vbCrLf & vbCrLf & _
                    "22,80,443,3389   or   8000-8010", False, ports, nPorts) Then Exit Sub

    Application.EnableCancelKey = xlErrorHandler
    On Error GoTo Handler
    wsaOk = WsaInit()

    n = items.Count
    ReDim ips(1 To n), todo(1 To n)
    For i = 1 To n
        Application.StatusBar = APP_NAME & ": resolving " & items(i)
        ips(i) = FirstIPv4(items(i))
        todo(i) = (Len(ips(i)) > 0)
    Next i
    res = TcpScan(ips, todo, ports, nPorts, "port check")

    ReDim headers(0 To 2 + nPorts)
    headers(0) = H_HOST: headers(1) = SC_IP
    For j = 0 To nPorts - 1
        headers(2 + j) = PortLabel(ports(j))
    Next j
    headers(2 + nPorts) = SC_OPEN
    ReDim out(1 To n, 1 To nPorts + 3)
    For i = 1 To n
        out(i, 1) = items(i)
        If todo(i) Then
            out(i, 2) = ips(i)
            openList = ""
            For j = 0 To nPorts - 1
                out(i, 3 + j) = res(i, j)
                If res(i, j) = "Open" Then
                    If Len(openList) > 0 Then openList = openList & ", "
                    openList = openList & ports(j)
                End If
            Next j
            out(i, 3 + nPorts) = openList
            If Len(openList) > 0 Then nOpenHosts = nOpenHosts + 1
        Else
            out(i, 2) = "(not found)"
            nBad = nBad + 1
        End If
    Next i

    Set ws = NewResultSheet("Port Check")
    Set lo = StartResultTable(ws, "Port Check  -  TCP " & JoinPorts(ports, nPorts) & "  -  " & srcDesc, headers, n)
    Application.ScreenUpdating = False
    lo.DataBodyRange.Value = out
    For i = 1 To n
        Set rowRng = lo.DataBodyRange.Rows(i)
        If Not todo(i) Then Bad rowRng.Cells(1, 2)
        For j = 0 To nPorts - 1
            ColourPortCell rowRng.Cells(1, 3 + j)
        Next j
    Next i
    FinishTable lo, n
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox n & " host(s) checked: " & nOpenHosts & " with at least one open port" & _
           IIf(nBad > 0, ", " & nBad & " not found in DNS", "") & "." & vbCrLf & vbCrLf & _
           "Open = accepted the connection.  Closed = refused it.  " & _
           "Filtered = no answer (firewall, or nothing there).", vbInformation, APP_NAME
    GoTo Cleanup

Handler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    If Err.Number = 18 Then
        MsgBox "Port check stopped.", vbInformation, APP_NAME
    Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, APP_NAME
    End If

Cleanup:
    On Error Resume Next
    If wsaOk Then WSACleanup
    Application.StatusBar = False
    Application.EnableCancelKey = xlInterrupt
End Sub

' First IPv4 address for a host name or IP ("" if none)
Private Function FirstIPv4(ByVal host As String) As String
    Dim ip As String, item As Variant
    ip = NumericAddress(host)
    If Len(ip) > 0 Then
        If InStr(ip, ":") = 0 Then FirstIPv4 = ip
        Exit Function
    End If
    For Each item In LookupAddresses(host)
        If InStr(CStr(item), ":") = 0 Then FirstIPv4 = CStr(item): Exit Function
    Next item
End Function

'==============================================================================
' MAC addresses: formatting and vendor lookup
'==============================================================================
Private Sub RunMacFormat()
    Dim items As Collection, srcDesc As String, style As Long, n As Long, i As Long
    Dim out() As Variant, h As String, ws As Worksheet, lo As ListObject, nBad As Long

    Set items = GetInputList("MAC addresses to format, separated by commas:", True, srcDesc)
    If items Is Nothing Then Exit Sub
    style = AskMacStyle()
    If style = 0 Then Exit Sub

    n = items.Count
    ReDim out(1 To n, 1 To 3)
    For i = 1 To n
        out(i, 1) = items(i)
        h = MacHex(items(i))
        If Len(h) = 0 Then
            out(i, 2) = "Not a MAC address"
            nBad = nBad + 1
        Else
            out(i, 2) = FormatMac(h, style)
            out(i, 3) = MacVendor(h)
        End If
    Next i

    On Error GoTo Handler
    Set ws = NewResultSheet("MAC Addresses")
    Set lo = StartResultTable(ws, "MAC addresses  -  " & IIf(style = MAC_CISCO, "xxxx.xxxx.xxxx", "XX:XX:XX:XX:XX:XX") & _
                              "  -  " & srcDesc, Array("MAC (as entered)", "Formatted MAC", SC_VENDOR), n)
    lo.DataBodyRange.Value = out
    For i = 1 To n
        If Len(MacHex(CStr(out(i, 1)))) = 0 Then Bad lo.DataBodyRange.Cells(i, 2)
    Next i
    FinishTable lo, n
    If nBad > 0 Then MsgBox nBad & " value(s) weren't MAC addresses - see the red cells.", vbExclamation, APP_NAME
    Exit Sub
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, APP_NAME
End Sub

' Asks which MAC format to use (remembers the last answer). 0 = cancelled.
Private Function AskMacStyle() As Long
    Dim answer As String
    answer = InputBox("Which format?" & vbCrLf & vbCrLf & _
                      "   1  =  xxxx.xxxx.xxxx         (e.g. aabb.ccdd.eeff)" & vbCrLf & vbCrLf & _
                      "   2  =  XX:XX:XX:XX:XX:XX   (e.g. AA:BB:CC:DD:EE:FF)" & vbCrLf & vbCrLf & _
                      "Type 1 or 2:", APP_NAME & " - MAC format", _
                      GetSetting(APP_NAME, "Mac", "Style", CStr(MAC_CISCO)))
    If StrPtr(answer) = 0 Then Exit Function
    Select Case Trim$(answer)
        Case "1": AskMacStyle = MAC_CISCO
        Case "2": AskMacStyle = MAC_COLON
        Case Else
            MsgBox "Please type 1 or 2.", vbExclamation, APP_NAME
            Exit Function
    End Select
    SaveSetting APP_NAME, "Mac", "Style", CStr(AskMacStyle)
End Function

' Any common MAC notation -> 12 upper-case hex digits ("" if it isn't a MAC)
Private Function MacHex(ByVal s As String) As String
    Dim i As Long, c As String, h As String
    s = Trim$(s)
    For i = 1 To Len(s)
        c = UCase$(Mid$(s, i, 1))
        If (c >= "0" And c <= "9") Or (c >= "A" And c <= "F") Then
            h = h & c
        ElseIf InStr(":-. ", c) = 0 Then
            Exit Function
        End If
    Next i
    If Len(h) = 12 Then MacHex = h
End Function

' 12 hex digits -> MAC_CISCO "aabb.ccdd.eeff" or MAC_COLON "AA:BB:CC:DD:EE:FF"
' (style 3 = "AA-BB-CC-DD-EE-FF", 4 = "AABBCCDDEEFF")
Private Function FormatMac(ByVal h As String, ByVal style As Long) As String
    Select Case style
        Case MAC_CISCO
            h = LCase$(h)
            FormatMac = Mid$(h, 1, 4) & "." & Mid$(h, 5, 4) & "." & Mid$(h, 9, 4)
        Case 3
            FormatMac = Mid$(h, 1, 2) & "-" & Mid$(h, 3, 2) & "-" & Mid$(h, 5, 2) & "-" & _
                        Mid$(h, 7, 2) & "-" & Mid$(h, 9, 2) & "-" & Mid$(h, 11, 2)
        Case 4
            FormatMac = h
        Case Else
            FormatMac = Mid$(h, 1, 2) & ":" & Mid$(h, 3, 2) & ":" & Mid$(h, 5, 2) & ":" & _
                        Mid$(h, 7, 2) & ":" & Mid$(h, 9, 2) & ":" & Mid$(h, 11, 2)
    End Select
End Function

' Vendor from the first 3 bytes (OUI), using the IEEE list installed with NetTools.
Private Function MacVendor(ByVal h As String) As String
    Dim oui As Object
    If Len(h) <> 12 Then Exit Function
    ' Second-lowest bit of the first byte set = locally administered (random/private) MAC
    If (CLng("&H" & Left$(h, 2)) And 2) <> 0 Then
        MacVendor = "(private / random MAC)"
        Exit Function
    End If
    Set oui = OuiList()
    If oui.Count = 0 Then
        MacVendor = ""
    ElseIf oui.Exists(Left$(h, 6)) Then
        MacVendor = oui(Left$(h, 6))
    Else
        MacVendor = "(unknown)"
    End If
End Function

' Loads oui.txt ("AABBCC<tab>Vendor" per line) once. Looks next to the add-in,
' then in %APPDATA%\NetTools (where the MSI puts it).
Private Function OuiList() As Object
    Dim fso As Object, path As Variant, st As Object, txt As String, line As Variant, l As String
    If Not mOuiTried Then
        mOuiTried = True
        Set mOui = CreateObject("Scripting.Dictionary")
        On Error Resume Next
        Set fso = CreateObject("Scripting.FileSystemObject")
        For Each path In Array(ThisWorkbook.Path & "\oui.txt", Environ$("APPDATA") & "\NetTools\oui.txt")
            If fso.FileExists(CStr(path)) Then
                Set st = CreateObject("ADODB.Stream")
                st.Type = 2
                st.Charset = "utf-8"
                st.Open
                st.LoadFromFile CStr(path)
                txt = st.ReadText
                st.Close
                For Each line In Split(txt, vbLf)
                    l = Replace(CStr(line), vbCr, "")
                    If Len(l) > 7 Then
                        If Mid$(l, 7, 1) = vbTab Then mOui(UCase$(Left$(l, 6))) = Mid$(l, 8)
                    End If
                Next line
                Exit For
            End If
        Next path
        On Error GoTo 0
    End If
    Set OuiList = mOui
End Function

' Reverse lookup straight from DNS (DnsQuery PTR) - quick even when there is
' no record, unlike getnameinfo which can fall back to NetBIOS/LLMNR.
Private Function PtrLookupDns(ByVal ip As String) As String
    Dim parts() As String, qname As String, pRec As LongPtr, p As LongPtr
    parts = Split(ip, ".")
    If UBound(parts) <> 3 Then Exit Function
    qname = parts(3) & "." & parts(2) & "." & parts(1) & "." & parts(0) & ".in-addr.arpa"
    If DnsQuery_A(qname, DNS_TYPE_PTR, 0, 0, pRec, 0) = 0 Then
        ' DNS_RECORD: pNext, pName, wType (Integer) ... Data.PTR.pNameHost at 2*ptr+16
        p = pRec
        Do While p <> 0
            If ReadInt(p + 2 * PTR_SIZE) = DNS_TYPE_PTR Then
                PtrLookupDns = ReadStringA(ReadPtr(p + 2 * PTR_SIZE + 16))
                Exit Do
            End If
            p = ReadPtr(p)
        Loop
        DnsRecordListFree pRec, 1
    End If
End Function

Private Function ReadInt(ByVal p As LongPtr) As Integer
    Dim v As Integer
    CopyMemory v, p, 2
    ReadInt = v
End Function

Private Function ReadStringA(ByVal p As LongPtr) As String
    Dim n As Long, b() As Byte
    If p = 0 Then Exit Function
    n = lstrlenA(p)
    If n = 0 Then Exit Function
    ReDim b(0 To n - 1)
    CopyMemory b(0), p, n
    ReadStringA = StrConv(b, vbUnicode)
End Function

'==============================================================================
' Subnet calculator
'==============================================================================
Private Sub RunSubnetCalc()
    Dim items As Collection, srcDesc As String, ws As Worksheet, lo As ListObject
    Dim headers As Variant, i As Long, info As SubnetInfo, v() As Variant, nBad As Long
    Dim rowRng As Range

    Set items = GetInputList("Subnets to calculate, separated by commas, for example:" & vbCrLf & vbCrLf & _
                             "10.1.4.0/22, 192.168.1.10 255.255.255.0, 172.16.0.0/12", True, srcDesc)
    If items Is Nothing Then Exit Sub

    headers = Array(S_INPUT, S_CIDR, S_NET, S_MASK, S_WILD, S_FIRST, S_LAST, S_BCAST, S_HOSTS, S_TOTAL)
    On Error GoTo Handler
    Set ws = NewResultSheet("Subnets")
    Set lo = StartResultTable(ws, "Subnet Calculator  -  " & srcDesc, headers, items.Count)

    For i = 1 To items.Count
        info = ParseSubnet(items(i))
        ReDim v(0 To UBound(headers))
        v(0) = items(i)
        If info.Ok Then
            v(1) = IPv4ToString(info.Network) & "/" & info.Prefix
            v(2) = IPv4ToString(info.Network)
            v(3) = IPv4ToString(info.Mask)
            v(4) = IPv4ToString(4294967295# - info.Mask)
            v(5) = IPv4ToString(info.FirstHost)
            v(6) = IPv4ToString(info.LastHost)
            v(7) = IPv4ToString(info.Broadcast)
            v(8) = info.Hosts
            v(9) = info.Broadcast - info.Network + 1
        Else
            v(1) = info.ErrorText
        End If
        Set rowRng = lo.DataBodyRange.Rows(i)
        rowRng.Value = v
        If Not info.Ok Then
            nBad = nBad + 1
            Bad rowRng.Cells(1, 2)
        End If
    Next i
    FinishTable lo, items.Count

    If nBad > 0 Then
        MsgBox items.Count - nBad & " subnet(s) calculated." & vbCrLf & nBad & _
               " couldn't be read - see the red cells.", vbExclamation, APP_NAME
    End If
    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, APP_NAME
End Sub

' Accepts "10.1.4.0/22", "10.1.4.7/255.255.252.0", "10.1.4.7 255.255.252.0".
' The address may be any host in the subnet; the network is worked out from it.
Private Function ParseSubnet(ByVal text As String) As SubnetInfo
    Dim info As SubnetInfo, p As Long, ipPart As String, maskPart As String
    Dim m As Double, block As Double

    text = Trim$(Replace(Replace(text, vbTab, " "), Chr$(160), " "))
    p = InStr(text, "/")
    If p = 0 Then p = InStr(text, " ")
    If p = 0 Then
        info.ErrorText = "Add a prefix or mask, e.g. 10.1.4.0/22"
        ParseSubnet = info
        Exit Function
    End If
    ipPart = Trim$(Left$(text, p - 1))
    maskPart = Trim$(Mid$(text, p + 1))

    If Not ParseIPv4(ipPart, info.Address) Then
        info.ErrorText = "Not an IPv4 address: " & ipPart
        ParseSubnet = info
        Exit Function
    End If

    If IsAllDigits(maskPart) And Len(maskPart) <= 2 Then
        info.Prefix = CLng(maskPart)
        If info.Prefix > 32 Then
            info.ErrorText = "Prefix must be 0-32"
            ParseSubnet = info
            Exit Function
        End If
    ElseIf ParseIPv4(maskPart, m) Then
        info.Prefix = MaskToPrefix(m)
        If info.Prefix < 0 Then
            info.ErrorText = "Not a valid subnet mask: " & maskPart
            ParseSubnet = info
            Exit Function
        End If
    Else
        info.ErrorText = "Not a prefix or subnet mask: " & maskPart
        ParseSubnet = info
        Exit Function
    End If

    block = 2 ^ (32 - info.Prefix)
    info.Mask = 4294967296# - block
    info.Network = Int(info.Address / block) * block
    info.Broadcast = info.Network + block - 1
    Select Case info.Prefix
        Case 32         ' single host
            info.FirstHost = info.Network: info.LastHost = info.Network: info.Hosts = 1
        Case 31         ' point-to-point link (RFC 3021): both addresses usable
            info.FirstHost = info.Network: info.LastHost = info.Broadcast: info.Hosts = 2
        Case Else
            info.FirstHost = info.Network + 1: info.LastHost = info.Broadcast - 1: info.Hosts = block - 2
    End Select
    info.Ok = True
    ParseSubnet = info
End Function

Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i
    IsAllDigits = True
End Function

' "10.1.2.3" -> 167838211. False if it isn't a dotted-quad IPv4 address.
Private Function ParseIPv4(ByVal s As String, ByRef value As Double) As Boolean
    Dim parts() As String, i As Long, n As Long
    parts = Split(Trim$(s), ".")
    If UBound(parts) <> 3 Then Exit Function
    value = 0
    For i = 0 To 3
        If Not IsAllDigits(parts(i)) Or Len(parts(i)) > 3 Then Exit Function
        n = CLng(parts(i))
        If n > 255 Then Exit Function
        value = value * 256 + n
    Next i
    ParseIPv4 = True
End Function

Private Function IPv4ToString(ByVal v As Double) As String
    Dim a As Double, b As Double, c As Double, d As Double
    a = Int(v / 16777216#): v = v - a * 16777216#
    b = Int(v / 65536#): v = v - b * 65536#
    c = Int(v / 256#): d = v - c * 256#
    IPv4ToString = a & "." & b & "." & c & "." & d
End Function

' 255.255.252.0 -> 22. -1 if the mask isn't contiguous.
Private Function MaskToPrefix(ByVal m As Double) As Long
    Dim n As Long
    For n = 0 To 32
        If 4294967296# - 2 ^ (32 - n) = m Then MaskToPrefix = n: Exit Function
    Next n
    MaskToPrefix = -1
End Function

'==============================================================================
' Worksheet functions
'==============================================================================

' =NT_RESOLVE("server1")  ->  "10.1.2.3"   (all addresses, comma separated)
Public Function NT_RESOLVE(ByVal host As String) As String
    Dim addrs As Collection
    host = CleanHost(host)
    If Len(host) = 0 Then Exit Function
    If Not WsaInit() Then NT_RESOLVE = "#WSA error": Exit Function
    Set addrs = LookupAddresses(host)
    If addrs.Count = 0 Then NT_RESOLVE = "(not found)" Else NT_RESOLVE = JoinCollection(addrs)
    WSACleanup
End Function

' =NT_PTR("10.1.2.3")  ->  "server1.corp.example.com"
' A hostname is resolved first and its first address looked up.
Public Function NT_PTR(ByVal hostOrIp As String) As String
    Dim ip As String, addrs As Collection
    hostOrIp = CleanHost(hostOrIp)
    If Len(hostOrIp) = 0 Then Exit Function
    If Not WsaInit() Then NT_PTR = "#WSA error": Exit Function
    ip = NumericAddress(hostOrIp)
    If Len(ip) = 0 Then
        Set addrs = LookupAddresses(hostOrIp)
        If addrs.Count > 0 Then ip = addrs(1)
    End If
    If Len(ip) = 0 Then
        NT_PTR = "(not found)"
    Else
        ip = ReverseLookup(ip)
        If Len(ip) = 0 Then NT_PTR = "(no PTR record)" Else NT_PTR = ip
    End If
    WSACleanup
End Function

' =NT_PING("server1")  ->  "Up (4/4, 3 ms)" / "Timed out" / "Not resolved"
Public Function NT_PING(ByVal host As String) As String
    Dim ip As String, addrs As Collection, lat As Variant, ttl As Variant, status As String
    Dim received As Long, sent As Long
    host = CleanHost(host)
    If Len(host) = 0 Then Exit Function
    If Not WsaInit() Then NT_PING = "#WSA error": Exit Function
    ip = NumericAddress(host)
    If Len(ip) = 0 Then
        Set addrs = LookupAddresses(host)
        If addrs.Count > 0 Then ip = addrs(1)
    End If
    WSACleanup
    If Len(ip) = 0 Then NT_PING = "Not resolved": Exit Function
    status = PingHost(ip, lat, ttl, received, sent)
    If received > 0 Then
        NT_PING = status & " (" & received & "/" & sent & ", " & lat & " ms)"
    Else
        NT_PING = status
    End If
End Function

' =NT_SUBNET("10.1.4.7/22")              -> "10.1.4.0/22 (10.1.4.1 - 10.1.7.254, 1022 hosts)"
' =NT_SUBNET("10.1.4.7/22", "network")   -> "10.1.4.0"
' field: cidr, network, mask, wildcard, first, last, broadcast, hosts, range, prefix
Public Function NT_SUBNET(ByVal subnet As String, Optional ByVal field As String = "") As Variant
    Dim info As SubnetInfo
    info = ParseSubnet(subnet)
    If Not info.Ok Then NT_SUBNET = info.ErrorText: Exit Function
    Select Case LCase$(Trim$(field))
        Case "":                       NT_SUBNET = IPv4ToString(info.Network) & "/" & info.Prefix & _
                                           " (" & IPv4ToString(info.FirstHost) & " - " & _
                                           IPv4ToString(info.LastHost) & ", " & info.Hosts & " hosts)"
        Case "cidr":                   NT_SUBNET = IPv4ToString(info.Network) & "/" & info.Prefix
        Case "network", "net":         NT_SUBNET = IPv4ToString(info.Network)
        Case "mask", "netmask":        NT_SUBNET = IPv4ToString(info.Mask)
        Case "wildcard":               NT_SUBNET = IPv4ToString(4294967295# - info.Mask)
        Case "first":                  NT_SUBNET = IPv4ToString(info.FirstHost)
        Case "last":                   NT_SUBNET = IPv4ToString(info.LastHost)
        Case "broadcast", "bcast":     NT_SUBNET = IPv4ToString(info.Broadcast)
        Case "hosts", "usable":        NT_SUBNET = info.Hosts
        Case "range":                  NT_SUBNET = IPv4ToString(info.FirstHost) & " - " & IPv4ToString(info.LastHost)
        Case "prefix":                 NT_SUBNET = info.Prefix
        Case Else:                     NT_SUBNET = "Unknown field: " & field
    End Select
End Function

' =NT_INSUBNET("10.1.5.20", "10.1.4.0/22")  -> TRUE
Public Function NT_INSUBNET(ByVal ip As String, ByVal subnet As String) As Variant
    Dim info As SubnetInfo, v As Double
    info = ParseSubnet(subnet)
    If Not info.Ok Then NT_INSUBNET = info.ErrorText: Exit Function
    If Not ParseIPv4(CleanHost(ip), v) Then NT_INSUBNET = "Not an IPv4 address: " & ip: Exit Function
    NT_INSUBNET = (v >= info.Network And v <= info.Broadcast)
End Function

' =NT_MAC("aa-bb-cc-dd-ee-ff", "cisco")  -> "aabb.ccdd.eeff"
' style: "cisco" (or 1), "colon" (or 2, the default), "dash", "bare"
Public Function NT_MAC(ByVal mac As String, Optional ByVal style As Variant = "colon") As String
    Dim h As String, st As Long
    h = MacHex(mac)
    If Len(h) = 0 Then NT_MAC = "Not a MAC address": Exit Function
    Select Case LCase$(Trim$(CStr(style)))
        Case "cisco", "1", "dot", "dotted": st = MAC_CISCO
        Case "dash", "hyphen", "windows":   st = 3
        Case "bare", "plain", "none":       st = 4
        Case Else:                           st = MAC_COLON
    End Select
    NT_MAC = FormatMac(h, st)
End Function

' =NT_VENDOR("00:1b:54:12:34:56")  -> "Cisco Systems, Inc"
Public Function NT_VENDOR(ByVal mac As String) As String
    Dim h As String
    h = MacHex(mac)
    If Len(h) = 0 Then NT_VENDOR = "Not a MAC address": Exit Function
    h = MacVendor(h)
    If Len(h) = 0 Then NT_VENDOR = "(vendor list not installed)" Else NT_VENDOR = h
End Function
