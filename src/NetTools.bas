Attribute VB_Name = "NetTools"
'==============================================================================
' NetTools - Excel ribbon add-in for network engineers
'
' Check DNS : forward lookup, reverse (PTR) lookup, forward/reverse match
'             and ping for a column of IP addresses or hostnames.
' DNS Only  : the lookups without the ping.
' Ping Only : status, latency and TTL.
' Traceroute: hop-by-hop path to a host (3 probes per hop, like tracert).
' Subnet Scan: ping + reverse DNS (+ forward check) every address in a subnet.
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
Private Const APP_NAME As String = "NetTools"
Private Const APP_VERSION As String = "2.2.0"

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
Private Const SCAN_TIMEOUT_MS As Long = 500      ' one ping per address when scanning
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

Private mWmi As Object

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
        "2. Click Check DNS, DNS Only, Ping Only, Traceroute, Subnet Scan" & vbCrLf & _
        "   or Subnet Calc." & vbCrLf & _
        "3. Results go on their own sheet as a table (rerunning overwrites" & vbCrLf & _
        "   that sheet) - the data you selected is never changed." & vbCrLf & _
        "   Press Esc to stop a long run." & vbCrLf & vbCrLf & _
        "A header row (e.g. 'Hostname' or 'IP Address') is detected and skipped." & vbCrLf & _
        "Each traceroute and each subnet scan gets its own sheet." & vbCrLf & vbCrLf & _
        "Worksheet functions: =NT_RESOLVE(A2)  =NT_PTR(A2)  =NT_PING(A2)" & vbCrLf & _
        "   =NT_SUBNET(A2, ""network"")  =NT_INSUBNET(B2, A2)" & vbCrLf & vbCrLf & _
        "Ping: " & PING_COUNT & " echo requests per host, " & PING_TIMEOUT_MS & " ms timeout each.", _
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
    Dim t0 As Double, wsaOk As Boolean, msg As String

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

    For i = 1 To total
        host = items(i)
        Application.StatusBar = APP_NAME & ": " & host & "  (" & i & " of " & total & ")   -   press Esc to stop"

        res = CheckHost(host, doDns, doPing)
        Set rowRng = lo.DataBodyRange.Rows(i)
        rowRng.Value = ResultToRow(host, res, doDns, doPing)
        ColourRow rowRng, headers, res
        done = i

        If res.DnsFailed Then nDnsFail = nDnsFail + 1
        If doDns And Left$(res.Match, 2) = "No" Then nMismatch = nMismatch + 1
        If doPing Then
            If res.PingUp Then nUp = nUp + 1 Else nDown = nDown + 1
            If res.PingLoss Then nLoss = nLoss + 1
        End If
        DoEvents
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
             "prefix", "prefixes", "range", "ranges", "vlan", "vlans"
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
    Dim lat As Variant, ttl As Variant, received As Long

    res.Latency = Empty
    res.Ttl = Empty
    canonical = NumericAddress(host)

    If Len(canonical) > 0 Then
        ' ---- Input is an IP address ----
        res.InputType = "IP"
        res.Addresses = canonical
        res.FirstAddr = canonical
        If doDns Then
            res.Ptr = ReverseLookup(canonical)
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
            res.Ptr = ReverseLookup(res.FirstAddr)
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
        res.PingStatus = PingHost(res.FirstAddr, lat, ttl, received)
        res.Replies = received & "/" & PING_COUNT
        res.PingUp = (received > 0)
        res.PingLoss = (received > 0 And received < PING_COUNT)
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
' Ping (WMI Win32_PingStatus)
'==============================================================================
' Sends PING_COUNT echo requests (one at a time, like Windows ping).
' Returns "Up", "Up - packet loss", or the failure (e.g. "Timed out").
' latency = average round trip of the replies (ms), ttl = TTL of the last reply,
' received = number of replies.
Private Function PingHost(ByVal addr As String, ByRef latency As Variant, ByRef ttl As Variant, _
        ByRef received As Long) As String
    Dim pings As Object, reply As Object, code As Variant
    Dim i As Long, totalMs As Double, lastFailure As String
    latency = Empty
    ttl = Empty
    received = 0
    lastFailure = "No response"
    On Error GoTo Fail

    If mWmi Is Nothing Then
        Set mWmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    End If

    For i = 1 To PING_COUNT
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
    Next i

    If received = 0 Then
        PingHost = lastFailure
    Else
        latency = Round(totalMs / received, 0)
        If received = PING_COUNT Then PingHost = "Up" Else PingHost = "Up - packet loss"
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
    Dim hIcmp As LongPtr, wsaOk As Boolean, i As Long, summary As String

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
    If totalAddr > SCAN_CONFIRM_ABOVE Then
        If MsgBox("Scan " & Format$(totalAddr, "#,##0") & " addresses?" & vbCrLf & vbCrLf & _
                  "Addresses that answer take a moment each; ones that don't wait " & _
                  SCAN_TIMEOUT_MS & " ms. Worst case about " & _
                  Format$(totalAddr * (SCAN_TIMEOUT_MS / 1000#) / 60#, "0") & " minute(s). Esc stops it.", _
                  vbOKCancel + vbQuestion, APP_NAME) <> vbOK Then Exit Sub
    End If

    Application.EnableCancelKey = xlErrorHandler
    On Error GoTo Handler
    wsaOk = WsaInit()
    hIcmp = IcmpCreateFile()
    If hIcmp = 0 Or hIcmp = -1 Then
        MsgBox "Could not open the Windows ICMP service (IcmpCreateFile failed).", vbCritical, APP_NAME
        GoTo Cleanup
    End If

    For i = 1 To subnets.Count
        summary = summary & ScanOne(subnets(i), hIcmp) & vbCrLf
    Next i
    Application.StatusBar = False
    MsgBox summary, vbInformation, APP_NAME
    GoTo Cleanup

Handler:
    Application.StatusBar = False
    If Err.Number = 18 Then
        MsgBox "Scan stopped. The addresses scanned so far are on the sheet.", vbInformation, APP_NAME
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

' Scans one subnet onto its own sheet ("Scan 10.1.4.0_24"). Returns a one-line summary.
Private Function ScanOne(ByVal subnetText As String, ByVal hIcmp As LongPtr) As String
    Dim info As SubnetInfo, cidr As String, ws As Worksheet, lo As ListObject
    Dim headers As Variant, n As Long, i As Long, a As Double, ip As String
    Dim fromAddr As String, status As Long, rtt As Long, replyTtl As Long, answered As Boolean
    Dim ptr As String, fwd As Collection, v() As Variant, rowRng As Range
    Dim nUp As Long, nPtr As Long, nMismatch As Long, done As Long
    Dim result As String, errNum As Long, errDesc As String

    info = ParseSubnet(subnetText)
    cidr = IPv4ToString(info.Network) & "/" & info.Prefix
    n = CLng(info.Hosts)
    headers = Array(SC_IP, H_PING, SC_LAT, H_TTL, H_PTR, H_MATCH)

    Set ws = NewResultSheet("Scan " & cidr)
    Set lo = StartResultTable(ws, "Subnet scan of " & cidr & "  (" & IPv4ToString(info.FirstHost) & _
                              " - " & IPv4ToString(info.LastHost) & ", " & Format$(n, "#,##0") & _
                              " addresses, 1 ping each, " & SCAN_TIMEOUT_MS & " ms timeout)", headers, n)
    On Error GoTo Stopped

    For i = 1 To n
        a = info.FirstHost + i - 1
        ip = IPv4ToString(a)
        If i Mod 4 = 1 Then
            Application.StatusBar = APP_NAME & ": scanning " & cidr & " - " & ip & "  (" & i & " of " & n & _
                                    ", " & nUp & " up)   -   press Esc to stop"
        End If

        ReDim v(0 To 5)
        v(0) = ip
        answered = SendProbe(hIcmp, IPv4ToNetLong(ip), 128, SCAN_TIMEOUT_MS, fromAddr, status, rtt, replyTtl)
        If answered And status = IP_SUCCESS Then
            nUp = nUp + 1
            v(1) = "Up"
            If rtt < 1 Then v(2) = "<1" Else v(2) = rtt
            v(3) = replyTtl
        ElseIf answered Then
            v(1) = PingStatusText(status)
        Else
            v(1) = "No reply"
        End If

        ptr = PtrLookupDns(ip)
        If Len(ptr) > 0 Then
            nPtr = nPtr + 1
            v(4) = ptr
            Set fwd = LookupAddresses(ptr)
            If InCollection(fwd, ip) Then
                v(5) = "Yes"
            ElseIf fwd.Count = 0 Then
                v(5) = "No - PTR name does not resolve"
                nMismatch = nMismatch + 1
            Else
                v(5) = "No - PTR name resolves to " & JoinCollection(fwd)
                nMismatch = nMismatch + 1
            End If
        Else
            v(4) = ""
            v(5) = ""
        End If

        Set rowRng = lo.DataBodyRange.Rows(i)
        rowRng.Value = v
        If v(1) = "Up" Then
            Good rowRng.Cells(1, 1): Good rowRng.Cells(1, 2)
        Else
            rowRng.Cells(1, 2).Font.Color = RGB(128, 128, 128)
        End If
        If v(5) = "Yes" Then
            Good rowRng.Cells(1, 6)
        ElseIf Left$(CStr(v(5)), 2) = "No" Then
            Warn rowRng.Cells(1, 6)
        End If
        done = i
        DoEvents
    Next i

    FinishTable lo, done
    result = cidr & ":  " & nUp & " of " & n & " up,  " & nPtr & " with a PTR record" & _
              IIf(nMismatch > 0, ",  " & nMismatch & " PTR mismatch(es)", "") & _
              "   (sheet """ & ws.Name & """)"
    ws.Cells(3, 1).Value = result
    ScanOne = result
    Exit Function

Stopped:
    ' Esc (or an error): keep what's done, then let the caller report it
    errNum = Err.Number: errDesc = Err.Description
    FinishTable lo, done
    ws.Cells(3, 1).Value = "Stopped after " & done & " of " & n & " addresses: " & nUp & " up, " & nPtr & " with a PTR record."
    Err.Raise errNum, , errDesc
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
    Dim received As Long
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
    status = PingHost(ip, lat, ttl, received)
    If received > 0 Then
        NT_PING = status & " (" & received & "/" & PING_COUNT & ", " & lat & " ms)"
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
