Option Explicit

' Set to False before deploying to server
Const DEBUG_MODE As Boolean = True

Private g_wsh As Object

Private Sub WriteLog(ByVal message As String)
    Dim logDir As String
    Dim logFile As String
    Dim fNum As Integer

    logDir = ThisWorkbook.Path & "\Logbook"
    If Dir(logDir, vbDirectory) = "" Then MkDir logDir

    logFile = logDir & "\" & Format(Date, "dd_mm_yyyy") & ".txt"
    fNum = FreeFile
    Open logFile For Append As #fNum
    Print #fNum, "[" & Format(Now, "dd/mm/yyyy hh:mm:ss") & "] " & message
    Close #fNum
End Sub

Private Function CleanSheetName(ByVal sName As String) As String
    Dim badChars As Variant
    Dim i As Integer
    Dim tmp As String

    tmp = sName
    badChars = Array(":", "\", "/", "?", "*", "[", "]")
    For i = LBound(badChars) To UBound(badChars)
        tmp = Replace(tmp, badChars(i), "_")
    Next i
    If Trim(tmp) = "" Then tmp = "Unnamed"
    If Len(tmp) > 31 Then tmp = Left(tmp, 31)
    CleanSheetName = tmp
End Function

Public Sub CreateAndScanSubnets()
    On Error GoTo ErrorHandler

    Dim wsOzet   As Worksheet
    Dim wsNew    As Worksheet
    Dim lastRow  As Long
    Dim i        As Long
    Dim rawName  As String
    Dim nameBase As String
    Dim safeName As String
    Dim cidr     As String
    Dim gwIP     As String
    Dim fwName   As String
    Dim backupDir  As String
    Dim backupFile As String

    backupDir = ThisWorkbook.Path & "\Backups"
    If Dir(backupDir, vbDirectory) = "" Then MkDir backupDir

    Dim logDir As String
    logDir = ThisWorkbook.Path & "\Logbook"
    If Dir(logDir, vbDirectory) = "" Then MkDir logDir

    backupFile = backupDir & "\" & Format(Now, "dd_mm_yyyy_hh-mm-ss") & ".xlsm"
    ThisWorkbook.SaveCopyAs backupFile
    WriteLog "=========================================="
    WriteLog "Macro started. Backup: " & backupFile

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Set g_wsh = CreateObject("WScript.Shell")

    Set wsOzet = Nothing
    On Error Resume Next
    Set wsOzet = ThisWorkbook.Sheets("Ozet")
    On Error GoTo ErrorHandler

    If wsOzet Is Nothing Then
        WriteLog "ERROR: 'Ozet' sheet not found."
        If DEBUG_MODE Then MsgBox "ERROR: 'Ozet' sheet not found!", vbCritical
        GoTo Cleanup
    End If

    lastRow = wsOzet.Cells(wsOzet.Rows.Count, "A").End(xlUp).row
    WriteLog "Total subnets to scan: " & (lastRow - 1)

    For i = 2 To lastRow
        rawName = Trim(wsOzet.Cells(i, 1).Value)
        cidr = Trim(wsOzet.Cells(i, 2).Value)
        gwIP = Trim(wsOzet.Cells(i, 4).Value)
        fwName = Trim(wsOzet.Cells(i, 5).Value)

        If rawName <> "" And cidr <> "" Then
            If InStr(rawName, "(") > 0 Then
                nameBase = Left(rawName, InStr(rawName, "(") - 1)
            Else
                nameBase = rawName
            End If

            safeName = CleanSheetName(Trim(nameBase))
            Set wsNew = PrepareSheet(safeName)

            Application.StatusBar = "[" & (i - 1) & "/" & (lastRow - 1) & "] Scanning: " & safeName
            WriteLog "Scanning: " & safeName & " | " & cidr

            Call ProcessSingleSubnet(wsNew, rawName, cidr, gwIP, fwName)

            ThisWorkbook.Save
            WriteLog "Saved: " & safeName
        End If
    Next i

    Application.StatusBar = "Calculating occupancy rates..."
    Call UpdateOccupancyRates
    ThisWorkbook.Save

    WriteLog "All done."
    If DEBUG_MODE Then MsgBox "Scan completed successfully!", vbInformation

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False

    If Not g_wsh Is Nothing Then Set g_wsh = Nothing
    Exit Sub

ErrorHandler:
    WriteLog "CRITICAL ERROR | Code: " & Err.Number & " | Description: " & Err.Description
    If DEBUG_MODE Then MsgBox "AN ERROR OCCURRED!" & vbCrLf & "Code: " & Err.Number & vbCrLf & Err.Description, vbCritical
    Resume Cleanup
End Sub

Private Sub UpdateOccupancyRates()
    Dim wsOzet   As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow  As Long
    Dim i        As Long
    Dim rawName  As String
    Dim nameBase As String
    Dim safeName As String
    Dim totalIPs As Long
    Dim usedIPs  As Long
    Dim rate     As Double

    Set wsOzet = Nothing
    On Error Resume Next
    Set wsOzet = ThisWorkbook.Sheets("Ozet")
    On Error GoTo 0
    If wsOzet Is Nothing Then Exit Sub

    wsOzet.Range("C1").Value = "Occupancy"
    wsOzet.Range("C1").Font.Bold = True
    lastRow = wsOzet.Cells(wsOzet.Rows.Count, "A").End(xlUp).row

    For i = 2 To lastRow
        rawName = Trim(wsOzet.Cells(i, 1).Value)
        If rawName <> "" Then
            If InStr(rawName, "(") > 0 Then
                nameBase = Left(rawName, InStr(rawName, "(") - 1)
            Else
                nameBase = rawName
            End If
            safeName = CleanSheetName(Trim(nameBase))

            Set wsTarget = Nothing
            On Error Resume Next
            Set wsTarget = ThisWorkbook.Sheets(safeName)
            On Error GoTo 0

            If Not wsTarget Is Nothing Then
                totalIPs = Application.WorksheetFunction.CountA(wsTarget.Range("A6:A1048576"))
                usedIPs = Application.WorksheetFunction.CountIf(wsTarget.Range("B6:B1048576"), "Dolu")

                rate = IIf(totalIPs > 0, usedIPs / totalIPs, 0)

                wsOzet.Cells(i, 3).Value = rate
                wsOzet.Cells(i, 3).NumberFormat = "0.00%"
            Else
                wsOzet.Cells(i, 3).Value = "Sheet Missing"
            End If
        End If
    Next i
    WriteLog "Occupancy rates updated."
End Sub

Private Function PrepareSheet(sName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sName
    End If
    Set PrepareSheet = ws
End Function

Private Sub ProcessSingleSubnet(ws As Worksheet, _
                                sName As String, _
                                cidr As String, _
                                gwIP As String, _
                                targetFirewall As String)

    Dim ipStr    As String
    Dim prefix   As Integer
    Dim o1 As Integer, o2 As Integer, o3 As Integer, o4 As Integer
    Dim hosts     As Long
    Dim hostIndex As Long
    Dim addrLow   As Long
    Dim netLow    As Long
    Dim cur3 As Integer, cur4 As Integer
    Dim currentIP As String
    Dim dnsName   As String
    Dim isDolu    As Boolean
    Dim row       As Long
    Dim evCnt     As Integer
    Dim parts()   As String
    Dim tableRange As Range

    Dim dictData  As Object
    Set dictData = CreateObject("Scripting.Dictionary")

    Dim lastExRow As Long
    Dim r As Long
    Dim exIP As String
    Dim vals(4) As Variant

    lastExRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    If lastExRow >= 6 Then
        For r = 6 To lastExRow
            exIP = Trim(ws.Cells(r, 1).Value)
            If exIP <> "" Then
                vals(0) = ws.Cells(r, 4).Value
                vals(1) = ws.Cells(r, 5).Value
                vals(2) = ws.Cells(r, 6).Value
                vals(3) = ws.Cells(r, 7).Value
                vals(4) = ws.Cells(r, 9).Value
                dictData(exIP) = vals
            End If
        Next r
    End If

    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Cells.ClearContents
    ws.Cells.ClearFormats

    parts = Split(cidr, "/")
    If UBound(parts) <> 1 Then
        ws.Range("A1").Value = "Invalid CIDR: " & cidr
        WriteLog "  ERROR: Invalid CIDR -> " & cidr
        GoTo SubCleanup
    End If

    ipStr = Trim(parts(0))
    prefix = CInt(Trim(parts(1)))

    If Not ParseIP(ipStr, o1, o2, o3, o4) Then
        ws.Range("A1").Value = "Invalid IP: " & ipStr
        WriteLog "  ERROR: Invalid IP -> " & ipStr
        GoTo SubCleanup
    End If

    hosts = 2& ^ (32 - prefix)
    netLow = CLng(o3) * 256& + CLng(o4)

    ws.Range("A1").Value = sName
    ws.Range("A2").Value = cidr
    ws.Range("A3").Value = "Gateway : " & gwIP
    ws.Range("A1:A3").Font.Bold = True
    ws.Range("A1:A3").Font.Size = 11

    ws.Range("A5").Value = "IP"
    ws.Range("B5").Value = "Status"
    ws.Range("C5").Value = "Device"
    ws.Range("D5").Value = "Responsible"
    ws.Range("E5").Value = "Environment Type"
    ws.Range("F5").Value = "Classification"
    ws.Range("G5").Value = "Type of Asset"
    ws.Range("H5").Value = "Firewall"
    ws.Range("I5").Value = "Notes"

    With ws.Range("A5:I5")
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(255, 192, 0)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 20
    End With

    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B").ColumnWidth = 10
    ws.Columns("C").ColumnWidth = 35
    ws.Columns("D").ColumnWidth = 15
    ws.Columns("E").ColumnWidth = 18
    ws.Columns("F").ColumnWidth = 15
    ws.Columns("G").ColumnWidth = 15
    ws.Columns("H").ColumnWidth = 15
    ws.Columns("I").ColumnWidth = 25

    row = 6
    evCnt = 0

    For hostIndex = 0& To (hosts - 1&)
        addrLow = netLow + hostIndex
        cur3 = CInt(addrLow \ 256&)
        cur4 = CInt(addrLow Mod 256&)
        currentIP = o1 & "." & o2 & "." & cur3 & "." & cur4

        ws.Cells(row, 1).Value = currentIP

        evCnt = evCnt + 1
        If evCnt >= 10 Then
            DoEvents
            evCnt = 0
        End If

        If hostIndex = 0 Then
            ws.Cells(row, 2).Value = "Used"
            ws.Cells(row, 3).Value = "Network ID"
            ws.Cells(row, 4).Value = "Firewall Team"
            ws.Cells(row, 8).Value = targetFirewall

        ElseIf currentIP = gwIP Then
            isDolu = IsPingableClassic(currentIP)
            ws.Cells(row, 2).Value = IIf(isDolu, "Used", "Free")
            DoEvents
            dnsName = ResolveDNSClassic(currentIP)
            ws.Cells(row, 3).Value = IIf(dnsName <> "", dnsName & " (Gateway)", "Gateway")
            ws.Cells(row, 4).Value = "Firewall Team"
            ws.Cells(row, 8).Value = targetFirewall

        ElseIf hostIndex = (hosts - 1&) Then
            ws.Cells(row, 2).Value = "Free"
            ws.Cells(row, 3).Value = "Broadcast IP"

        Else
            isDolu = IsPingableClassic(currentIP)
            ws.Cells(row, 2).Value = IIf(isDolu, "Used", "Free")
            DoEvents
            dnsName = ResolveDNSClassic(currentIP)
            ws.Cells(row, 3).Value = dnsName
        End If

        If dictData.Exists(currentIP) Then
            Dim saved As Variant
            saved = dictData(currentIP)
            If ws.Cells(row, 4).Value = "" Then ws.Cells(row, 4).Value = saved(0)
            ws.Cells(row, 5).Value = saved(1)
            ws.Cells(row, 6).Value = saved(2)
            ws.Cells(row, 7).Value = saved(3)
            ws.Cells(row, 9).Value = saved(4)
        End If

        row = row + 1
    Next hostIndex

    Set tableRange = ws.Range("A6:I" & (row - 1))
    With tableRange
        .Interior.Color = RGB(255, 242, 204)
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Color = vbWhite
            .Weight = xlThick
        End With
    End With
    With ws.Range("A5:I5").Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Color = vbWhite
        .Weight = xlThick
    End With

SubCleanup:
    Set dictData = Nothing
End Sub

Private Function ParseIP(ip As String, ByRef o1 As Integer, ByRef o2 As Integer, ByRef o3 As Integer, ByRef o4 As Integer) As Boolean
    Dim parts() As String
    parts = Split(Trim(ip), ".")
    If UBound(parts) <> 3 Then ParseIP = False: Exit Function
    On Error GoTo ErrH
    o1 = CInt(parts(0)): o2 = CInt(parts(1))
    o3 = CInt(parts(2)): o4 = CInt(parts(3))
    If o1 < 0 Or o1 > 255 Or o2 < 0 Or o2 > 255 Or o3 < 0 Or o3 > 255 Or o4 < 0 Or o4 > 255 Then ParseIP = False: Exit Function
    ParseIP = True: Exit Function
ErrH:
    ParseIP = False
End Function

Private Function IsPingableClassic(ip As String) As Boolean
    Dim ret As Integer
    On Error GoTo ErrHandler
    If g_wsh Is Nothing Then Set g_wsh = CreateObject("WScript.Shell")
    ret = g_wsh.Run("cmd /c ping -n 1 -w 200 " & ip, 0, True)
    IsPingableClassic = (ret = 0)
    Exit Function
ErrHandler:
    IsPingableClassic = False
End Function

Private Function ResolveDNSClassic(ip As String) As String
    Dim execObj  As Object
    Dim line     As String
    Dim hostName As String
    Dim deadline As Date

    ResolveDNSClassic = ""
    On Error GoTo ErrH
    If g_wsh Is Nothing Then Set g_wsh = CreateObject("WScript.Shell")
    hostName = ""

    Set execObj = g_wsh.Exec("cmd /c nslookup -timeout=1 -retry=1 " & ip & " 2>nul")
    deadline = Now + TimeValue("00:00:05")

    Do While Not execObj.StdOut.AtEndOfStream
        If Now > deadline Then Exit Do
        line = execObj.StdOut.ReadLine
        If InStr(1, line, "name =", vbTextCompare) > 0 Then
            hostName = Trim(Mid(line, InStr(1, line, "name =", vbTextCompare) + 6))
        ElseIf InStr(1, line, "Name:", vbTextCompare) > 0 Then
            hostName = Trim(Mid(line, InStr(1, line, "Name:", vbTextCompare) + 5))
        End If
    Loop

    If Len(hostName) > 0 Then
        If Right(hostName, 1) = "." Then hostName = Left(hostName, Len(hostName) - 1)
    End If

    Do While Not execObj.StdErr.AtEndOfStream: execObj.StdErr.ReadLine: Loop
    On Error Resume Next: execObj.Terminate: On Error GoTo ErrH
    Set execObj = Nothing

    If hostName = "" Then
        Set execObj = g_wsh.Exec("cmd /c ping -a -n 1 -w 200 " & ip & " 2>nul")
        deadline = Now + TimeValue("00:00:05")

        Do While Not execObj.StdOut.AtEndOfStream
            If Now > deadline Then Exit Do
            line = execObj.StdOut.ReadLine
            If InStr(line, " [") > 0 And (InStr(1, line, "ping", vbTextCompare) > 0) Then
                Dim parts() As String: parts = Split(line, " [")
                Dim tempName As String: tempName = Trim(parts(0))
                If Left(tempName, 8) = "Pinging " Then tempName = Mid(tempName, 9)
                If tempName <> "" And tempName <> ip Then hostName = tempName
                Exit Do
            End If
        Loop

        Do While Not execObj.StdErr.AtEndOfStream: execObj.StdErr.ReadLine: Loop
        On Error Resume Next: execObj.Terminate: On Error GoTo ErrH
        Set execObj = Nothing
    End If

    ResolveDNSClassic = hostName
    Exit Function

ErrH:
    ResolveDNSClassic = ""
    If Not execObj Is Nothing Then
        On Error Resume Next
        execObj.Terminate
        Set execObj = Nothing
    End If
End Function
