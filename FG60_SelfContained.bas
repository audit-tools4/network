Option Explicit

' FortiGate詳細設計書（ネットワーク／ルーティングシート）とconfigの照合ツール
' 標準モジュールへ、このファイル全体をそのまま貼り付けてください。
' Python、PowerShell、外部ライブラリ、外部通信は使用しません。

Private Const SHEET_DESIGN As String = "FG_設計正規化"
Private Const SHEET_CONFIG As String = "FG_Config正規化"
Private Const SHEET_RESULT As String = "FG_比較結果"
Private Const SHEET_LOG As String = "FG_実行ログ"
Private Const SECTION_INTERFACE As String = "system interface"
Private Const SECTION_STATIC_ROUTE As String = "router static"

Public Sub RunAllChecks()
    Dim designPath As Variant, configPath As Variant

    designPath = Application.GetOpenFilename( _
        "Excel,*.xlsx;*.xlsm;*.xls", , "詳細設計書を選択")
    If VarType(designPath) = vbBoolean Then Exit Sub

    configPath = Application.GetOpenFilename( _
        "FortiGate config,*.conf;*.txt;*.cfg;*.*", , "FortiGate configを選択")
    If VarType(configPath) = vbBoolean Then Exit Sub

    Dim oldScreen As Boolean, oldEvents As Boolean
    Dim oldAlerts As Boolean, oldCalculation As XlCalculation
    Dim oldSecurity As Long

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
    oldSecurity = Application.AutomationSecurity

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.AutomationSecurity = 3   ' msoAutomationSecurityForceDisable

    On Error GoTo ErrorHandler

    InitializeSheets
    LogMessage "INFO", "処理開始", "Design=" & CStr(designPath)
    ExtractNetworkSheet CStr(designPath)
    ExtractRoutingSheet CStr(designPath)
    ParseFortiGateConfig CStr(configPath)
    CompareNormalizedData
    FormatAllSheets

    ThisWorkbook.Worksheets(SHEET_RESULT).Activate
    MsgBox "照合が完了しました。FG_比較結果を確認してください。", vbInformation

CleanExit:
    Application.AutomationSecurity = oldSecurity
    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldAlerts
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Exit Sub

ErrorHandler:
    LogMessage "ERROR", "処理中断", Err.Number & ": " & Err.Description
    MsgBox "処理中にエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Private Sub ExtractRoutingSheet(ByVal workbookPath As String)
    Dim sourceBook As Workbook, sourceSheet As Worksheet, outputSheet As Worksheet
    Dim outputRow As Long, rowNumber As Long, blockEnd As Long
    Dim lastRow As Long, routeCount As Long, sectionEnd As Long
    Set outputSheet = ThisWorkbook.Worksheets(SHEET_DESIGN)

    On Error GoTo ErrorHandler
    Set sourceBook = Workbooks.Open( _
        Filename:=workbookPath, UpdateLinks:=0, ReadOnly:=True, _
        AddToMru:=False, IgnoreReadOnlyRecommended:=True)

    On Error Resume Next
    Set sourceSheet = sourceBook.Worksheets("ルーティング")
    On Error GoTo ErrorHandler
    If sourceSheet Is Nothing Then
        LogMessage "WARN", "ルーティング抽出", "ルーティングシートなし（スキップ）"
        GoTo CleanExit
    End If

    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "C").End(xlUp).Row
    sectionEnd = lastRow
    For rowNumber = 1 To lastRow
        If InStr(1, NormalizeText(ReadMerged(sourceSheet.Cells(rowNumber, "B"))), _
                 "BGP", vbTextCompare) > 0 Then
            sectionEnd = rowNumber - 1
            Exit For
        End If
    Next rowNumber

    outputRow = outputSheet.Cells(outputSheet.Rows.Count, 1).End(xlUp).Row + 1
    For rowNumber = 1 To sectionEnd
        If Left$(NormalizeText(ReadMerged(sourceSheet.Cells(rowNumber, "C"))), 5) = "VDOM【" Then
            blockEnd = FindNextVdomRow(sourceSheet, rowNumber, sectionEnd)
            ExtractStaticRouteBlock sourceSheet, rowNumber, blockEnd, _
                                    outputSheet, outputRow
            routeCount = routeCount + 1
        End If
    Next rowNumber

    LogMessage "INFO", "ルーティング抽出", _
        "StaticRoutes=" & routeCount & ", TotalDesignParameters=" & (outputRow - 2)

CleanExit:
    On Error Resume Next
    If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    On Error GoTo 0
    Exit Sub

ErrorHandler:
    Dim errorNumber As Long, errorDescription As String
    errorNumber = Err.Number
    errorDescription = Err.Description
    ResumeAfterClose sourceBook
    Err.Raise errorNumber, , errorDescription
End Sub

Private Function FindNextVdomRow(ByVal ws As Worksheet, ByVal startRow As Long, _
                                 ByVal sectionEnd As Long) As Long
    Dim rowNumber As Long
    For rowNumber = startRow + 1 To sectionEnd
        If Left$(NormalizeText(ReadMerged(ws.Cells(rowNumber, "C"))), 5) = "VDOM【" Then
            FindNextVdomRow = rowNumber - 1
            Exit Function
        End If
    Next rowNumber
    FindNextVdomRow = sectionEnd
End Function

Private Sub ExtractStaticRouteBlock(ByVal ws As Worksheet, ByVal startRow As Long, _
                                    ByVal endRow As Long, ByVal outputSheet As Worksheet, _
                                    ByRef outputRow As Long)
    Dim heading As String, vdom As String, destination As String
    Dim gateway As String, device As String, distance As String
    Dim comment As String, status As String, location As String, objectKey As String

    heading = NormalizeText(ReadMerged(ws.Cells(startRow, "C")))
    vdom = Replace(Replace(heading, "VDOM【", ""), "】", "")
    destination = NormalizeIpMask(GetRoutingLabelValue(ws, startRow, endRow, "宛先", location))
    gateway = GetRoutingLabelValue(ws, startRow, endRow, "ゲートウェイアドレス", location)
    device = GetRoutingLabelValue(ws, startRow, endRow, "インターフェース", location)
    distance = GetRoutingLabelValue(ws, startRow, endRow, _
                             "アドミニストレーティブ・ディスタンス", location)
    comment = GetRoutingLabelValue(ws, startRow, endRow, "コメント", location)
    status = MapRouteStatus(SelectedLabel(ws, _
                 FindLabelRow(ws, startRow, endRow, "ステータス")))

    If destination = "" Then Exit Sub
    objectKey = BuildStaticRouteObjectKey(destination, gateway, device)
    AddDesignParameter outputSheet, outputRow, vdom, SECTION_STATIC_ROUTE, _
        objectKey, "dst", destination, "ip-mask", "error", _
        FindRoutingLabelLocation(ws, startRow, endRow, "宛先"), "経路識別キーの一部"
    AddDesignParameter outputSheet, outputRow, vdom, SECTION_STATIC_ROUTE, _
        objectKey, "gateway", gateway, "exact", "error", _
        FindRoutingLabelLocation(ws, startRow, endRow, "ゲートウェイアドレス"), "経路識別キーの一部"
    AddDesignParameter outputSheet, outputRow, vdom, SECTION_STATIC_ROUTE, _
        objectKey, "device", device, "exact", "error", _
        FindRoutingLabelLocation(ws, startRow, endRow, "インターフェース"), "経路識別キーの一部"
    AddDesignParameter outputSheet, outputRow, vdom, SECTION_STATIC_ROUTE, _
        objectKey, "distance", distance, "exact", "default-review", _
        FindRoutingLabelLocation(ws, startRow, endRow, _
        "アドミニストレーティブ・ディスタンス"), "FortiOS既定値は10"
    AddDesignParameter outputSheet, outputRow, vdom, SECTION_STATIC_ROUTE, _
        objectKey, "comment", comment, "exact", "empty-is-match", _
        FindRoutingLabelLocation(ws, startRow, endRow, "コメント"), ""
    AddDesignParameter outputSheet, outputRow, vdom, SECTION_STATIC_ROUTE, _
        objectKey, "status", status, "exact", "default-review", _
        FindRoutingLabelLocation(ws, startRow, endRow, "ステータス"), "FortiOS既定値はenable"
End Sub

Private Function GetRoutingLabelValue(ByVal ws As Worksheet, ByVal startRow As Long, _
                                      ByVal endRow As Long, ByVal label As String, _
                                      ByRef location As String) As String
    Dim rowNumber As Long
    rowNumber = FindLabelRow(ws, startRow, endRow, label)
    If rowNumber > 0 Then
        location = CellLocation(ws, rowNumber, 14)
        GetRoutingLabelValue = NormalizeText(ReadMerged(ws.Cells(rowNumber, 14)))
    End If
End Function

Private Function FindRoutingLabelLocation(ByVal ws As Worksheet, _
                                          ByVal startRow As Long, _
                                          ByVal endRow As Long, _
                                          ByVal label As String) As String
    Dim rowNumber As Long
    rowNumber = FindLabelRow(ws, startRow, endRow, label)
    If rowNumber > 0 Then FindRoutingLabelLocation = CellLocation(ws, rowNumber, 14)
End Function

Private Function MapRouteStatus(ByVal value As String) As String
    Select Case value
        Case "有効化済み", "有効": MapRouteStatus = "enable"
        Case "無効化済み", "無効": MapRouteStatus = "disable"
    End Select
End Function

Private Function BuildStaticRouteObjectKey(ByVal destination As String, _
                                           ByVal gateway As String, _
                                           ByVal device As String) As String
    BuildStaticRouteObjectKey = NormalizeIpMask(destination) & "|" & _
                                LCase$(NormalizeText(gateway)) & "|" & _
                                LCase$(NormalizeText(device))
End Function

Public Sub InitializeSheets()
    PrepareSheet SHEET_DESIGN, Array( _
        "Device", "VDOM", "Section", "ObjectKey", "Parameter", _
        "DesignValue", "CompareMode", "MissingPolicy", "SourceLocation", "Note", "Key")

    PrepareSheet SHEET_CONFIG, Array( _
        "Device", "VDOM", "Section", "ObjectKey", "Parameter", _
        "ConfigValue", "ConfigLocation", "Key")

    PrepareSheet SHEET_RESULT, Array( _
        "判定", "VDOM", "Section", "ObjectKey", "Parameter", _
        "設計値", "Config値", "設計場所", "Config場所", "比較方法", "未記載時", "Key")

    PrepareSheet SHEET_LOG, Array("日時", "レベル", "処理", "内容")
End Sub

Private Sub PrepareSheet(ByVal sheetName As String, ByVal headers As Variant)
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(sheetName)
    ws.Cells.Clear

    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        WriteText ws.Cells(1, i + 1), CStr(headers(i))
    Next i

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(headers) + 1))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(23, 54, 93)
        .AutoFilter
    End With
End Sub

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function

Private Sub ExtractNetworkSheet(ByVal workbookPath As String)
    Dim sourceBook As Workbook, sourceSheet As Worksheet
    Dim outputSheet As Worksheet
    Set outputSheet = ThisWorkbook.Worksheets(SHEET_DESIGN)

    On Error GoTo ErrorHandler

    Set sourceBook = Workbooks.Open( _
        Filename:=workbookPath, UpdateLinks:=0, ReadOnly:=True, _
        AddToMru:=False, IgnoreReadOnlyRecommended:=True)

    On Error Resume Next
    Set sourceSheet = sourceBook.Worksheets("ネットワーク")
    On Error GoTo ErrorHandler

    If sourceSheet Is Nothing Then
        Err.Raise vbObjectError + 101, , "ネットワークシートが見つかりません。"
    End If

    Dim lastRow As Long, rowNumber As Long, blockEnd As Long
    Dim outputRow As Long, interfaceCount As Long
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "C").End(xlUp).Row
    outputRow = 2

    For rowNumber = 1 To lastRow
        If NormalizeText(ReadMerged(sourceSheet.Cells(rowNumber, "C"))) = "名前" _
           And NormalizeText(ReadMerged(sourceSheet.Cells(rowNumber, "K"))) <> "" Then

            blockEnd = FindBlockEnd(sourceSheet, rowNumber, lastRow)
            ExtractInterfaceBlock sourceSheet, rowNumber, blockEnd, outputSheet, outputRow
            interfaceCount = interfaceCount + 1
        End If
    Next rowNumber

    If interfaceCount = 0 Then
        Err.Raise vbObjectError + 102, , _
            "インターフェースブロックを検出できませんでした。"
    End If

    LogMessage "INFO", "詳細設計書抽出", _
        "Interfaces=" & interfaceCount & ", Parameters=" & (outputRow - 2)

CleanExit:
    On Error Resume Next
    If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    On Error GoTo 0
    Exit Sub

ErrorHandler:
    Dim errorNumber As Long, errorDescription As String
    errorNumber = Err.Number
    errorDescription = Err.Description
    ResumeAfterClose sourceBook
    Err.Raise errorNumber, , errorDescription
End Sub

Private Sub ResumeAfterClose(ByVal sourceBook As Workbook)
    On Error Resume Next
    If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Function FindBlockEnd(ByVal ws As Worksheet, ByVal startRow As Long, _
                              ByVal lastRow As Long) As Long
    Dim rowNumber As Long
    For rowNumber = startRow + 1 To lastRow
        If NormalizeText(ReadMerged(ws.Cells(rowNumber, "C"))) = "名前" _
           And NormalizeText(ReadMerged(ws.Cells(rowNumber, "K"))) <> "" Then
            FindBlockEnd = rowNumber - 1
            Exit Function
        End If
    Next rowNumber
    FindBlockEnd = lastRow
End Function

Private Sub ExtractInterfaceBlock(ByVal ws As Worksheet, ByVal startRow As Long, _
                                  ByVal endRow As Long, ByVal outputSheet As Worksheet, _
                                  ByRef outputRow As Long)
    Dim interfaceName As String, interfaceType As String
    Dim vdom As String, value As String, location As String

    interfaceName = NormalizeText(ReadMerged(ws.Cells(startRow, "K")))
    interfaceType = LCase$(GetLabelValue(ws, startRow, endRow, "タイプ", location))
    vdom = GetLabelValue(ws, startRow, endRow, "バーチャルドメイン", location)
    If vdom = "" Then vdom = "root"

    value = GetLabelValue(ws, startRow, endRow, "エイリアス", location)
    AddDesignRow outputSheet, outputRow, vdom, interfaceName, "alias", value, _
        "exact", "default-review", location, ""

    If UCase$(interfaceType) = "VLAN" Then
        AddDesignRow outputSheet, outputRow, vdom, interfaceName, "type", "vlan", _
            "exact", "error", FindLabelLocation(ws, startRow, endRow, "タイプ"), ""

        value = GetLabelValue(ws, startRow, endRow, "インターフェース", location)
        AddDesignRow outputSheet, outputRow, vdom, interfaceName, "interface", value, _
            "exact", "error", location, "VLAN親インターフェース"

        value = GetLabelValue(ws, startRow, endRow, "VLAN ID", location)
        AddDesignRow outputSheet, outputRow, vdom, interfaceName, "vlanid", value, _
            "exact", "error", location, ""
    End If

    AddDesignRow outputSheet, outputRow, vdom, interfaceName, "vdom", vdom, _
        "exact", "error", FindLabelLocation(ws, startRow, endRow, "バーチャルドメイン"), ""

    value = GetLabelValue(ws, startRow, endRow, "VRF ID", location)
    AddDesignRow outputSheet, outputRow, vdom, interfaceName, "vrf", value, _
        "exact", "default-review", location, ""

    value = LCase$(GetLabelValue(ws, startRow, endRow, "ロール", location))
    If value = "未定義" Then value = "undefined"
    AddDesignRow outputSheet, outputRow, vdom, interfaceName, "role", value, _
        "exact", "default-review", location, ""

    value = GetLabelValue(ws, startRow, endRow, "アドレッシングモード", location)
    If value = "マニュアル" Then value = "static"
    AddDesignRow outputSheet, outputRow, vdom, interfaceName, "mode", LCase$(value), _
        "exact", "default-review", location, ""

    value = GetFirstLabelValue(ws, startRow, endRow, "IP/ネットワーク", location)
    AddDesignRow outputSheet, outputRow, vdom, interfaceName, "ip", _
        NormalizeIpMask(value), "ip-mask", "default-review", location, ""

    Dim labelRow As Long
    labelRow = FindLabelRow(ws, startRow, endRow, "セカンダリIPアドレス")
    value = SelectedEnableDisable(ws, labelRow)
    AddDesignRow outputSheet, outputRow, vdom, interfaceName, "secondary-IP", value, _
        "exact", "default-review", CellLocation(ws, labelRow, 11), ""

    labelRow = FindLabelRow(ws, startRow, endRow, "IPv4")
    value = CollectAdminAccess(ws, labelRow, endRow)
    AddDesignRow outputSheet, outputRow, vdom, interfaceName, "allowaccess", value, _
        "list", "empty-is-match", CellLocation(ws, labelRow, 11), _
        "選択済みの管理アクセスのみ"

    labelRow = FindLabelRow(ws, startRow, endRow, "LLDP受信")
    value = MapLldpValue(SelectedLabel(ws, labelRow))
    AddDesignRow outputSheet, outputRow, vdom, interfaceName, "lldp-reception", value, _
        "exact", "default-review", CellLocation(ws, labelRow, 11), ""

    labelRow = FindLabelRow(ws, startRow, endRow, "LLDP送信")
    value = MapLldpValue(SelectedLabel(ws, labelRow))
    AddDesignRow outputSheet, outputRow, vdom, interfaceName, "lldp-transmission", value, _
        "exact", "default-review", CellLocation(ws, labelRow, 11), ""

    labelRow = FindLabelRow(ws, startRow, endRow, "ステータス")
    value = MapStatusValue(SelectedLabel(ws, labelRow))
    AddDesignRow outputSheet, outputRow, vdom, interfaceName, "status", value, _
        "exact", "default-review", CellLocation(ws, labelRow, 11), ""
End Sub

Private Sub AddDesignRow(ByVal ws As Worksheet, ByRef outputRow As Long, _
                         ByVal vdom As String, ByVal objectKey As String, _
                         ByVal parameter As String, ByVal value As String, _
                         ByVal compareMode As String, ByVal missingPolicy As String, _
                         ByVal sourceLocation As String, ByVal note As String)
    If value = "" And parameter = "alias" Then Exit Sub
    AddDesignParameter ws, outputRow, vdom, SECTION_INTERFACE, objectKey, _
                       parameter, value, compareMode, missingPolicy, _
                       sourceLocation, note
End Sub

Private Sub AddDesignParameter(ByVal ws As Worksheet, ByRef outputRow As Long, _
                               ByVal vdom As String, ByVal section As String, _
                               ByVal objectKey As String, ByVal parameter As String, _
                               ByVal value As String, ByVal compareMode As String, _
                               ByVal missingPolicy As String, _
                               ByVal sourceLocation As String, ByVal note As String)
    WriteText ws.Cells(outputRow, 1), "FG60"
    WriteText ws.Cells(outputRow, 2), vdom
    WriteText ws.Cells(outputRow, 3), section
    WriteText ws.Cells(outputRow, 4), objectKey
    WriteText ws.Cells(outputRow, 5), parameter
    WriteText ws.Cells(outputRow, 6), value
    WriteText ws.Cells(outputRow, 7), compareMode
    WriteText ws.Cells(outputRow, 8), missingPolicy
    WriteText ws.Cells(outputRow, 9), sourceLocation
    WriteText ws.Cells(outputRow, 10), note
    WriteText ws.Cells(outputRow, 11), BuildKey(vdom, section, objectKey, parameter)
    outputRow = outputRow + 1
End Sub

Private Function GetLabelValue(ByVal ws As Worksheet, ByVal startRow As Long, _
                               ByVal endRow As Long, ByVal label As String, _
                               ByRef location As String) As String
    Dim rowNumber As Long, columnNumber As Long
    For rowNumber = startRow To endRow
        For columnNumber = 3 To 5
            If NormalizeText(ReadMerged(ws.Cells(rowNumber, columnNumber))) = label Then
                location = CellLocation(ws, rowNumber, 11)
                GetLabelValue = NormalizeText(ReadMerged(ws.Cells(rowNumber, 11)))
                Exit Function
            End If
        Next columnNumber
    Next rowNumber
End Function

Private Function GetFirstLabelValue(ByVal ws As Worksheet, ByVal startRow As Long, _
                                    ByVal endRow As Long, ByVal label As String, _
                                    ByRef location As String) As String
    GetFirstLabelValue = GetLabelValue(ws, startRow, endRow, label, location)
End Function

Private Function FindLabelRow(ByVal ws As Worksheet, ByVal startRow As Long, _
                              ByVal endRow As Long, ByVal label As String) As Long
    Dim rowNumber As Long, columnNumber As Long
    For rowNumber = startRow To endRow
        For columnNumber = 3 To 5
            If NormalizeText(ReadMerged(ws.Cells(rowNumber, columnNumber))) = label Then
                FindLabelRow = rowNumber
                Exit Function
            End If
        Next columnNumber
    Next rowNumber
End Function

Private Function FindLabelLocation(ByVal ws As Worksheet, ByVal startRow As Long, _
                                   ByVal endRow As Long, ByVal label As String) As String
    Dim rowNumber As Long
    rowNumber = FindLabelRow(ws, startRow, endRow, label)
    If rowNumber > 0 Then FindLabelLocation = CellLocation(ws, rowNumber, 11)
End Function

Private Function ReadMerged(ByVal target As Range) As Variant
    If target.MergeCells Then
        ReadMerged = target.MergeArea.Cells(1, 1).Value2
    Else
        ReadMerged = target.Value2
    End If
End Function

Private Function SelectedEnableDisable(ByVal ws As Worksheet, ByVal rowNumber As Long) As String
    Dim selected As String
    selected = SelectedLabel(ws, rowNumber)
    Select Case selected
        Case "有効", "有効化済み": SelectedEnableDisable = "enable"
        Case "無効", "無効化済み": SelectedEnableDisable = "disable"
    End Select
End Function

Private Function SelectedLabel(ByVal ws As Worksheet, ByVal rowNumber As Long) As String
    If rowNumber = 0 Then Exit Function
    Dim columnNumber As Long
    For columnNumber = 11 To 29
        If NormalizeText(ReadMerged(ws.Cells(rowNumber, columnNumber))) = "■" Then
            SelectedLabel = NormalizeText(ReadMerged(ws.Cells(rowNumber, columnNumber + 1)))
            Exit Function
        End If
    Next columnNumber
End Function

Private Function CollectAdminAccess(ByVal ws As Worksheet, ByVal startRow As Long, _
                                    ByVal blockEnd As Long) As String
    If startRow = 0 Then Exit Function

    Dim values As Object
    Set values = CreateObject("Scripting.Dictionary")

    Dim rowNumber As Long, columnNumber As Long, label As String, mapped As String
    For rowNumber = startRow To WorksheetFunction.Min(startRow + 2, blockEnd)
        For columnNumber = 11 To 29
            If NormalizeText(ReadMerged(ws.Cells(rowNumber, columnNumber))) = "■" Then
                label = NormalizeText(ReadMerged(ws.Cells(rowNumber, columnNumber + 1)))
                mapped = MapAdminAccess(label)
                If mapped <> "" Then values(mapped) = True
            End If
        Next columnNumber
    Next rowNumber
    CollectAdminAccess = SortedDictionaryKeys(values)
End Function

Private Function MapAdminAccess(ByVal label As String) As String
    Select Case UCase$(label)
        Case "HTTPS": MapAdminAccess = "https"
        Case "PING": MapAdminAccess = "ping"
        Case "HTTP": MapAdminAccess = "http"
        Case "SSH": MapAdminAccess = "ssh"
        Case "SNMP": MapAdminAccess = "snmp"
        Case "TELNET": MapAdminAccess = "telnet"
        Case "FMGアクセス": MapAdminAccess = "fgfm"
        Case "RADIUSアカウンティング": MapAdminAccess = "radius-acct"
        Case "セキュリティファブリック接続": MapAdminAccess = "fabric"
        Case "スピードテスト": MapAdminAccess = "speed-test"
    End Select
End Function

Private Function MapLldpValue(ByVal value As String) As String
    Select Case value
        Case "VDOM設定を使用": MapLldpValue = "vdom"
        Case "有効": MapLldpValue = "enable"
        Case "無効": MapLldpValue = "disable"
    End Select
End Function

Private Function MapStatusValue(ByVal value As String) As String
    Select Case value
        Case "有効", "有効化済み": MapStatusValue = "up"
        Case "無効", "無効化済み": MapStatusValue = "down"
    End Select
End Function

Private Sub ParseFortiGateConfig(ByVal configPath As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONFIG)

    Dim fileNo As Integer, lineText As String, trimmed As String
    Dim lineNumber As Long, outputRow As Long
    Dim savedErrorNumber As Long, savedErrorDescription As String
    Dim sections As New Collection
    Dim currentObject As String, currentObjectSection As String
    Dim objectDepth As Long, editLine As Long
    Dim currentVdom As String, vdomDepth As Long
    Dim params As Object, locations As Object

    outputRow = 2
    fileNo = FreeFile
    On Error GoTo ErrorHandler
    Open configPath For Input As #fileNo

    Do Until EOF(fileNo)
        Line Input #fileNo, lineText
        lineNumber = lineNumber + 1
        trimmed = Trim$(lineText)

        If Left$(trimmed, 7) = "config " Then
            sections.Add LCase$(Trim$(Mid$(trimmed, 8)))

        ElseIf Left$(trimmed, 5) = "edit " Then
            If CurrentSection(sections) = "vdom" And currentObject = "" Then
                currentVdom = Unquote(Trim$(Mid$(trimmed, 6)))
                vdomDepth = sections.Count
            ElseIf CurrentSection(sections) = SECTION_INTERFACE _
                Or CurrentSection(sections) = SECTION_STATIC_ROUTE Then
                currentObject = Unquote(Trim$(Mid$(trimmed, 6)))
                currentObjectSection = CurrentSection(sections)
                objectDepth = sections.Count
                editLine = lineNumber
                Set params = CreateObject("Scripting.Dictionary")
                Set locations = CreateObject("Scripting.Dictionary")
            End If

        ElseIf Left$(trimmed, 4) = "set " Then
            If currentObject <> "" And sections.Count = objectDepth Then
                ApplySetLine trimmed, lineNumber, params, locations
            End If

        ElseIf Left$(trimmed, 6) = "unset " Then
            If currentObject <> "" And sections.Count = objectDepth Then
                Dim unsetParameter As String
                unsetParameter = CanonicalParameter(Trim$(Mid$(trimmed, 7)))
                params(unsetParameter) = ""
                locations(unsetParameter) = lineNumber
            End If

        ElseIf Left$(trimmed, 7) = "append " Then
            If currentObject <> "" And sections.Count = objectDepth Then
                ApplyListChange trimmed, lineNumber, params, locations, True
            End If

        ElseIf Left$(trimmed, 9) = "unselect " Then
            If currentObject <> "" And sections.Count = objectDepth Then
                ApplyListChange trimmed, lineNumber, params, locations, False
            End If

        ElseIf trimmed = "next" Then
            If currentObject <> "" And sections.Count = objectDepth Then
                FlushConfigObject ws, outputRow, currentVdom, currentObjectSection, _
                                  currentObject, editLine, params, locations
                currentObject = ""
                currentObjectSection = ""
                objectDepth = 0
                Set params = Nothing
                Set locations = Nothing
            ElseIf currentVdom <> "" And sections.Count = vdomDepth Then
                currentVdom = ""
                vdomDepth = 0
            End If

        ElseIf trimmed = "end" Then
            If sections.Count > 0 Then sections.Remove sections.Count
        End If
    Loop

    Close #fileNo
    LogMessage "INFO", "config解析", _
        "Lines=" & lineNumber & ", Parameters=" & (outputRow - 2)
    Exit Sub

ErrorHandler:
    savedErrorNumber = Err.Number
    savedErrorDescription = Err.Description
    On Error Resume Next
    If fileNo > 0 Then Close #fileNo
    On Error GoTo 0
    Err.Raise savedErrorNumber, , savedErrorDescription
End Sub

Private Sub ApplySetLine(ByVal lineText As String, ByVal lineNumber As Long, _
                         ByVal params As Object, ByVal locations As Object)
    Dim parameter As String, rawValue As String
    SplitCommandBody Mid$(lineText, 5), parameter, rawValue
    parameter = CanonicalParameter(parameter)
    params(parameter) = NormalizeConfigValue(parameter, rawValue)
    locations(parameter) = lineNumber
End Sub

Private Sub ApplyListChange(ByVal lineText As String, ByVal lineNumber As Long, _
                            ByVal params As Object, ByVal locations As Object, _
                            ByVal appendMode As Boolean)
    Dim commandLength As Long, parameter As String, rawValue As String
    If appendMode Then commandLength = 8 Else commandLength = 10
    SplitCommandBody Mid$(lineText, commandLength), parameter, rawValue
    parameter = CanonicalParameter(parameter)

    Dim values As Object, token As Variant
    Set values = CreateObject("Scripting.Dictionary")

    If params.Exists(parameter) Then
        For Each token In Split(CStr(params(parameter)), "|")
            If CStr(token) <> "" Then values(LCase$(CStr(token))) = True
        Next token
    End If

    Dim newTokens As Collection
    Set newTokens = TokenizeCli(rawValue)
    For Each token In newTokens
        If appendMode Then
            values(LCase$(CStr(token))) = True
        ElseIf values.Exists(LCase$(CStr(token))) Then
            values.Remove LCase$(CStr(token))
        End If
    Next token

    params(parameter) = SortedDictionaryKeys(values)
    locations(parameter) = lineNumber
End Sub

Private Sub SplitCommandBody(ByVal body As String, ByRef parameter As String, _
                             ByRef rawValue As String)
    body = Trim$(body)
    Dim position As Long
    position = InStr(body, " ")
    If position = 0 Then
        parameter = body
        rawValue = ""
    Else
        parameter = Left$(body, position - 1)
        rawValue = Trim$(Mid$(body, position + 1))
    End If
End Sub

Private Sub FlushConfigObject(ByVal ws As Worksheet, ByRef outputRow As Long, _
                              ByVal outerVdom As String, ByVal section As String, _
                              ByVal editName As String, ByVal editLine As Long, _
                              ByVal params As Object, ByVal locations As Object)
    Dim vdom As String, objectName As String
    vdom = outerVdom
    If vdom = "" Then vdom = "root"
    If section = SECTION_INTERFACE And params.Exists("vdom") Then _
        vdom = CStr(params("vdom"))

    objectName = editName
    If section = SECTION_STATIC_ROUTE Then
        If Not params.Exists("dst") Then params("dst") = "0.0.0.0 0.0.0.0"
        If Not params.Exists("gateway") Then params("gateway") = ""
        If Not params.Exists("device") Then params("device") = ""
        If Not params.Exists("distance") Then params("distance") = "10"
        If Not params.Exists("status") Then params("status") = "enable"
        objectName = BuildStaticRouteObjectKey(CStr(params("dst")), _
                     CStr(params("gateway")), CStr(params("device")))
    End If

    AddConfigParameter ws, outputRow, vdom, section, objectName, _
                       "_exists", "true", editLine

    Dim parameter As Variant
    For Each parameter In params.Keys
        If IsComparisonParameter(section, CStr(parameter)) Then
            Dim parameterLine As Long
            parameterLine = editLine
            If locations.Exists(CStr(parameter)) Then _
                parameterLine = CLng(locations(parameter))
            AddConfigParameter ws, outputRow, vdom, section, objectName, _
                CStr(parameter), CStr(params(parameter)), parameterLine
        End If
    Next parameter
End Sub

Private Sub AddConfigParameter(ByVal ws As Worksheet, ByRef outputRow As Long, _
                               ByVal vdom As String, ByVal section As String, _
                               ByVal objectKey As String, ByVal parameter As String, _
                               ByVal value As String, ByVal lineNumber As Long)
    WriteText ws.Cells(outputRow, 1), "FG60"
    WriteText ws.Cells(outputRow, 2), vdom
    WriteText ws.Cells(outputRow, 3), section
    WriteText ws.Cells(outputRow, 4), objectKey
    WriteText ws.Cells(outputRow, 5), parameter
    WriteText ws.Cells(outputRow, 6), value
    WriteText ws.Cells(outputRow, 7), "config line " & lineNumber
    WriteText ws.Cells(outputRow, 8), BuildKey(vdom, section, objectKey, parameter)
    outputRow = outputRow + 1
End Sub

Private Sub CompareNormalizedData()
    Dim designSheet As Worksheet, configSheet As Worksheet, resultSheet As Worksheet
    Set designSheet = ThisWorkbook.Worksheets(SHEET_DESIGN)
    Set configSheet = ThisWorkbook.Worksheets(SHEET_CONFIG)
    Set resultSheet = ThisWorkbook.Worksheets(SHEET_RESULT)

    Dim configValues As Object, configLocations As Object, objects As Object
    Set configValues = CreateObject("Scripting.Dictionary")
    Set configLocations = CreateObject("Scripting.Dictionary")
    Set objects = CreateObject("Scripting.Dictionary")

    Dim lastConfigRow As Long, rowNumber As Long, key As String, objectKey As String
    lastConfigRow = configSheet.Cells(configSheet.Rows.Count, 1).End(xlUp).Row

    For rowNumber = 2 To lastConfigRow
        key = CStr(configSheet.Cells(rowNumber, 8).Value2)
        If key <> "" Then
            configValues(key) = CStr(configSheet.Cells(rowNumber, 6).Value2)
            configLocations(key) = CStr(configSheet.Cells(rowNumber, 7).Value2)
        End If
        objectKey = BuildObjectKey( _
            CStr(configSheet.Cells(rowNumber, 2).Value2), _
            CStr(configSheet.Cells(rowNumber, 3).Value2), _
            CStr(configSheet.Cells(rowNumber, 4).Value2))
        If objectKey <> "||" Then objects(objectKey) = True
    Next rowNumber

    Dim lastDesignRow As Long, resultRow As Long
    Dim designValue As String, configValue As String
    Dim missingPolicy As String, result As String
    lastDesignRow = designSheet.Cells(designSheet.Rows.Count, 1).End(xlUp).Row
    resultRow = 2

    For rowNumber = 2 To lastDesignRow
        key = CStr(designSheet.Cells(rowNumber, 11).Value2)
        designValue = CStr(designSheet.Cells(rowNumber, 6).Value2)
        missingPolicy = CStr(designSheet.Cells(rowNumber, 8).Value2)
        objectKey = BuildObjectKey( _
            CStr(designSheet.Cells(rowNumber, 2).Value2), _
            CStr(designSheet.Cells(rowNumber, 3).Value2), _
            CStr(designSheet.Cells(rowNumber, 4).Value2))

        configValue = ""
        If Not objects.Exists(objectKey) Then
            result = "CONFIGオブジェクトなし"
        ElseIf Not configValues.Exists(key) Then
            Select Case missingPolicy
                Case "empty-is-match"
                    If designValue = "" Then result = "一致(未設定)" Else result = "CONFIG未記載"
                Case "default-review"
                    result = "要確認(省略可能)"
                Case Else
                    result = "CONFIG未記載"
            End Select
        Else
            configValue = CStr(configValues(key))
            If designValue = configValue Then
                result = "一致"
            Else
                result = "不一致"
            End If
        End If

        WriteText resultSheet.Cells(resultRow, 1), result
        WriteText resultSheet.Cells(resultRow, 2), CStr(designSheet.Cells(rowNumber, 2).Value2)
        WriteText resultSheet.Cells(resultRow, 3), CStr(designSheet.Cells(rowNumber, 3).Value2)
        WriteText resultSheet.Cells(resultRow, 4), CStr(designSheet.Cells(rowNumber, 4).Value2)
        WriteText resultSheet.Cells(resultRow, 5), CStr(designSheet.Cells(rowNumber, 5).Value2)
        WriteText resultSheet.Cells(resultRow, 6), designValue
        WriteText resultSheet.Cells(resultRow, 7), configValue
        WriteText resultSheet.Cells(resultRow, 8), CStr(designSheet.Cells(rowNumber, 9).Value2)
        If configLocations.Exists(key) Then _
            WriteText resultSheet.Cells(resultRow, 9), CStr(configLocations(key))
        WriteText resultSheet.Cells(resultRow, 10), CStr(designSheet.Cells(rowNumber, 7).Value2)
        WriteText resultSheet.Cells(resultRow, 11), missingPolicy
        WriteText resultSheet.Cells(resultRow, 12), key
        resultRow = resultRow + 1
    Next rowNumber

    LogMessage "INFO", "比較", "Results=" & (resultRow - 2)
End Sub

Private Function CurrentSection(ByVal sections As Collection) As String
    If sections.Count > 0 Then CurrentSection = CStr(sections(sections.Count))
End Function

Private Function CanonicalParameter(ByVal parameter As String) As String
    Select Case LCase$(Trim$(parameter))
        Case "secondary-ip": CanonicalParameter = "secondary-IP"
        Case Else: CanonicalParameter = LCase$(Trim$(parameter))
    End Select
End Function

Private Function IsComparisonParameter(ByVal section As String, _
                                       ByVal parameter As String) As Boolean
    If section = SECTION_INTERFACE Then
        Select Case parameter
            Case "alias", "type", "interface", "vlanid", "vdom", "vrf", _
                 "role", "mode", "ip", "secondary-IP", "allowaccess", _
                 "lldp-reception", "lldp-transmission", "status"
                IsComparisonParameter = True
        End Select
    ElseIf section = SECTION_STATIC_ROUTE Then
        Select Case parameter
            Case "dst", "gateway", "device", "distance", "comment", "status"
                IsComparisonParameter = True
        End Select
    End If
End Function

Private Function NormalizeConfigValue(ByVal parameter As String, _
                                      ByVal rawValue As String) As String
    Dim tokens As Collection
    Set tokens = TokenizeCli(rawValue)

    If parameter = "allowaccess" Then
        Dim values As Object, token As Variant
        Set values = CreateObject("Scripting.Dictionary")
        For Each token In tokens
            values(LCase$(CStr(token))) = True
        Next token
        NormalizeConfigValue = SortedDictionaryKeys(values)
    ElseIf parameter = "ip" Or parameter = "dst" Then
        NormalizeConfigValue = NormalizeIpMask(JoinTokens(tokens, " "))
    Else
        NormalizeConfigValue = JoinTokens(tokens, " ")
    End If
End Function

Private Function TokenizeCli(ByVal text As String) As Collection
    Dim result As New Collection
    Dim token As String, character As String
    Dim inQuote As Boolean, escaped As Boolean, index As Long

    For index = 1 To Len(text)
        character = Mid$(text, index, 1)
        If escaped Then
            token = token & character
            escaped = False
        ElseIf character = "\" Then
            escaped = True
        ElseIf character = Chr$(34) Then
            inQuote = Not inQuote
        ElseIf character = " " And Not inQuote Then
            If token <> "" Then result.Add token: token = ""
        Else
            token = token & character
        End If
    Next index
    If token <> "" Then result.Add token
    Set TokenizeCli = result
End Function

Private Function JoinTokens(ByVal tokens As Collection, ByVal delimiter As String) As String
    Dim result As String, index As Long
    For index = 1 To tokens.Count
        If result <> "" Then result = result & delimiter
        result = result & CStr(tokens(index))
    Next index
    JoinTokens = result
End Function

Private Function SortedDictionaryKeys(ByVal dictionary As Object) As String
    If dictionary.Count = 0 Then Exit Function

    Dim values() As String, keys As Variant
    Dim i As Long, j As Long, temp As String
    keys = dictionary.Keys
    ReDim values(LBound(keys) To UBound(keys))

    For i = LBound(keys) To UBound(keys)
        values(i) = CStr(keys(i))
    Next i
    For i = LBound(values) To UBound(values) - 1
        For j = i + 1 To UBound(values)
            If values(j) < values(i) Then
                temp = values(i): values(i) = values(j): values(j) = temp
            End If
        Next j
    Next i
    For i = LBound(values) To UBound(values)
        If SortedDictionaryKeys <> "" Then SortedDictionaryKeys = SortedDictionaryKeys & "|"
        SortedDictionaryKeys = SortedDictionaryKeys & values(i)
    Next i
End Function

Private Function NormalizeText(ByVal value As Variant) As String
    If IsError(value) Or IsEmpty(value) Then Exit Function
    Dim result As String
    result = CStr(value)
    result = Replace(result, ChrW(&H3000), " ")
    result = Replace(result, vbCrLf, " ")
    result = Replace(result, vbCr, " ")
    result = Replace(result, vbLf, " ")
    result = Replace(result, vbTab, " ")
    result = Trim$(result)
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    NormalizeText = result
End Function

Private Function NormalizeIpMask(ByVal value As String) As String
    Dim normalized As String, slashPosition As Long
    Dim prefixLength As Long, addressPart As String
    normalized = NormalizeText(value)
    slashPosition = InStr(normalized, "/")
    If slashPosition > 0 Then
        addressPart = Left$(normalized, slashPosition - 1)
        prefixLength = Val(Mid$(normalized, slashPosition + 1))
        If prefixLength >= 0 And prefixLength <= 32 Then
            NormalizeIpMask = addressPart & " " & PrefixToMask(prefixLength)
            Exit Function
        End If
    End If
    NormalizeIpMask = normalized
End Function

Private Function PrefixToMask(ByVal prefixLength As Long) As String
    Dim octet As Long, remaining As Long, value As Long, result As String
    remaining = prefixLength
    For octet = 1 To 4
        If remaining >= 8 Then
            value = 255
            remaining = remaining - 8
        ElseIf remaining > 0 Then
            value = 256 - (2 ^ (8 - remaining))
            remaining = 0
        Else
            value = 0
        End If
        If result <> "" Then result = result & "."
        result = result & CStr(value)
    Next octet
    PrefixToMask = result
End Function

Private Function Unquote(ByVal value As String) As String
    If Len(value) >= 2 Then
        If Left$(value, 1) = Chr$(34) And Right$(value, 1) = Chr$(34) Then
            Unquote = Mid$(value, 2, Len(value) - 2)
            Exit Function
        End If
    End If
    Unquote = value
End Function

Private Function BuildKey(ByVal vdom As String, ByVal section As String, _
                          ByVal objectKey As String, ByVal parameter As String) As String
    BuildKey = vdom & "|" & section & "|" & objectKey & "|" & parameter
End Function

Private Function BuildObjectKey(ByVal vdom As String, ByVal section As String, _
                                ByVal objectKey As String) As String
    BuildObjectKey = vdom & "|" & section & "|" & objectKey
End Function

Private Function CellLocation(ByVal ws As Worksheet, ByVal rowNumber As Long, _
                              ByVal columnNumber As Long) As String
    If rowNumber > 0 Then
        CellLocation = ws.Name & "!" & ws.Cells(rowNumber, columnNumber).Address(False, False)
    End If
End Function

Private Sub WriteText(ByVal target As Range, ByVal value As String)
    target.NumberFormat = "@"
    target.Value2 = value
End Sub

Private Sub LogMessage(ByVal level As String, ByVal operation As String, _
                       ByVal detail As String)
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(SHEET_LOG)

    If ws.Cells(1, 1).Value2 = "" Then
        WriteText ws.Cells(1, 1), "日時"
        WriteText ws.Cells(1, 2), "レベル"
        WriteText ws.Cells(1, 3), "処理"
        WriteText ws.Cells(1, 4), "内容"
    End If

    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 1).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    WriteText ws.Cells(nextRow, 2), level
    WriteText ws.Cells(nextRow, 3), operation
    WriteText ws.Cells(nextRow, 4), detail
End Sub

Private Sub FormatAllSheets()
    Dim sheetName As Variant, ws As Worksheet
    For Each sheetName In Array(SHEET_DESIGN, SHEET_CONFIG, SHEET_RESULT, SHEET_LOG)
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
        ws.Rows(1).Font.Bold = True
        ws.Rows(1).Interior.Color = RGB(23, 54, 93)
        ws.Rows(1).Font.Color = RGB(255, 255, 255)
        ws.Columns.AutoFit
        If ws.AutoFilterMode = False Then ws.Rows(1).AutoFilter
    Next sheetName

    With ThisWorkbook.Worksheets(SHEET_RESULT)
        .Columns("A").ColumnWidth = 24
        .Columns("F:G").ColumnWidth = 28
        .Columns("H:I").ColumnWidth = 22
        .Columns("L").ColumnWidth = 55
    End With
End Sub
