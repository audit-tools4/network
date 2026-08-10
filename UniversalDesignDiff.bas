Option Explicit

' ============================================================
' 汎用 詳細設計書／Config Diff
' - FortiGate、SW、APなど機器種別に依存しないExcelセル比較
' - 任意のテキストConfigの追加／削除比較
' - 外部通信、Python、Java、外部ライブラリ不要
' ============================================================

Private Const UD_SUMMARY As String = "UD_サマリ"
Private Const UD_EXCEL As String = "UD_Excel差分"
Private Const UD_CONFIG As String = "UD_Config差分"
Private Const UD_RULES As String = "UD_比較ルール"
Private Const UD_LOG As String = "UD_実行ログ"
Private Const UD_MAX_CELLS As Long = 300000

Public Sub RunUniversalDiff()
    Dim oldBookPath As Variant, newBookPath As Variant
    Dim oldConfigPath As Variant, newConfigPath As Variant

    oldBookPath = Application.GetOpenFilename( _
        "Excel,*.xlsx;*.xlsm;*.xls", , "過去の詳細設計書を選択")
    If VarType(oldBookPath) = vbBoolean Then Exit Sub

    newBookPath = Application.GetOpenFilename( _
        "Excel,*.xlsx;*.xlsm;*.xls", , "今回の詳細設計書を選択")
    If VarType(newBookPath) = vbBoolean Then Exit Sub

    oldConfigPath = ""
    newConfigPath = ""
    If MsgBox("Configも比較しますか？", vbYesNo + vbQuestion) = vbYes Then
        oldConfigPath = Application.GetOpenFilename( _
            "Config,*.conf;*.cfg;*.txt;*.*", , "過去Configを選択")
        If VarType(oldConfigPath) = vbBoolean Then Exit Sub
        newConfigPath = Application.GetOpenFilename( _
            "Config,*.conf;*.cfg;*.txt;*.*", , "今回Configを選択")
        If VarType(newConfigPath) = vbBoolean Then Exit Sub
    End If

    RunUniversalDiffFromPaths CStr(oldBookPath), CStr(newBookPath), _
        CStr(oldConfigPath), CStr(newConfigPath), True
End Sub

Public Sub RunUniversalDiffFromPaths(ByVal oldBookPath As String, _
                                     ByVal newBookPath As String, _
                                     Optional ByVal oldConfigPath As String = "", _
                                     Optional ByVal newConfigPath As String = "", _
                                     Optional ByVal showMessage As Boolean = True)
    Dim oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean
    Dim oldCalculation As XlCalculation, oldSecurity As Long
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
    oldSecurity = Application.AutomationSecurity

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.AutomationSecurity = 3
    On Error GoTo ErrorHandler

    PrepareUniversalSheets
    UDLog "INFO", "処理開始", "過去=" & FileNameOnly(oldBookPath) & _
          ", 今回=" & FileNameOnly(newBookPath)
    CompareExcelBooks oldBookPath, newBookPath
    If oldConfigPath <> "" And newConfigPath <> "" Then
        CompareConfigFiles oldConfigPath, newConfigPath
    End If
    BuildSummary oldBookPath, newBookPath, oldConfigPath, newConfigPath
    FormatUniversalSheets
    ThisWorkbook.Worksheets(UD_SUMMARY).Activate
    If showMessage Then MsgBox "比較が完了しました。UD_サマリを確認してください。", vbInformation

CleanExit:
    Application.AutomationSecurity = oldSecurity
    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldAlerts
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Exit Sub

ErrorHandler:
    UDLog "ERROR", "処理中断", Err.Number & ": " & Err.Description
    If showMessage Then MsgBox "比較中にエラーが発生しました。" & vbCrLf & _
                               Err.Description, vbExclamation
    Resume CleanExit
End Sub

Public Sub ResetUniversalRules()
    InitializeUniversalRules True
End Sub

Private Sub InitializeUniversalRules(ByVal showMessage As Boolean)
    Dim ws As Worksheet
    Set ws = UDGetOrCreateSheet(UD_RULES)
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Cells.Clear
    UDWriteHeaders ws, Array("有効", "対象", "シート/ファイル", "項目/文言", _
                             "処理", "説明")
    Dim defaults As Variant, r As Long, c As Long
    defaults = Array( _
        Array("○", "Excel", "表紙", "*", "比較除外", "表紙全体を除外"), _
        Array("○", "Excel", "改訂履歴", "*", "比較除外", "改訂履歴を除外"), _
        Array("○", "Excel", "*", "作成日", "拠点差分", "日付変更を重点対象外にする"), _
        Array("○", "Excel", "*", "ホスト名", "拠点差分", "拠点ごとに異なる値"), _
        Array("○", "Excel", "*", "管理IP", "拠点差分", "拠点ごとに異なる値"), _
        Array("○", "Config", "*", "password", "秘密情報", "値をマスク"), _
        Array("○", "Config", "*", "secret", "秘密情報", "値をマスク"), _
        Array("○", "Config", "*", "psk", "秘密情報", "値をマスク"), _
        Array("○", "Config", "*", "community", "秘密情報", "値をマスク"), _
        Array("○", "Config", "*", "private-key", "秘密情報", "値をマスク"), _
        Array("○", "Config", "*", "token", "秘密情報", "値をマスク"), _
        Array("○", "Config", "*", "certificate", "秘密情報", "値をマスク"))
    For r = LBound(defaults) To UBound(defaults)
        For c = LBound(defaults(r)) To UBound(defaults(r))
            ws.Cells(r + 2, c + 1).Value2 = defaults(r)(c)
        Next c
    Next r
    UDFormatTable ws, 6
    If showMessage Then _
        MsgBox "UD_比較ルールを初期化しました。必要に応じて編集してください。", vbInformation
End Sub

Private Sub PrepareUniversalSheets()
    UDPrepareSheet UD_SUMMARY, Array("項目", "値")
    UDPrepareSheet UD_EXCEL, Array( _
        "判定", "扱い", "シート", "項目候補", "過去場所", "今回場所", _
        "過去値", "今回値", "過去数式", "今回数式", "確度", "キー")
    UDPrepareSheet UD_CONFIG, Array( _
        "判定", "扱い", "正規化行", "過去行", "今回行", _
        "過去内容", "今回内容", "秘密情報", "キー")
    UDPrepareSheet UD_LOG, Array("日時", "レベル", "処理", "内容")
    If Not UDSheetExists(UD_RULES) Then InitializeUniversalRules False
End Sub

Private Sub CompareExcelBooks(ByVal oldPath As String, ByVal newPath As String)
    Dim oldIndex As Object, newIndex As Object
    Set oldIndex = LoadWorkbookIndex(oldPath)
    Set newIndex = LoadWorkbookIndex(newPath)

    Dim matchedOld As Object, matchedNew As Object, oldByValue As Object
    Set matchedOld = CreateObject("Scripting.Dictionary")
    Set matchedNew = CreateObject("Scripting.Dictionary")
    Set oldByValue = CreateObject("Scripting.Dictionary")

    Dim ws As Worksheet, outputRow As Long, key As Variant
    Dim oldItem As Variant, newItem As Variant, result As String, handling As String
    Set ws = ThisWorkbook.Worksheets(UD_EXCEL)
    outputRow = 2

    ' 同じシート・セルを先に比較する。
    For Each key In newIndex.Keys
        If oldIndex.Exists(CStr(key)) Then
            oldItem = oldIndex(CStr(key))
            newItem = newIndex(CStr(key))
            matchedOld(CStr(key)) = True
            matchedNew(CStr(key)) = True
            handling = ResolveExcelHandling(CStr(newItem(0)), CStr(newItem(5)))
            If handling <> "比較除外" Then
                If CStr(oldItem(2)) = CStr(newItem(2)) And _
                   CStr(oldItem(3)) = CStr(newItem(3)) Then
                    result = "同一"
                ElseIf CStr(oldItem(2)) = CStr(newItem(2)) Then
                    result = "数式変更"
                Else
                    result = "変更"
                End If
                If result <> "同一" Then
                    WriteExcelDiff ws, outputRow, result, handling, oldItem, newItem, _
                                   "高", CStr(key)
                End If
            End If
        End If
    Next key

    ' 未対応の旧セルを正規化値ごとに索引化し、移動候補を探す。
    For Each key In oldIndex.Keys
        If Not matchedOld.Exists(CStr(key)) Then
            oldItem = oldIndex(CStr(key))
            If CStr(oldItem(2)) <> "" Then AddCollectionItem oldByValue, _
                CStr(oldItem(0)) & "|" & CStr(oldItem(2)), CStr(key)
        End If
    Next key

    Dim candidateKey As String, oldLocationKey As String
    For Each key In newIndex.Keys
        If Not matchedNew.Exists(CStr(key)) Then
            newItem = newIndex(CStr(key))
            candidateKey = CStr(newItem(0)) & "|" & CStr(newItem(2))
            handling = ResolveExcelHandling(CStr(newItem(0)), CStr(newItem(5)))
            If handling = "比較除外" Then
                matchedNew(CStr(key)) = True
            ElseIf oldByValue.Exists(candidateKey) Then
                If oldByValue(candidateKey).Count > 0 Then
                    oldLocationKey = CStr(oldByValue(candidateKey)(1))
                    oldByValue(candidateKey).Remove 1
                    oldItem = oldIndex(oldLocationKey)
                    matchedOld(oldLocationKey) = True
                    matchedNew(CStr(key)) = True
                    WriteExcelDiff ws, outputRow, "移動候補", handling, oldItem, _
                                   newItem, "中", oldLocationKey & "->" & CStr(key)
                End If
            End If
        End If
    Next key

    For Each key In newIndex.Keys
        If Not matchedNew.Exists(CStr(key)) Then
            newItem = newIndex(CStr(key))
            handling = ResolveExcelHandling(CStr(newItem(0)), CStr(newItem(5)))
            If handling <> "比較除外" Then
                WriteExcelDiff ws, outputRow, "追加", handling, EmptyExcelItem(), _
                               newItem, "高", CStr(key)
            End If
        End If
    Next key

    For Each key In oldIndex.Keys
        If Not matchedOld.Exists(CStr(key)) Then
            oldItem = oldIndex(CStr(key))
            handling = ResolveExcelHandling(CStr(oldItem(0)), CStr(oldItem(5)))
            If handling <> "比較除外" Then
                WriteExcelDiff ws, outputRow, "削除", handling, oldItem, _
                               EmptyExcelItem(), "高", CStr(key)
            End If
        End If
    Next key
    UDLog "INFO", "Excel比較", "差分=" & (outputRow - 2)
End Sub

Private Function LoadWorkbookIndex(ByVal workbookPath As String) As Object
    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    Dim sourceBook As Workbook, ws As Worksheet, used As Range, cell As Range
    Dim totalCells As Double, key As String, value As String, formula As String
    Dim address As String, label As String, item As Variant
    On Error GoTo ErrorHandler
    Set sourceBook = Workbooks.Open(workbookPath, UpdateLinks:=0, ReadOnly:=True, _
                                   AddToMru:=False, IgnoreReadOnlyRecommended:=True)
    For Each ws In sourceBook.Worksheets
        Set used = ws.UsedRange
        totalCells = totalCells + CDbl(used.Cells.CountLarge)
        If totalCells > UD_MAX_CELLS Then
            Err.Raise vbObjectError + 701, , _
                "比較セル数が上限を超えました。不要シートを比較ルールで除外するか、UsedRangeを整理してください。"
        End If
        If ResolveExcelHandling(ws.Name, "") <> "比較除外" Then
            For Each cell In used.Cells
                If IsMergeAnchor(cell) Then
                    value = NormalizeCellValue(cell.Value2)
                    formula = NormalizeCellValue(cell.Formula)
                    If value <> "" Or formula <> "" Then
                        address = cell.Address(False, False)
                        label = FindRowLabel(ws, cell.Row, cell.Column)
                        key = ws.Name & "|" & address
                        item = Array(ws.Name, address, value, formula, _
                                     CellDisplayValue(cell), label)
                        result(key) = item
                    End If
                End If
            Next cell
        End If
    Next ws
    sourceBook.Close SaveChanges:=False
    Set LoadWorkbookIndex = result
    Exit Function

ErrorHandler:
    Dim number As Long, description As String
    number = Err.Number: description = Err.Description
    On Error Resume Next
    If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise number, , description
End Function

Private Sub WriteExcelDiff(ByVal ws As Worksheet, ByRef rowNumber As Long, _
                           ByVal result As String, ByVal handling As String, _
                           ByVal oldItem As Variant, ByVal newItem As Variant, _
                           ByVal confidence As String, ByVal key As String)
    Dim sheetName As String, label As String
    sheetName = CStr(newItem(0)): If sheetName = "" Then sheetName = CStr(oldItem(0))
    label = CStr(newItem(5)): If label = "" Then label = CStr(oldItem(5))
    ws.Cells(rowNumber, 1).Value2 = result
    ws.Cells(rowNumber, 2).Value2 = handling
    ws.Cells(rowNumber, 3).Value2 = sheetName
    ws.Cells(rowNumber, 4).Value2 = label
    ws.Cells(rowNumber, 5).Value2 = CStr(oldItem(1))
    ws.Cells(rowNumber, 6).Value2 = CStr(newItem(1))
    ws.Cells(rowNumber, 7).Value2 = CStr(oldItem(4))
    ws.Cells(rowNumber, 8).Value2 = CStr(newItem(4))
    ws.Cells(rowNumber, 9).Value2 = CStr(oldItem(3))
    ws.Cells(rowNumber, 10).Value2 = CStr(newItem(3))
    ws.Cells(rowNumber, 11).Value2 = confidence
    ws.Cells(rowNumber, 12).Value2 = key
    rowNumber = rowNumber + 1
End Sub

Private Function EmptyExcelItem() As Variant
    EmptyExcelItem = Array("", "", "", "", "", "")
End Function

Private Sub CompareConfigFiles(ByVal oldPath As String, ByVal newPath As String)
    Dim oldLines As Object, newLines As Object
    Set oldLines = LoadConfigLines(oldPath)
    Set newLines = LoadConfigLines(newPath)

    Dim ws As Worksheet, outputRow As Long, key As Variant
    Dim oldCount As Long, newCount As Long, index As Long
    Set ws = ThisWorkbook.Worksheets(UD_CONFIG)
    outputRow = 2

    For Each key In oldLines.Keys
        oldCount = oldLines(key).Count
        If newLines.Exists(CStr(key)) Then newCount = newLines(key).Count Else newCount = 0
        If oldCount > newCount Then
            For index = newCount + 1 To oldCount
                WriteConfigDiff ws, outputRow, "削除", CStr(key), _
                    CStr(oldLines(key)(index)), "", CStr(key)
            Next index
        End If
    Next key

    For Each key In newLines.Keys
        newCount = newLines(key).Count
        If oldLines.Exists(CStr(key)) Then oldCount = oldLines(key).Count Else oldCount = 0
        If newCount > oldCount Then
            For index = oldCount + 1 To newCount
                WriteConfigDiff ws, outputRow, "追加", CStr(key), "", _
                    CStr(newLines(key)(index)), CStr(key)
            Next index
        End If
    Next key
    UDLog "INFO", "Config比較", "差分=" & (outputRow - 2)
End Sub

Private Function LoadConfigLines(ByVal filePath As String) As Object
    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    Dim fileNo As Integer, lineText As String, lineNumber As Long
    Dim normalized As String, handling As String, stored As String
    On Error GoTo ErrorHandler
    fileNo = FreeFile
    Open filePath For Input As #fileNo
    Do Until EOF(fileNo)
        Line Input #fileNo, lineText
        lineNumber = lineNumber + 1
        normalized = NormalizeConfigLine(lineText)
        If normalized <> "" Then
            handling = ResolveConfigHandling(FileNameOnly(filePath), normalized)
            If handling <> "比較除外" Then
                If handling = "秘密情報" Then normalized = MaskSensitiveLine(normalized)
                stored = CStr(lineNumber) & "|" & MaskSensitiveLine(Trim$(lineText))
                AddCollectionItem result, normalized, stored
            End If
        End If
    Loop
    Close #fileNo
    Set LoadConfigLines = result
    Exit Function
ErrorHandler:
    Dim number As Long, description As String
    number = Err.Number: description = Err.Description
    On Error Resume Next
    If fileNo > 0 Then Close #fileNo
    On Error GoTo 0
    Err.Raise number, , description
End Function

Private Sub WriteConfigDiff(ByVal ws As Worksheet, ByRef rowNumber As Long, _
                            ByVal result As String, ByVal normalized As String, _
                            ByVal oldStored As String, ByVal newStored As String, _
                            ByVal key As String)
    Dim oldLine As String, oldText As String, newLine As String, newText As String
    SplitStoredLine oldStored, oldLine, oldText
    SplitStoredLine newStored, newLine, newText
    Dim handling As String
    handling = ResolveConfigHandling("*", normalized)
    ws.Cells(rowNumber, 1).Value2 = result
    ws.Cells(rowNumber, 2).Value2 = handling
    ws.Cells(rowNumber, 3).Value2 = normalized
    ws.Cells(rowNumber, 4).Value2 = oldLine
    ws.Cells(rowNumber, 5).Value2 = newLine
    ws.Cells(rowNumber, 6).Value2 = oldText
    ws.Cells(rowNumber, 7).Value2 = newText
    ws.Cells(rowNumber, 8).Value2 = IIf(handling = "秘密情報", "マスク済み", "")
    ws.Cells(rowNumber, 9).Value2 = key
    rowNumber = rowNumber + 1
End Sub

Private Sub SplitStoredLine(ByVal stored As String, ByRef lineNumber As String, _
                            ByRef lineText As String)
    Dim position As Long: position = InStr(stored, "|")
    If position > 0 Then
        lineNumber = Left$(stored, position - 1)
        lineText = Mid$(stored, position + 1)
    End If
End Sub

Private Function NormalizeConfigLine(ByVal value As String) As String
    Dim result As String
    result = NormalizeWhitespace(value)
    If result = "" Then Exit Function
    If Left$(result, 1) = "#" Or Left$(result, 1) = "!" Then Exit Function
    NormalizeConfigLine = LCase$(result)
End Function

Private Function MaskSensitiveLine(ByVal value As String) As String
    If ResolveConfigHandling("*", value) <> "秘密情報" Then
        MaskSensitiveLine = value
        Exit Function
    End If
    Dim firstSpace As Long, secondSpace As Long
    firstSpace = InStr(value, " ")
    If firstSpace > 0 Then secondSpace = InStr(firstSpace + 1, value, " ")
    If secondSpace > 0 Then
        MaskSensitiveLine = Left$(value, secondSpace - 1) & " ********"
    ElseIf firstSpace > 0 Then
        MaskSensitiveLine = Left$(value, firstSpace - 1) & " ********"
    Else
        MaskSensitiveLine = "********"
    End If
End Function

Private Function ResolveExcelHandling(ByVal sheetName As String, _
                                      ByVal context As String) As String
    ResolveExcelHandling = ResolveRule("Excel", sheetName, context, "一致必須")
End Function

Private Function ResolveConfigHandling(ByVal fileName As String, _
                                       ByVal lineText As String) As String
    ResolveConfigHandling = ResolveRule("Config", fileName, lineText, "通常比較")
End Function

Private Function ResolveRule(ByVal target As String, ByVal container As String, _
                             ByVal text As String, ByVal defaultAction As String) As String
    ResolveRule = defaultAction
    If Not UDSheetExists(UD_RULES) Then Exit Function
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(UD_RULES)
    Dim lastRow As Long, rowNumber As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For rowNumber = 2 To lastRow
        If IsRuleEnabled(CStr(ws.Cells(rowNumber, 1).Value2)) _
           And LCase$(NormalizeWhitespace(CStr(ws.Cells(rowNumber, 2).Value2))) = LCase$(target) Then
            If WildcardMatch(container, CStr(ws.Cells(rowNumber, 3).Value2)) And _
               ContainsPattern(text, CStr(ws.Cells(rowNumber, 4).Value2)) Then
                ResolveRule = NormalizeWhitespace(CStr(ws.Cells(rowNumber, 5).Value2))
                Exit Function
            End If
        End If
    Next rowNumber
End Function

Private Function IsRuleEnabled(ByVal value As String) As Boolean
    Select Case LCase$(NormalizeWhitespace(value))
        Case "○", "on", "yes", "true", "1", "有効"
            IsRuleEnabled = True
    End Select
End Function

Private Function WildcardMatch(ByVal value As String, ByVal pattern As String) As Boolean
    pattern = NormalizeWhitespace(pattern)
    If pattern = "" Or pattern = "*" Then WildcardMatch = True: Exit Function
    WildcardMatch = (LCase$(value) Like LCase$(pattern))
End Function

Private Function ContainsPattern(ByVal value As String, ByVal pattern As String) As Boolean
    pattern = NormalizeWhitespace(pattern)
    If pattern = "" Or pattern = "*" Then ContainsPattern = True: Exit Function
    ContainsPattern = (InStr(1, value, pattern, vbTextCompare) > 0)
End Function

Private Function FindRowLabel(ByVal ws As Worksheet, ByVal rowNumber As Long, _
                              ByVal valueColumn As Long) As String
    Dim columnNumber As Long, candidate As String
    For columnNumber = valueColumn - 1 To 1 Step -1
        candidate = NormalizeCellValue(ReadMergeAnchor(ws.Cells(rowNumber, columnNumber)).Value2)
        If candidate <> "" Then FindRowLabel = candidate: Exit Function
    Next columnNumber
End Function

Private Function IsMergeAnchor(ByVal cell As Range) As Boolean
    If Not cell.MergeCells Then
        IsMergeAnchor = True
    Else
        IsMergeAnchor = (cell.Address = cell.MergeArea.Cells(1, 1).Address)
    End If
End Function

Private Function ReadMergeAnchor(ByVal cell As Range) As Range
    If cell.MergeCells Then
        Set ReadMergeAnchor = cell.MergeArea.Cells(1, 1)
    Else
        Set ReadMergeAnchor = cell
    End If
End Function

Private Function CellDisplayValue(ByVal cell As Range) As String
    On Error Resume Next
    CellDisplayValue = NormalizeWhitespace(CStr(cell.Text))
    If Err.Number <> 0 Then Err.Clear: CellDisplayValue = NormalizeCellValue(cell.Value2)
    On Error GoTo 0
End Function

Private Function NormalizeCellValue(ByVal value As Variant) As String
    If IsError(value) Or IsEmpty(value) Then Exit Function
    NormalizeCellValue = NormalizeWhitespace(CStr(value))
End Function

Private Function NormalizeWhitespace(ByVal value As String) As String
    Dim result As String
    result = Replace(value, ChrW(&H3000), " ")
    result = Replace(result, vbCrLf, " ")
    result = Replace(result, vbCr, " ")
    result = Replace(result, vbLf, " ")
    result = Replace(result, vbTab, " ")
    result = Trim$(result)
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    NormalizeWhitespace = result
End Function

Private Sub AddCollectionItem(ByVal dictionary As Object, ByVal key As String, _
                              ByVal value As String)
    Dim items As Collection
    If dictionary.Exists(key) Then
        Set items = dictionary(key)
    Else
        Set items = New Collection
        dictionary.Add key, items
    End If
    items.Add value
End Sub

Private Sub BuildSummary(ByVal oldBookPath As String, ByVal newBookPath As String, _
                         ByVal oldConfigPath As String, ByVal newConfigPath As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(UD_SUMMARY)
    Dim rowNumber As Long: rowNumber = 2
    UDSummaryRow ws, rowNumber, "実行日時", Format$(Now, "yyyy/mm/dd hh:nn:ss")
    UDSummaryRow ws, rowNumber, "過去設計書", FileNameOnly(oldBookPath)
    UDSummaryRow ws, rowNumber, "今回設計書", FileNameOnly(newBookPath)
    UDSummaryRow ws, rowNumber, "Excel差分件数", CStr(UDDataRowCount(UD_EXCEL))
    If oldConfigPath <> "" Then
        UDSummaryRow ws, rowNumber, "過去Config", FileNameOnly(oldConfigPath)
        UDSummaryRow ws, rowNumber, "今回Config", FileNameOnly(newConfigPath)
        UDSummaryRow ws, rowNumber, "Config差分件数", CStr(UDDataRowCount(UD_CONFIG))
    Else
        UDSummaryRow ws, rowNumber, "Config比較", "未実施"
    End If
    UDSummaryRow ws, rowNumber, "重要", _
        "差分は候補です。拠点固有値・設定順序・メーカー仕様を人間が最終確認してください。"
End Sub

Private Sub UDSummaryRow(ByVal ws As Worksheet, ByRef rowNumber As Long, _
                         ByVal label As String, ByVal value As String)
    ws.Cells(rowNumber, 1).Value2 = label
    ws.Cells(rowNumber, 2).Value2 = value
    rowNumber = rowNumber + 1
End Sub

Private Function UDDataRowCount(ByVal sheetName As String) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(sheetName)
    UDDataRowCount = Application.Max(0, ws.Cells(ws.Rows.Count, 1).End(xlUp).Row - 1)
End Function

Private Sub UDPrepareSheet(ByVal sheetName As String, ByVal headers As Variant)
    Dim ws As Worksheet: Set ws = UDGetOrCreateSheet(sheetName)
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Cells.Clear
    UDWriteHeaders ws, headers
End Sub

Private Sub UDWriteHeaders(ByVal ws As Worksheet, ByVal headers As Variant)
    Dim index As Long
    For index = LBound(headers) To UBound(headers)
        ws.Cells(1, index + 1).Value2 = CStr(headers(index))
    Next index
End Sub

Private Function UDGetOrCreateSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set UDGetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If UDGetOrCreateSheet Is Nothing Then
        Set UDGetOrCreateSheet = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        UDGetOrCreateSheet.Name = sheetName
    End If
End Function

Private Function UDSheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    UDSheetExists = Not ws Is Nothing
End Function

Private Sub UDFormatTable(ByVal ws As Worksheet, ByVal columnCount As Long)
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, columnCount))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .AutoFilter
    End With
    ws.Rows(1).RowHeight = 24
    ws.Rows(1).VerticalAlignment = xlCenter
    ws.Columns.AutoFit
    ws.Activate
    ActiveWindow.SplitColumn = 0
    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True
End Sub

Private Sub FormatUniversalSheets()
    Dim sheetName As Variant, ws As Worksheet
    For Each sheetName In Array(UD_SUMMARY, UD_EXCEL, UD_CONFIG, UD_RULES, UD_LOG)
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
        UDFormatTable ws, ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        ws.Cells.VerticalAlignment = xlTop
        ws.Cells.WrapText = False
    Next sheetName
    With ThisWorkbook.Worksheets(UD_SUMMARY)
        .Columns("A").ColumnWidth = 22
        .Columns("B").ColumnWidth = 80
        .Columns("B").WrapText = True
    End With
    With ThisWorkbook.Worksheets(UD_EXCEL)
        .Columns("A:D").ColumnWidth = 16
        .Columns("E:F").ColumnWidth = 18
        .Columns("G:J").ColumnWidth = 32
        .Columns("L").ColumnWidth = 45
        .Columns("G:J").WrapText = True
    End With
    With ThisWorkbook.Worksheets(UD_CONFIG)
        .Columns("A:B").ColumnWidth = 15
        .Columns("C").ColumnWidth = 55
        .Columns("F:G").ColumnWidth = 55
        .Columns("C:G").WrapText = True
    End With
    ApplyResultColors ThisWorkbook.Worksheets(UD_EXCEL)
    ApplyResultColors ThisWorkbook.Worksheets(UD_CONFIG)
End Sub

Private Sub ApplyResultColors(ByVal ws As Worksheet)
    Dim lastRow As Long, rowNumber As Long, result As String
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For rowNumber = 2 To lastRow
        result = CStr(ws.Cells(rowNumber, 1).Value2)
        Select Case result
            Case "追加": ws.Cells(rowNumber, 1).Interior.Color = RGB(221, 235, 247)
            Case "削除": ws.Cells(rowNumber, 1).Interior.Color = RGB(244, 204, 204)
            Case "変更", "数式変更": _
                ws.Cells(rowNumber, 1).Interior.Color = RGB(255, 242, 204)
            Case "移動候補": ws.Cells(rowNumber, 1).Interior.Color = RGB(226, 239, 218)
        End Select
    Next rowNumber
End Sub

Private Sub UDLog(ByVal level As String, ByVal operation As String, _
                  ByVal detail As String)
    Dim ws As Worksheet: Set ws = UDGetOrCreateSheet(UD_LOG)
    If ws.Cells(1, 1).Value2 = "" Then _
        UDWriteHeaders ws, Array("日時", "レベル", "処理", "内容")
    Dim nextRow As Long: nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 1).NumberFormat = "yyyy/mm/dd hh:mm:ss"
    ws.Cells(nextRow, 2).Value2 = level
    ws.Cells(nextRow, 3).Value2 = operation
    ws.Cells(nextRow, 4).Value2 = detail
End Sub

Private Function FileNameOnly(ByVal filePath As String) As String
    Dim position As Long
    position = InStrRev(filePath, Application.PathSeparator)
    If position > 0 Then FileNameOnly = Mid$(filePath, position + 1) Else FileNameOnly = filePath
End Function
