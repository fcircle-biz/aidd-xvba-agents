Attribute VB_Name = "modDashboard"
'=============================================================================
' modDashboard.bas - ダッシュボードUI・KPI管理モジュール
'=============================================================================
' 概要:
'   Dashboardシートの UI操作、KPI表示、ボタンイベント処理
'   統計情報の収集・表示、レポート生成機能を提供
'=============================================================================
Option Explicit

'=============================================================================
' ダッシュボード初期化
'=============================================================================

' ダッシュボードシート初期化
Public Sub InitializeDashboard()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_DASHBOARD)
    If ws Is Nothing Then
        Call modCmn.LogError("InitializeDashboard", "Dashboardシートが見つかりません")
        Exit Sub
    End If
    
    ' シート名設定
    ws.Name = SHEET_DASHBOARD
    
    ' フォント適用
    Call modCmn.ApplySheetFont(ws)
    
    ' UI要素作成
    Call CreateDashboardLayout(ws)
    Call CreateDashboardButtons(ws)
    
    ' 初期KPI表示
    Call RefreshKPI()
    
    Call modCmn.LogInfo("InitializeDashboard", "ダッシュボード初期化完了")
    Exit Sub
    
ErrHandler:
    Call modCmn.LogError("InitializeDashboard", "ダッシュボード初期化エラー: " & Err.Description)
End Sub

' ダッシュボードレイアウト作成
Private Sub CreateDashboardLayout(ByVal ws As Worksheet)
    On Error Resume Next
    
    ' ヘッダー
    With ws.Range("A1")
        .Value = SYSTEM_NAME
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(0, 100, 200)
    End With
    
    With ws.Range("A2")
        .Value = "バージョン: " & SYSTEM_VERSION
        .Font.Size = 10
        .Font.Color = RGB(100, 100, 100)
    End With
    
    ' KPIセクションヘッダー
    With ws.Range("A4")
        .Value = "=== システム状況 ==="
        .Font.Size = 12
        .Font.Bold = True
    End With
    
    ' KPI項目
    ws.Range("B5").Value = "総顧客数:"
    ws.Range("B6").Value = "追加件数:"
    ws.Range("B7").Value = "更新件数:"
    ws.Range("B8").Value = "重複検出:"
    ws.Range("B9").Value = "エラー件数:"
    ws.Range("B10").Value = "無効化件数:"
    ws.Range("B11").Value = "最終取込日時:"
    ws.Range("B12").Value = "処理時間:"
    
    ' KPI値セル（右揃え）
    ws.Range("D5:D12").HorizontalAlignment = xlRight
    
    ' ボタンセクションヘッダー
    With ws.Range("A14")
        .Value = "=== 操作メニュー ==="
        .Font.Size = 12
        .Font.Bold = True
    End With
    
    ' 列幅調整
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 15
    ws.Columns("C").ColumnWidth = 2
    ws.Columns("D").ColumnWidth = 15
    ws.Columns("E").ColumnWidth = 20
End Sub

' ダッシュボードボタン作成
Private Sub CreateDashboardButtons(ByVal ws As Worksheet)
    On Error GoTo ErrHandler
    
    Dim btn As Button
    
    ' メインボタン: CSV取り込み→整形→検証→反映
    Set btn = ws.Buttons.Add(30, 240, 200, 30) ' B15セル位置
    With btn
        .Caption = "CSV一括取込・更新実行"
        .OnAction = "Sheet1.ExecuteFullImportProcess"
        .Font.Bold = True
    End With
    
    ' サブボタン1: Stagingクリア
    Set btn = ws.Buttons.Add(30, 280, 150, 25) ' B16セル位置
    With btn
        .Caption = "Stagingデータクリア"
        .OnAction = "Sheet1.ClearStagingData"
    End With
    
    ' サブボタン2: 設定画面
    Set btn = ws.Buttons.Add(30, 315, 150, 25) ' B17セル位置
    With btn
        .Caption = "設定画面を開く"
        .OnAction = "Sheet1.OpenConfigSheet"
    End With
    
    ' サブボタン3: レポート出力
    Set btn = ws.Buttons.Add(30, 350, 150, 25) ' B18セル位置
    With btn
        .Caption = "差分レポート出力"
        .OnAction = "Sheet1.ExportDifferenceReport"
    End With
    
    ' サブボタン4: KPI更新
    Set btn = ws.Buttons.Add(30, 385, 150, 25) ' B19セル位置
    With btn
        .Caption = "KPI表示更新"
        .OnAction = "Sheet1.RefreshKPI"
    End With
    
    ' サブボタン5: 期限切れ無効化
    Set btn = ws.Buttons.Add(220, 280, 150, 25) ' D16セル位置
    With btn
        .Caption = "期限切れ顧客無効化"
        .OnAction = "Sheet1.InactivateStaleCustomers"
    End With
    
    Exit Sub
    
ErrHandler:
    Call modCmn.LogError("CreateDashboardButtons", "ボタン作成エラー: " & Err.Description)
End Sub

'=============================================================================
' KPI表示・更新
'=============================================================================

' KPI表示更新
Public Sub RefreshKPI()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim customerCount As Long
    Dim stagingStats As Object
    Dim lastImportDate As Date
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_DASHBOARD)
    If ws Is Nothing Then Exit Sub
    
    ' 顧客総数取得
    customerCount = GetTotalCustomerCount()
    ws.Range(KPI_TOTAL_CUSTOMERS_CELL).Value = Format(customerCount, NUMBER_FORMAT_COUNT)
    
    ' Staging統計取得
    Set stagingStats = GetStagingStatistics()
    
    ' KPI値設定
    ws.Range(KPI_ADDED_COUNT_CELL).Value = Format(GetRecentAddedCount(), NUMBER_FORMAT_COUNT)
    ws.Range(KPI_UPDATED_COUNT_CELL).Value = Format(GetRecentUpdatedCount(), NUMBER_FORMAT_COUNT)
    ws.Range(KPI_DUPLICATE_COUNT_CELL).Value = Format(GetDuplicateDetectionCount(), NUMBER_FORMAT_COUNT)
    ws.Range(KPI_ERROR_COUNT_CELL).Value = Format(stagingStats("ErrorCount"), NUMBER_FORMAT_COUNT)
    ws.Range(KPI_INACTIVE_COUNT_CELL).Value = Format(GetInactiveCustomerCount(), NUMBER_FORMAT_COUNT)
    
    ' 最終取込日時
    lastImportDate = GetLastImportDateTime()
    If lastImportDate > 0 Then
        ws.Range(KPI_LAST_IMPORT_CELL).Value = Format(lastImportDate, DATE_FORMAT_DISPLAY)
    Else
        ws.Range(KPI_LAST_IMPORT_CELL).Value = "未実行"
    End If
    
    ' 処理時間
    ws.Range(KPI_PROCESS_TIME_CELL).Value = GetLastProcessTime() & " 秒"
    
    ' 更新日時記録
    ws.Range("D13").Value = "更新: " & Format(Now, DATE_FORMAT_DISPLAY)
    ws.Range("D13").Font.Size = 8
    ws.Range("D13").Font.Color = RGB(150, 150, 150)
    
    Call modCmn.LogInfo("RefreshKPI", "KPI表示更新完了")
    Exit Sub
    
ErrHandler:
    Call modCmn.LogError("RefreshKPI", "KPI更新エラー: " & Err.Description)
End Sub

'=============================================================================
' 統計データ取得
'=============================================================================

' 総顧客数取得
Private Function GetTotalCustomerCount() As Long
    On Error Resume Next
    
    Dim tbl As ListObject
    Set tbl = modData.GetCustomersTable()
    
    If Not tbl Is Nothing Then
        GetTotalCustomerCount = tbl.ListRows.Count
    Else
        GetTotalCustomerCount = 0
    End If
End Function

' Staging統計取得
Private Function GetStagingStatistics() As Object
    On Error Resume Next
    
    Dim stats As Object
    Dim tbl As ListObject
    Dim row As ListRow
    Dim totalCount As Long
    Dim validCount As Long
    Dim errorCount As Long
    
    Set GetStagingStatistics = CreateObject("Scripting.Dictionary")
    Set stats = GetStagingStatistics
    
    stats("TotalCount") = 0
    stats("ValidCount") = 0
    stats("ErrorCount") = 0
    
    Set tbl = modData.GetStagingTable()
    If tbl Is Nothing Then Exit Function
    
    For Each row In tbl.ListRows
        totalCount = totalCount + 1
        If modCmn.GetRowValue(row, COL_IS_VALID) = True Then
            validCount = validCount + 1
        Else
            errorCount = errorCount + 1
        End If
    Next row
    
    stats("TotalCount") = totalCount
    stats("ValidCount") = validCount
    stats("ErrorCount") = errorCount
End Function

' 最近の追加件数取得
Private Function GetRecentAddedCount() As Long
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    Dim addedCount As Long
    Dim today As Date
    
    GetRecentAddedCount = 0
    today = Date
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_LOGS)
    If ws Is Nothing Then Exit Function
    
    Set tbl = modCmn.GetTable(ws, TABLE_LOGS)
    If tbl Is Nothing Then Exit Function
    
    ' 本日のアップサートログから追加件数を集計
    For Each row In tbl.ListRows
        Dim logDate As Date
        Dim message As String
        
        logDate = modCmn.SafeDate(modCmn.GetRowText(row, "Timestamp"))
        message = modCmn.GetRowText(row, "Message")
        
        If DateValue(logDate) = today And InStr(message, "追加:") > 0 Then
            ' メッセージから追加件数を抽出（簡易パターンマッチング）
            Dim parts As Variant
            parts = Split(message, "追加:")
            If UBound(parts) > 0 Then
                Dim countPart As String
                countPart = Trim(Split(parts(1), ",")(0))
                addedCount = addedCount + modCmn.SafeLong(countPart)
            End If
        End If
    Next row
    
    GetRecentAddedCount = addedCount
End Function

' 最近の更新件数取得
Private Function GetRecentUpdatedCount() As Long
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    Dim updatedCount As Long
    Dim today As Date
    
    GetRecentUpdatedCount = 0
    today = Date
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_LOGS)
    If ws Is Nothing Then Exit Function
    
    Set tbl = modCmn.GetTable(ws, TABLE_LOGS)
    If tbl Is Nothing Then Exit Function
    
    ' 本日のアップサートログから更新件数を集計
    For Each row In tbl.ListRows
        Dim logDate As Date
        Dim message As String
        
        logDate = modCmn.SafeDate(modCmn.GetRowText(row, "Timestamp"))
        message = modCmn.GetRowText(row, "Message")
        
        If DateValue(logDate) = today And InStr(message, "更新:") > 0 Then
            ' メッセージから更新件数を抽出
            Dim parts As Variant
            parts = Split(message, "更新:")
            If UBound(parts) > 0 Then
                Dim countPart As String
                countPart = Trim(Split(parts(1), ",")(0))
                updatedCount = updatedCount + modCmn.SafeLong(countPart)
            End If
        End If
    Next row
    
    GetRecentUpdatedCount = updatedCount
End Function

' 重複検出件数取得
Private Function GetDuplicateDetectionCount() As Long
    On Error Resume Next
    
    Dim tbl As ListObject
    Dim row As ListRow
    Dim duplicateCount As Long
    
    GetDuplicateDetectionCount = 0
    
    Set tbl = modData.GetStagingTable()
    If tbl Is Nothing Then Exit Function
    
    For Each row In tbl.ListRows
        Dim errorMessage As String
        errorMessage = modCmn.GetRowText(row, COL_ERROR_MESSAGE)
        If InStr(errorMessage, "重複") > 0 Then
            duplicateCount = duplicateCount + 1
        End If
    Next row
    
    GetDuplicateDetectionCount = duplicateCount
End Function

' 無効顧客数取得
Private Function GetInactiveCustomerCount() As Long
    On Error Resume Next
    
    Dim tbl As ListObject
    Dim row As ListRow
    Dim inactiveCount As Long
    
    GetInactiveCustomerCount = 0
    
    Set tbl = modData.GetCustomersTable()
    If tbl Is Nothing Then Exit Function
    
    For Each row In tbl.ListRows
        If modCmn.GetRowText(row, COL_STATUS) = STATUS_INACTIVE Then
            inactiveCount = inactiveCount + 1
        End If
    Next row
    
    GetInactiveCustomerCount = inactiveCount
End Function

' 最終取込日時取得
Private Function GetLastImportDateTime() As Date
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    Dim lastDateTime As Date
    Dim rowDateTime As Date
    
    GetLastImportDateTime = 0
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_LOGS)
    If ws Is Nothing Then Exit Function
    
    Set tbl = modCmn.GetTable(ws, TABLE_LOGS)
    If tbl Is Nothing Then Exit Function
    
    ' ログから最新のCSV取込記録を検索
    For Each row In tbl.ListRows
        Dim message As String
        message = modCmn.GetRowText(row, "Message")
        If InStr(message, "CSV取り込み") > 0 Or InStr(message, "アップサート") > 0 Then
            rowDateTime = modCmn.SafeDate(modCmn.GetRowText(row, "Timestamp"))
            If rowDateTime > lastDateTime Then
                lastDateTime = rowDateTime
            End If
        End If
    Next row
    
    GetLastImportDateTime = lastDateTime
End Function

' 最終処理時間取得
Private Function GetLastProcessTime() As String
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    Dim lastDateTime As Date
    Dim lastProcessTime As String
    Dim rowDateTime As Date
    
    GetLastProcessTime = "N/A"
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_LOGS)
    If ws Is Nothing Then Exit Function
    
    Set tbl = modCmn.GetTable(ws, TABLE_LOGS)
    If tbl Is Nothing Then Exit Function
    
    ' 最新のアップサート処理時間を取得
    For Each row In tbl.ListRows
        Dim message As String
        message = modCmn.GetRowText(row, "Message")
        If InStr(message, "アップサート") > 0 Then
            rowDateTime = modCmn.SafeDate(modCmn.GetRowText(row, "Timestamp"))
            If rowDateTime > lastDateTime Then
                lastDateTime = rowDateTime
                lastProcessTime = modCmn.GetRowText(row, "ProcessTime")
            End If
        End If
    Next row
    
    GetLastProcessTime = lastProcessTime
End Function

'=============================================================================
' レポート生成
'=============================================================================

' 差分レポート生成・出力
Public Sub GenerateDifferenceReport()
    On Error GoTo ErrHandler
    
    Dim reportText As String
    Dim filePath As String
    Dim fileNum As Integer
    
    ' レポート内容生成
    reportText = CreateDifferenceReportContent()
    
    ' ファイル出力
    filePath = ThisWorkbook.Path & "\customer_report_" & Format(Now, DATE_FORMAT_FILE) & ".txt"
    fileNum = FreeFile
    
    Open filePath For Output As fileNum
    Print #fileNum, reportText
    Close fileNum
    
    MsgBox "差分レポートを出力しました。" & vbCrLf & filePath, vbInformation
    Call modCmn.LogInfo("GenerateDifferenceReport", "レポート出力完了: " & filePath)
    Exit Sub
    
ErrHandler:
    If fileNum > 0 Then Close fileNum
    Call modCmn.LogError("GenerateDifferenceReport", "レポート出力エラー: " & Err.Description)
End Sub

' 差分レポート内容作成
Private Function CreateDifferenceReportContent() As String
    On Error Resume Next
    
    Dim report As String
    Dim stagingStats As Object
    
    Set stagingStats = GetStagingStatistics()
    
    report = "=== 顧客データ管理システム 処理レポート ===" & vbCrLf & vbCrLf
    report = report & "レポート作成日時: " & modCmn.GetCurrentDateTimeString() & vbCrLf
    report = report & "システムバージョン: " & SYSTEM_VERSION & vbCrLf & vbCrLf
    
    report = report & "【現在の状況】" & vbCrLf
    report = report & "総顧客数: " & Format(GetTotalCustomerCount(), NUMBER_FORMAT_COUNT) & " 件" & vbCrLf
    report = report & "有効顧客数: " & Format(GetTotalCustomerCount() - GetInactiveCustomerCount(), NUMBER_FORMAT_COUNT) & " 件" & vbCrLf
    report = report & "無効顧客数: " & Format(GetInactiveCustomerCount(), NUMBER_FORMAT_COUNT) & " 件" & vbCrLf & vbCrLf
    
    report = report & "【最新処理結果】" & vbCrLf
    report = report & "本日追加件数: " & Format(GetRecentAddedCount(), NUMBER_FORMAT_COUNT) & " 件" & vbCrLf
    report = report & "本日更新件数: " & Format(GetRecentUpdatedCount(), NUMBER_FORMAT_COUNT) & " 件" & vbCrLf
    report = report & "重複検出件数: " & Format(GetDuplicateDetectionCount(), NUMBER_FORMAT_COUNT) & " 件" & vbCrLf
    report = report & "エラー件数: " & Format(stagingStats("ErrorCount"), NUMBER_FORMAT_COUNT) & " 件" & vbCrLf & vbCrLf
    
    report = report & "【処理履歴】" & vbCrLf
    Dim lastImport As Date
    lastImport = GetLastImportDateTime()
    If lastImport > 0 Then
        report = report & "最終取込日時: " & Format(lastImport, DATE_FORMAT_DISPLAY) & vbCrLf
        report = report & "処理時間: " & GetLastProcessTime() & vbCrLf
    Else
        report = report & "取込履歴: なし" & vbCrLf
    End If
    
    report = report & vbCrLf & "【検証結果詳細】" & vbCrLf
    report = report & modValidation.GenerateValidationReport()
    
    CreateDifferenceReportContent = report
End Function

'=============================================================================
' ユーティリティ関数
'=============================================================================

' システム状態チェック
Public Function CheckSystemHealth() As Boolean
    On Error Resume Next
    
    CheckSystemHealth = True
    
    ' 必要なシートの存在チェック
    If modCmn.GetWorksheetByIndex(SHEET_INDEX_CUSTOMERS) Is Nothing Then CheckSystemHealth = False
    If modCmn.GetWorksheetByIndex(SHEET_INDEX_STAGING) Is Nothing Then CheckSystemHealth = False
    If modCmn.GetWorksheetByIndex(SHEET_INDEX_CONFIG) Is Nothing Then CheckSystemHealth = False
    If modCmn.GetWorksheetByIndex(SHEET_INDEX_LOGS) Is Nothing Then CheckSystemHealth = False
    
    ' 必要なテーブルの存在チェック
    If modData.GetCustomersTable() Is Nothing Then CheckSystemHealth = False
    If modData.GetStagingTable() Is Nothing Then CheckSystemHealth = False
End Function

' ダッシュボード表示メッセージ
Public Sub ShowStatusMessage(ByVal message As String, Optional ByVal isError As Boolean = False)
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_DASHBOARD)
    If ws Is Nothing Then Exit Sub
    
    With ws.Range("A20")
        .Value = message
        .Font.Size = 10
        .Font.Bold = True
        If isError Then
            .Font.Color = RGB(200, 0, 0)
        Else
            .Font.Color = RGB(0, 150, 0)
        End If
    End With
    
    ' 一定時間後にメッセージをクリア
    Application.OnTime Now + TimeValue("00:00:05"), "modDashboard.ClearStatusMessage"
End Sub

' ステータスメッセージクリア
Public Sub ClearStatusMessage()
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_DASHBOARD)
    If Not ws Is Nothing Then
        ws.Range("A20").ClearContents
    End If
End Sub

' 処理確認ダイアログ
Public Function ConfirmOperation(ByVal operation As String) As Boolean
    On Error Resume Next
    
    Dim result As VbMsgBoxResult
    result = MsgBox("以下の操作を実行しますか？" & vbCrLf & vbCrLf & operation & vbCrLf & vbCrLf & _
                    "この操作は取り消すことができません。", vbQuestion + vbYesNo, "操作確認")
    
    ConfirmOperation = (result = vbYes)
End Function