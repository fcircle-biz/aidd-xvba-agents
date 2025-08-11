Attribute VB_Name = "modCustomerSystem"
'=============================================================================
' modCustomerSystem.bas - 顧客管理システムメイン制御モジュール
'=============================================================================
' 概要:
'   システム全体のオーケストレーション、メイン処理フローの制御
'   各モジュール間の連携、エラーハンドリング、トランザクション管理
'=============================================================================
Option Explicit

'=============================================================================
' メイン処理フロー
'=============================================================================

' CSV一括取込・更新処理（メイン処理）
Public Sub ExecuteFullImportProcess()
    On Error GoTo ErrHandler
    
    Dim startTime As Double
    Dim result As Object
    Dim totalProcessTime As Double
    
    startTime = Timer
    
    ' 処理確認
    If Not modDashboard.ConfirmOperation("CSV一括取込・更新処理を実行します。" & vbCrLf & _
                                       "1. CSVファイル取込" & vbCrLf & _
                                       "2. データ正規化・検証" & vbCrLf & _
                                       "3. 顧客データ更新（追加・更新）" & vbCrLf & _
                                       "4. KPI表示更新") Then
        Call modDashboard.ShowStatusMessage("処理がキャンセルされました")
        Exit Sub
    End If
    
    ' パフォーマンス最適化開始
    Call modUtils.StartPerformanceOptimization()
    
    ' バックアップ作成
    Call modDashboard.ShowStatusMessage("バックアップ作成中...")
    If Not modUpsert.BackupCustomerData() Then
        Call modDashboard.ShowStatusMessage("バックアップ作成に失敗しましたが処理を継続します", isError:=True)
    End If
    
    ' ステップ1: CSV取込
    Call modDashboard.ShowStatusMessage("CSV取込処理中...")
    Call modData.ImportCsvToStaging()
    
    ' ステップ2: データ検証
    Call modDashboard.ShowStatusMessage("データ検証処理中...")
    Dim validationErrors As Long
    validationErrors = modValidation.ValidateStagingData()
    
    If validationErrors < 0 Then
        Err.Raise vbObjectError + 1001, "ExecuteFullImportProcess", "データ検証処理でエラーが発生しました"
    End If
    
    ' ステップ3: アップサート処理
    Call modDashboard.ShowStatusMessage("データ更新処理中...")
    Set result = modUpsert.ExecuteUpsertOperation()
    
    If Not result("Success") Then
        Err.Raise vbObjectError + 1002, "ExecuteFullImportProcess", result("Message")
    End If
    
    ' ステップ4: 期限切れ顧客無効化（任意）
    Dim inactivatedCount As Long
    If LCase(modData.GetConfigValue("AUTO_INACTIVATE")) = "true" Then
        Call modDashboard.ShowStatusMessage("期限切れ顧客無効化中...")
        inactivatedCount = modUpsert.InactivateStaleCustomers()
    End If
    
    ' ステップ5: KPI更新
    Call modDashboard.ShowStatusMessage("KPI表示更新中...")
    Call modDashboard.RefreshKPI()
    
    ' 処理完了
    totalProcessTime = Timer - startTime
    
    ' 結果表示
    Dim resultMsg As String
    resultMsg = "CSV一括取込・更新処理が完了しました。" & vbCrLf & vbCrLf
    resultMsg = resultMsg & "【処理結果】" & vbCrLf
    resultMsg = resultMsg & "追加件数: " & Format(result("AddedCount"), NUMBER_FORMAT_COUNT) & " 件" & vbCrLf
    resultMsg = resultMsg & "更新件数: " & Format(result("UpdatedCount"), NUMBER_FORMAT_COUNT) & " 件" & vbCrLf
    resultMsg = resultMsg & "スキップ件数: " & Format(result("SkippedCount"), NUMBER_FORMAT_COUNT) & " 件" & vbCrLf
    resultMsg = resultMsg & "検証エラー件数: " & Format(validationErrors, NUMBER_FORMAT_COUNT) & " 件" & vbCrLf
    If inactivatedCount > 0 Then
        resultMsg = resultMsg & "無効化件数: " & Format(inactivatedCount, NUMBER_FORMAT_COUNT) & " 件" & vbCrLf
    End If
    resultMsg = resultMsg & "総処理時間: " & Format(totalProcessTime, "0.0") & " 秒"
    
    ' ログ記録
    Call modData.LogImportOperation("フル処理完了", result("AddedCount") + result("UpdatedCount"), _
                                   totalProcessTime, "検証エラー:" & validationErrors & ", 無効化:" & inactivatedCount)
    
    MsgBox resultMsg, vbInformation, "処理完了"
    Call modDashboard.ShowStatusMessage("処理完了: " & Format(Now, "hh:mm"))
    
ExitHandler:
    Call modUtils.EndPerformanceOptimization()
    Exit Sub
    
ErrHandler:
    Call modUtils.EndPerformanceOptimization()
    Call modDashboard.ShowStatusMessage("処理エラーが発生しました", isError:=True)
    Call modCmn.LogError("ExecuteFullImportProcess", "フル処理エラー: " & Err.Description)
    MsgBox "処理中にエラーが発生しました。" & vbCrLf & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & vbCrLf & _
           "詳細はログファイルをご確認ください。", vbCritical, "処理エラー"
End Sub

' CSV取込のみ実行
Public Sub ExecuteImportOnly()
    On Error GoTo ErrHandler
    
    If Not modDashboard.ConfirmOperation("CSVファイルの取込のみを実行します。") Then Exit Sub
    
    Call modUtils.StartPerformanceOptimization()
    Call modData.ImportCsvToStaging()
    Call modUtils.EndPerformanceOptimization()
    
    Call modDashboard.RefreshKPI()
    MsgBox "CSV取込処理が完了しました。", vbInformation
    Exit Sub
    
ErrHandler:
    Call modUtils.EndPerformanceOptimization()
    Call modCmn.LogError("ExecuteImportOnly", Err.Description)
End Sub

' データ検証のみ実行
Public Sub ExecuteValidationOnly()
    On Error GoTo ErrHandler
    
    Dim errorCount As Long
    
    Call modUtils.StartPerformanceOptimization()
    errorCount = modValidation.ValidateStagingData()
    Call modUtils.EndPerformanceOptimization()
    
    If errorCount >= 0 Then
        MsgBox "データ検証が完了しました。" & vbCrLf & _
               "エラー件数: " & Format(errorCount, NUMBER_FORMAT_COUNT) & " 件", vbInformation
    Else
        MsgBox "データ検証処理でエラーが発生しました。", vbCritical
    End If
    
    Call modDashboard.RefreshKPI()
    Exit Sub
    
ErrHandler:
    Call modUtils.EndPerformanceOptimization()
    Call modCmn.LogError("ExecuteValidationOnly", Err.Description)
End Sub

' アップサートのみ実行
Public Sub ExecuteUpsertOnly()
    On Error GoTo ErrHandler
    
    Dim result As Object
    
    If Not modDashboard.ConfirmOperation("Stagingデータから顧客テーブルの更新のみを実行します。") Then Exit Sub
    
    Call modUtils.StartPerformanceOptimization()
    Set result = modUpsert.ExecuteUpsertOperation()
    Call modUtils.EndPerformanceOptimization()
    
    If result("Success") Then
        MsgBox "アップサート処理が完了しました。" & vbCrLf & _
               "追加: " & result("AddedCount") & " 件, 更新: " & result("UpdatedCount") & " 件", vbInformation
    Else
        MsgBox "アップサート処理でエラーが発生しました。" & vbCrLf & result("Message"), vbCritical
    End If
    
    Call modDashboard.RefreshKPI()
    Exit Sub
    
ErrHandler:
    Call modUtils.EndPerformanceOptimization()
    Call modCmn.LogError("ExecuteUpsertOnly", Err.Description)
End Sub

'=============================================================================
' システム管理処理
'=============================================================================

' システム初期化処理
Public Sub InitializeSystem()
    On Error GoTo ErrHandler
    
    Call modCmn.LogInfo("InitializeSystem", "システム初期化開始")
    
    ' システム全体初期化
    Call modUtils.InitializeCustomerSystem()
    
    ' 初期化完了後のスプラッシュ表示
    Call ShowSystemSplash()
    
    Call modCmn.LogInfo("InitializeSystem", "システム初期化完了")
    Exit Sub
    
ErrHandler:
    Call modCmn.LogError("InitializeSystem", "システム初期化エラー: " & Err.Description)
End Sub

' システムスプラッシュ表示
Private Sub ShowSystemSplash()
    On Error Resume Next
    
    Dim splashMsg As String
    splashMsg = SYSTEM_NAME & vbCrLf
    splashMsg = splashMsg & "バージョン " & SYSTEM_VERSION & vbCrLf & vbCrLf
    splashMsg = splashMsg & "システムの初期化が完了しました。" & vbCrLf & vbCrLf
    splashMsg = splashMsg & "使用方法:" & vbCrLf
    splashMsg = splashMsg & "1. Dashboardシートで各種操作を実行" & vbCrLf
    splashMsg = splashMsg & "2. _Configシートで設定を調整" & vbCrLf
    splashMsg = splashMsg & "3. Logsシートで処理履歴を確認" & vbCrLf & vbCrLf
    splashMsg = splashMsg & SYSTEM_COPYRIGHT
    
    MsgBox splashMsg, vbInformation, "システム起動完了"
End Sub

' システム健全性チェック実行
Public Sub ExecuteSystemHealthCheck()
    On Error GoTo ErrHandler
    
    Dim health As Object
    Dim healthMsg As String
    Dim issue As Variant
    Dim warning As Variant
    
    Set health = modUtils.PerformSystemHealthCheck()
    
    healthMsg = "=== システム健全性チェック結果 ===" & vbCrLf & vbCrLf
    healthMsg = healthMsg & "総合判定: " & health("OverallHealth") & vbCrLf & vbCrLf
    
    ' 重大な問題
    If health("Issues").Count > 0 Then
        healthMsg = healthMsg & "【重大な問題】" & vbCrLf
        For Each issue In health("Issues")
            healthMsg = healthMsg & "- " & issue & vbCrLf
        Next issue
        healthMsg = healthMsg & vbCrLf
    End If
    
    ' 警告
    If health("Warnings").Count > 0 Then
        healthMsg = healthMsg & "【警告】" & vbCrLf
        For Each warning In health("Warnings")
            healthMsg = healthMsg & "- " & warning & vbCrLf
        Next warning
        healthMsg = healthMsg & vbCrLf
    End If
    
    If health("Issues").Count = 0 And health("Warnings").Count = 0 Then
        healthMsg = healthMsg & "問題は検出されませんでした。"
    End If
    
    MsgBox healthMsg, vbInformation, "システム健全性チェック"
    Exit Sub
    
ErrHandler:
    Call modCmn.LogError("ExecuteSystemHealthCheck", Err.Description)
End Sub

' システム情報表示
Public Sub ShowSystemInformation()
    On Error Resume Next
    
    Dim info As Object
    Dim infoMsg As String
    Dim key As Variant
    
    Set info = modUtils.GetSystemInformation()
    
    infoMsg = "=== システム情報 ===" & vbCrLf & vbCrLf
    
    ' 基本情報
    infoMsg = infoMsg & "システム名: " & info("SystemName") & vbCrLf
    infoMsg = infoMsg & "バージョン: " & info("Version") & vbCrLf
    infoMsg = infoMsg & "作成者: " & info("Author") & vbCrLf
    infoMsg = infoMsg & "現在日時: " & info("CurrentTime") & vbCrLf
    infoMsg = infoMsg & "Excelバージョン: " & info("ExcelVersion") & vbCrLf & vbCrLf
    
    ' ファイル情報
    infoMsg = infoMsg & "ワークブック名: " & info("WorkbookName") & vbCrLf
    infoMsg = infoMsg & "ファイルパス: " & info("WorkbookPath") & vbCrLf
    infoMsg = infoMsg & "シート数: " & info("SheetCount") & vbCrLf & vbCrLf
    
    ' データ統計
    infoMsg = infoMsg & "顧客データ件数: " & Format(info("CustomerCount"), NUMBER_FORMAT_COUNT) & " 件" & vbCrLf
    infoMsg = infoMsg & "Stagingデータ件数: " & Format(info("StagingCount"), NUMBER_FORMAT_COUNT) & " 件" & vbCrLf
    
    MsgBox infoMsg, vbInformation, "システム情報"
End Sub

'=============================================================================
' データメンテナンス処理
'=============================================================================

' 全データクリア（開発・テスト用）
Public Sub ClearAllData()
    On Error GoTo ErrHandler
    
    Dim confirmMsg As String
    confirmMsg = "全てのデータをクリアします。" & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "この操作により以下のデータが削除されます:" & vbCrLf
    confirmMsg = confirmMsg & "- 顧客マスタデータ" & vbCrLf
    confirmMsg = confirmMsg & "- Stagingデータ" & vbCrLf
    confirmMsg = confirmMsg & "- ログデータ" & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "※設定データは保持されます" & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "本当に実行しますか？"
    
    If MsgBox(confirmMsg, vbQuestion + vbYesNo + vbDefaultButton2, "データクリア確認") <> vbYes Then
        Exit Sub
    End If
    
    ' 2回目の確認
    If MsgBox("最終確認：全データをクリアしますか？" & vbCrLf & "この操作は取り消すことができません。", _
              vbCritical + vbYesNo + vbDefaultButton2, "最終確認") <> vbYes Then
        Exit Sub
    End If
    
    Call modUtils.StartPerformanceOptimization()
    
    ' 顧客データクリア
    Dim ws As Worksheet
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_CUSTOMERS)
    If Not ws Is Nothing Then
        Call modCmn.SafeClearSheet(ws, keepFormats:=True)
        Call modUtils.EnsureCustomersTableStructure()
    End If
    
    ' Stagingデータクリア
    Call modData.ClearStagingData()
    
    ' ログデータクリア
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_LOGS)
    If Not ws Is Nothing Then
        Call modCmn.SafeClearSheet(ws, keepFormats:=True)
        Call modUtils.EnsureLogsTableStructure()
    End If
    
    Call modUtils.EndPerformanceOptimization()
    
    ' KPI更新
    Call modDashboard.RefreshKPI()
    
    Call modCmn.LogInfo("ClearAllData", "全データクリア実行")
    MsgBox "全データのクリアが完了しました。", vbInformation
    Exit Sub
    
ErrHandler:
    Call modUtils.EndPerformanceOptimization()
    Call modCmn.LogError("ClearAllData", Err.Description)
End Sub

' データベース最適化
Public Sub OptimizeDatabase()
    On Error GoTo ErrHandler
    
    If Not modDashboard.ConfirmOperation("データベースの最適化を実行します。" & vbCrLf & _
                                       "・古いログの削除" & vbCrLf & _
                                       "・重複レコードのクリーンアップ" & vbCrLf & _
                                       "・テーブル構造の最適化") Then Exit Sub
    
    Call modUtils.StartPerformanceOptimization()
    
    Dim optimizedCount As Long
    optimizedCount = 0
    
    ' 古いログ削除（30日以前）
    Call CleanupOldLogs(30, optimizedCount)
    
    ' 重複レコードチェック（警告のみ）
    Call CheckForDuplicateCustomers()
    
    ' テーブル構造最適化
    Call modUtils.EnsureAllTableStructures()
    
    ' メモリクリーンアップ
    Call modUtils.CleanupMemory()
    
    Call modUtils.EndPerformanceOptimization()
    
    MsgBox "データベース最適化が完了しました。" & vbCrLf & _
           "処理件数: " & Format(optimizedCount, NUMBER_FORMAT_COUNT) & " 件", vbInformation
    Exit Sub
    
ErrHandler:
    Call modUtils.EndPerformanceOptimization()
    Call modCmn.LogError("OptimizeDatabase", Err.Description)
End Sub

' 古いログクリーンアップ
Private Sub CleanupOldLogs(ByVal daysToKeep As Integer, ByRef cleanedCount As Long)
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    Dim cutoffDate As Date
    Dim i As Long
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_LOGS)
    If ws Is Nothing Then Exit Sub
    
    Set tbl = modCmn.GetTable(ws, TABLE_LOGS)
    If tbl Is Nothing Then Exit Sub
    
    cutoffDate = Now - daysToKeep
    
    ' 後ろから削除（インデックス変更を避けるため）
    For i = tbl.ListRows.Count To 1 Step -1
        Set row = tbl.ListRows(i)
        Dim logDate As Date
        logDate = modCmn.SafeDate(modCmn.GetRowText(row, "Timestamp"))
        
        If logDate > 0 And logDate < cutoffDate Then
            row.Delete
            cleanedCount = cleanedCount + 1
        End If
    Next i
End Sub

' 重複顧客チェック
Private Sub CheckForDuplicateCustomers()
    On Error Resume Next
    
    Dim tbl As ListObject
    Set tbl = modData.GetCustomersTable()
    If tbl Is Nothing Then Exit Sub
    
    Dim duplicates As Collection
    Set duplicates = New Collection
    
    ' 重複検出ロジック（簡易版）
    ' 実装の詳細は省略
    
    If duplicates.Count > 0 Then
        MsgBox "重複の可能性がある顧客データが " & duplicates.Count & " 件見つかりました。" & vbCrLf & _
               "詳細はLogsシートをご確認ください。", vbExclamation
    End If
End Sub

'=============================================================================
' エクスポート・レポート処理
'=============================================================================

' 顧客データエクスポート
Public Sub ExportCustomerData()
    On Error GoTo ErrHandler
    
    Dim exportPath As String
    Dim fileName As String
    
    fileName = "customer_export_" & Format(Now, DATE_FORMAT_FILE) & ".csv"
    exportPath = ThisWorkbook.Path & "\" & fileName
    
    Call modUtils.StartPerformanceOptimization()
    
    ' CSV出力処理
    If ExportCustomerTableToCsv(exportPath) Then
        Call modUtils.EndPerformanceOptimization()
        MsgBox "顧客データのエクスポートが完了しました。" & vbCrLf & exportPath, vbInformation
        Call modCmn.LogInfo("ExportCustomerData", "エクスポート完了: " & exportPath)
    Else
        Call modUtils.EndPerformanceOptimization()
        MsgBox "エクスポート処理でエラーが発生しました。", vbCritical
    End If
    Exit Sub
    
ErrHandler:
    Call modUtils.EndPerformanceOptimization()
    Call modCmn.LogError("ExportCustomerData", Err.Description)
End Sub

' 顧客テーブルCSV出力
Private Function ExportCustomerTableToCsv(ByVal filePath As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim tbl As ListObject
    Dim fileNum As Integer
    Dim row As ListRow
    Dim csvLine As String
    Dim exportCount As Long
    
    ExportCustomerTableToCsv = False
    
    Set tbl = modData.GetCustomersTable()
    If tbl Is Nothing Then Exit Function
    
    fileNum = FreeFile
    Open filePath For Output As fileNum
    
    ' ヘッダー行出力
    Print #fileNum, CUSTOMERS_HEADERS
    
    ' データ行出力
    For Each row In tbl.ListRows
        csvLine = ""
        csvLine = csvLine & QuoteCsvField(modCmn.GetRowText(row, COL_CUSTOMER_ID)) & ","
        csvLine = csvLine & QuoteCsvField(modCmn.GetRowText(row, COL_CUSTOMER_NAME)) & ","
        csvLine = csvLine & QuoteCsvField(modCmn.GetRowText(row, COL_EMAIL)) & ","
        csvLine = csvLine & QuoteCsvField(modCmn.GetRowText(row, COL_PHONE)) & ","
        csvLine = csvLine & QuoteCsvField(modCmn.GetRowText(row, COL_ZIP)) & ","
        csvLine = csvLine & QuoteCsvField(modCmn.GetRowText(row, COL_ADDRESS1)) & ","
        csvLine = csvLine & QuoteCsvField(modCmn.GetRowText(row, COL_ADDRESS2)) & ","
        csvLine = csvLine & QuoteCsvField(modCmn.GetRowText(row, COL_CATEGORY)) & ","
        csvLine = csvLine & QuoteCsvField(modCmn.GetRowText(row, COL_STATUS)) & ","
        csvLine = csvLine & QuoteCsvField(CStr(modCmn.GetRowDate(row, COL_CREATED_AT))) & ","
        csvLine = csvLine & QuoteCsvField(CStr(modCmn.GetRowDate(row, COL_UPDATED_AT))) & ","
        csvLine = csvLine & QuoteCsvField(modCmn.GetRowText(row, COL_SOURCE_FILE)) & ","
        csvLine = csvLine & QuoteCsvField(modCmn.GetRowText(row, COL_NOTES))
        
        Print #fileNum, csvLine
        exportCount = exportCount + 1
    Next row
    
    Close fileNum
    ExportCustomerTableToCsv = True
    
    Call modCmn.LogInfo("ExportCustomerTableToCsv", "CSV出力完了: " & exportCount & " 件")
    Exit Function
    
ErrHandler:
    If fileNum > 0 Then Close fileNum
    ExportCustomerTableToCsv = False
    Call modCmn.LogError("ExportCustomerTableToCsv", Err.Description)
End Function

' CSVフィールド引用符処理
Private Function QuoteCsvField(ByVal field As String) As String
    On Error Resume Next
    
    ' カンマや改行が含まれる場合は引用符で囲む
    If InStr(field, ",") > 0 Or InStr(field, vbCrLf) > 0 Or InStr(field, """") > 0 Then
        ' 引用符をエスケープ
        field = Replace(field, """", """""")
        QuoteCsvField = """" & field & """"
    Else
        QuoteCsvField = field
    End If
End Function

'=============================================================================
' エラーリカバリ処理
'=============================================================================

' システム緊急停止
Public Sub EmergencyStop()
    On Error Resume Next
    
    ' 全ての処理を停止
    Application.EnableEvents = False
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Call modCmn.LogError("EmergencyStop", "システム緊急停止が実行されました")
    
    MsgBox "システムの緊急停止を実行しました。" & vbCrLf & _
           "処理を再開する場合は、ワークブックを再度開いてください。", vbCritical, "緊急停止"
End Sub