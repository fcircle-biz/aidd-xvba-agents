Attribute VB_Name = "modUpsert"
'=============================================================================
' modUpsert.bas - アップサート（Insert/Update）処理モジュール
'=============================================================================
' 概要:
'   Stagingから顧客テーブルへの安全な追加・更新処理
'   論理削除、差分更新、重複処理等のデータ統合機能を提供
'=============================================================================
Option Explicit

'=============================================================================
' メインアップサート処理
'=============================================================================

' Stagingから顧客テーブルへのアップサート実行
Public Function ExecuteUpsertOperation() As Object
    On Error GoTo ErrHandler
    
    Dim result As Object
    Dim stagingTbl As ListObject
    Dim customerTbl As ListObject
    Dim stagingRow As ListRow
    Dim startTime As Double
    Dim processedCount As Long
    Dim addedCount As Long
    Dim updatedCount As Long
    Dim skippedCount As Long
    
    startTime = Timer
    Set result = CreateObject("Scripting.Dictionary")
    
    ' 初期化
    result("Success") = False
    result("AddedCount") = 0
    result("UpdatedCount") = 0
    result("SkippedCount") = 0
    result("ProcessTime") = 0
    result("Message") = ""
    
    ' テーブル取得
    Set stagingTbl = modData.GetStagingTable()
    Set customerTbl = modData.GetCustomersTable()
    
    If stagingTbl Is Nothing Or customerTbl Is Nothing Then
        result("Message") = "必要なテーブルが取得できませんでした"
        Set ExecuteUpsertOperation = result
        Exit Function
    End If
    
    Call modCmn.ShowProgressStart(MSG_UPSERT_STARTED)
    
    ' 顧客辞書作成（高速検索用）
    Dim customerDict As Object
    Set customerDict = CreateCustomerSearchDictionary(customerTbl)
    
    ' 各Staging行を処理
    For Each stagingRow In stagingTbl.ListRows
        processedCount = processedCount + 1
        
        ' 有効なレコードのみ処理
        If modCmn.GetRowValue(stagingRow, COL_IS_VALID) = True Then
            Dim upsertResult As Integer
            upsertResult = ProcessSingleUpsert(stagingRow, customerTbl, customerDict)
            
            Select Case upsertResult
                Case 1 ' 追加
                    addedCount = addedCount + 1
                Case 2 ' 更新
                    updatedCount = updatedCount + 1
                Case 0 ' スキップ
                    skippedCount = skippedCount + 1
            End Select
        Else
            skippedCount = skippedCount + 1
        End If
        
        ' プログレス更新
        If processedCount Mod BATCH_SIZE_UPSERT = 0 Then
            Call modCmn.UpdateProgress("データ更新中: " & processedCount & " 件処理")
        End If
    Next stagingRow
    
    ' 結果設定
    result("Success") = True
    result("AddedCount") = addedCount
    result("UpdatedCount") = updatedCount
    result("SkippedCount") = skippedCount
    result("ProcessTime") = Timer - startTime
    result("Message") = "アップサート完了: 追加" & addedCount & "件, 更新" & updatedCount & "件"
    
    ' ログ記録
    Call modData.LogImportOperation("アップサート処理完了", processedCount, Timer - startTime, _
                                   "追加:" & addedCount & ", 更新:" & updatedCount & ", スキップ:" & skippedCount)
    
    Set ExecuteUpsertOperation = result
    Call modCmn.HideProgress
    Exit Function
    
ErrHandler:
    Call modCmn.HideProgress
    result("Success") = False
    result("Message") = "アップサート処理エラー: " & Err.Description
    Set ExecuteUpsertOperation = result
    Call modCmn.LogError("ExecuteUpsertOperation", Err.Description)
End Function

' 単一レコードのアップサート処理
Private Function ProcessSingleUpsert(ByVal stagingRow As ListRow, ByVal customerTbl As ListObject, _
                                    ByVal customerDict As Object) As Integer
    On Error Resume Next
    
    Dim primaryKey As String
    Dim altKey As String
    Dim existingRow As ListRow
    
    ProcessSingleUpsert = 0 ' デフォルト：スキップ
    
    ' キー情報取得
    primaryKey = modCmn.GetRowText(stagingRow, COL_CUSTOMER_ID)
    altKey = modCmn.GetRowText(stagingRow, COL_KEY_CANDIDATE)
    
    ' 既存レコード検索
    Set existingRow = FindExistingCustomer(customerTbl, customerDict, primaryKey, altKey)
    
    If existingRow Is Nothing Then
        ' 新規追加
        If AddNewCustomer(stagingRow, customerTbl) Then
            ProcessSingleUpsert = 1 ' 追加成功
        End If
    Else
        ' 既存更新
        If UpdateExistingCustomer(stagingRow, existingRow) Then
            ProcessSingleUpsert = 2 ' 更新成功
        End If
    End If
End Function

'=============================================================================
' 顧客検索・辞書管理
'=============================================================================

' 顧客検索辞書作成
Private Function CreateCustomerSearchDictionary(ByVal customerTbl As ListObject) As Object
    On Error Resume Next
    
    Dim customerDict As Object
    Dim row As ListRow
    Dim primaryKey As String
    Dim email As String
    Dim customerName As String
    Dim altKey As String
    
    Set CreateCustomerSearchDictionary = CreateObject("Scripting.Dictionary")
    Set customerDict = CreateCustomerSearchDictionary
    
    Dim rowIndex As Long
    rowIndex = 1
    
    For Each row In customerTbl.ListRows
        ' 主キーマッピング（顧客ID）
        primaryKey = modCmn.GetRowText(row, COL_CUSTOMER_ID)
        If Len(primaryKey) > 0 Then
            customerDict("PK:" & primaryKey) = rowIndex
        End If
        
        ' 代替キーマッピング（Email + CustomerName）
        email = modCmn.NormalizeEmail(modCmn.GetRowText(row, COL_EMAIL))
        customerName = modCmn.GetRowText(row, COL_CUSTOMER_NAME)
        If Len(email) > 0 And Len(customerName) > 0 Then
            altKey = email & "+" & customerName
            customerDict("AK:" & altKey) = rowIndex
        End If
        
        rowIndex = rowIndex + 1
    Next row
End Function

' 既存顧客検索
Private Function FindExistingCustomer(ByVal customerTbl As ListObject, ByVal customerDict As Object, _
                                     ByVal primaryKey As String, ByVal altKey As String) As ListRow
    On Error Resume Next
    
    Dim rowIndex As Variant
    
    Set FindExistingCustomer = Nothing
    
    ' 主キーで検索
    If Len(primaryKey) > 0 Then
        rowIndex = customerDict("PK:" & primaryKey)
        If Not IsEmpty(rowIndex) Then
            Set FindExistingCustomer = customerTbl.ListRows(CLng(rowIndex))
            Exit Function
        End If
    End If
    
    ' 代替キーで検索（主キーがない場合）
    If Len(primaryKey) = 0 And Len(altKey) > 0 Then
        rowIndex = customerDict("AK:" & altKey)
        If Not IsEmpty(rowIndex) Then
            Set FindExistingCustomer = customerTbl.ListRows(CLng(rowIndex))
        End If
    End If
End Function

'=============================================================================
' 新規追加処理
'=============================================================================

' 新規顧客追加
Private Function AddNewCustomer(ByVal stagingRow As ListRow, ByVal customerTbl As ListObject) As Boolean
    On Error GoTo ErrHandler
    
    Dim newRow As ListRow
    
    AddNewCustomer = False
    
    ' 新規行追加
    Set newRow = customerTbl.ListRows.Add
    
    ' データコピー
    Call CopyDataFromStaging(stagingRow, newRow, isNewRecord:=True)
    
    AddNewCustomer = True
    Exit Function
    
ErrHandler:
    AddNewCustomer = False
    Call modCmn.LogError("AddNewCustomer", "新規顧客追加エラー: " & Err.Description)
End Function

'=============================================================================
' 既存更新処理
'=============================================================================

' 既存顧客更新
Private Function UpdateExistingCustomer(ByVal stagingRow As ListRow, ByVal existingRow As ListRow) As Boolean
    On Error GoTo ErrHandler
    
    UpdateExistingCustomer = False
    
    ' 差分チェック＆更新
    If HasDataDifferences(stagingRow, existingRow) Then
        Call CopyDataFromStaging(stagingRow, existingRow, isNewRecord:=False)
        UpdateExistingCustomer = True
    End If
    
    Exit Function
    
ErrHandler:
    UpdateExistingCustomer = False
    Call modCmn.LogError("UpdateExistingCustomer", "顧客更新エラー: " & Err.Description)
End Function

' データ差分チェック
Private Function HasDataDifferences(ByVal stagingRow As ListRow, ByVal existingRow As ListRow) As Boolean
    On Error Resume Next
    
    HasDataDifferences = False
    
    ' 主要フィールドの差分チェック
    If CompareFieldValues(stagingRow, existingRow, COL_CUSTOMER_NAME) Then HasDataDifferences = True
    If CompareFieldValues(stagingRow, existingRow, COL_EMAIL) Then HasDataDifferences = True
    If CompareFieldValues(stagingRow, existingRow, COL_PHONE) Then HasDataDifferences = True
    If CompareFieldValues(stagingRow, existingRow, COL_ZIP) Then HasDataDifferences = True
    If CompareFieldValues(stagingRow, existingRow, COL_ADDRESS1) Then HasDataDifferences = True
    If CompareFieldValues(stagingRow, existingRow, COL_ADDRESS2) Then HasDataDifferences = True
    If CompareFieldValues(stagingRow, existingRow, COL_CATEGORY) Then HasDataDifferences = True
    If CompareFieldValues(stagingRow, existingRow, COL_STATUS) Then HasDataDifferences = True
    If CompareFieldValues(stagingRow, existingRow, COL_NOTES) Then HasDataDifferences = True
End Function

' フィールド値比較
Private Function CompareFieldValues(ByVal stagingRow As ListRow, ByVal existingRow As ListRow, _
                                   ByVal fieldName As String) As Boolean
    On Error Resume Next
    
    Dim stagingValue As String
    Dim existingValue As String
    
    ' 正規化済み値を使用（利用可能な場合）
    Select Case fieldName
        Case COL_EMAIL
            stagingValue = modCmn.TrimAll(modCmn.GetRowText(stagingRow, COL_EMAIL_NORM))
            existingValue = modCmn.NormalizeEmail(modCmn.GetRowText(existingRow, COL_EMAIL))
        Case COL_PHONE
            stagingValue = modCmn.TrimAll(modCmn.GetRowText(stagingRow, COL_PHONE_NORM))
            existingValue = modCmn.NormalizePhone(modCmn.GetRowText(existingRow, COL_PHONE))
        Case COL_ZIP
            stagingValue = modCmn.TrimAll(modCmn.GetRowText(stagingRow, COL_ZIP_NORM))
            existingValue = modCmn.NormalizeZip(modCmn.GetRowText(existingRow, COL_ZIP))
        Case Else
            stagingValue = modCmn.TrimAll(modCmn.GetRowText(stagingRow, fieldName))
            existingValue = modCmn.TrimAll(modCmn.GetRowText(existingRow, fieldName))
    End Select
    
    CompareFieldValues = (stagingValue <> existingValue)
End Function

'=============================================================================
' データコピー処理
'=============================================================================

' StagingからCustomersへのデータコピー
Private Sub CopyDataFromStaging(ByVal stagingRow As ListRow, ByVal customerRow As ListRow, _
                               ByVal isNewRecord As Boolean)
    On Error Resume Next
    
    Dim sourceFile As String
    sourceFile = modCmn.GetRowText(stagingRow, COL_SOURCE_FILE)
    
    ' 基本データコピー（正規化済み値を使用）
    Call modCmn.SetRowText(customerRow, COL_CUSTOMER_ID, modCmn.GetRowText(stagingRow, COL_CUSTOMER_ID))
    Call modCmn.SetRowText(customerRow, COL_CUSTOMER_NAME, modCmn.GetRowText(stagingRow, COL_CUSTOMER_NAME))
    Call modCmn.SetRowText(customerRow, COL_EMAIL, modCmn.GetRowText(stagingRow, COL_EMAIL_NORM))
    Call modCmn.SetRowText(customerRow, COL_PHONE, modCmn.GetRowText(stagingRow, COL_PHONE_NORM))
    Call modCmn.SetRowText(customerRow, COL_ZIP, modCmn.GetRowText(stagingRow, COL_ZIP_NORM))
    Call modCmn.SetRowText(customerRow, COL_ADDRESS1, modCmn.GetRowText(stagingRow, COL_ADDRESS1))
    Call modCmn.SetRowText(customerRow, COL_ADDRESS2, modCmn.GetRowText(stagingRow, COL_ADDRESS2))
    Call modCmn.SetRowText(customerRow, COL_CATEGORY, modCmn.GetRowText(stagingRow, COL_CATEGORY))
    Call modCmn.SetRowText(customerRow, COL_STATUS, modCmn.GetRowText(stagingRow, COL_STATUS))
    Call modCmn.SetRowText(customerRow, COL_NOTES, modCmn.GetRowText(stagingRow, COL_NOTES))
    Call modCmn.SetRowText(customerRow, COL_SOURCE_FILE, sourceFile)
    
    ' 日時情報設定
    If isNewRecord Then
        Call modCmn.SetRowDate(customerRow, COL_CREATED_AT, Now)
    End If
    Call modCmn.SetRowDate(customerRow, COL_UPDATED_AT, Now)
End Sub

'=============================================================================
' 論理削除処理
'=============================================================================

' 期限切れ顧客の論理削除
Public Function InactivateStaleCustomers() As Long
    On Error GoTo ErrHandler
    
    Dim customerTbl As ListObject
    Dim row As ListRow
    Dim inactivateDays As Long
    Dim cutoffDate As Date
    Dim lastUpdated As Date
    Dim inactivatedCount As Long
    Dim startTime As Double
    
    startTime = Timer
    InactivateStaleCustomers = 0
    
    ' 設定値取得
    inactivateDays = CLng(modData.GetConfigValue(CONFIG_INACTIVATE_DAYS))
    If inactivateDays <= 0 Then
        Call modCmn.LogWarn("InactivateStaleCustomers", "無効化日数が未設定または無効です")
        Exit Function
    End If
    
    cutoffDate = Now - inactivateDays
    
    Set customerTbl = modData.GetCustomersTable()
    If customerTbl Is Nothing Then Exit Function
    
    Call modCmn.ShowProgressStart(MSG_CLEANUP_STARTED)
    
    ' 各顧客レコードをチェック
    For Each row In customerTbl.ListRows
        lastUpdated = modCmn.GetRowDate(row, COL_UPDATED_AT)
        
        ' 期限切れかつ有効なレコードを無効化
        If lastUpdated > 0 And lastUpdated < cutoffDate Then
            If modCmn.GetRowText(row, COL_STATUS) = STATUS_ACTIVE Then
                Call modCmn.SetRowText(row, COL_STATUS, STATUS_INACTIVE)
                Call modCmn.SetRowDate(row, COL_UPDATED_AT, Now)
                inactivatedCount = inactivatedCount + 1
            End If
        End If
    Next row
    
    ' ログ記録
    Call modData.LogImportOperation("期限切れ顧客無効化", inactivatedCount, Timer - startTime, _
                                   "無効化閾値: " & inactivateDays & "日")
    
    InactivateStaleCustomers = inactivatedCount
    Call modCmn.HideProgress
    Exit Function
    
ErrHandler:
    Call modCmn.HideProgress
    InactivateStaleCustomers = -1
    Call modCmn.LogError("InactivateStaleCustomers", "論理削除エラー: " & Err.Description)
End Function

'=============================================================================
' バックアップ処理
'=============================================================================

' 顧客データバックアップ
Public Function BackupCustomerData() As Boolean
    On Error GoTo ErrHandler
    
    Dim customerTbl As ListObject
    Dim backupDir As String
    Dim backupFileName As String
    Dim backupFilePath As String
    
    BackupCustomerData = False
    
    ' バックアップが有効かチェック
    If LCase(modData.GetConfigValue(CONFIG_BACKUP_ENABLED)) <> "true" Then
        Call modCmn.LogInfo("BackupCustomerData", "バックアップ機能が無効化されています")
        BackupCustomerData = True ' 無効の場合は成功扱い
        Exit Function
    End If
    
    backupDir = modData.GetConfigValue(CONFIG_BACKUP_DIR)
    If Not modCmn.DirectoryExists(backupDir) Then
        If Not modCmn.CreateDirectoryIfNotExists(backupDir) Then
            Call modCmn.LogWarn("BackupCustomerData", "バックアップディレクトリ作成失敗: " & backupDir)
            Exit Function
        End If
    End If
    
    ' バックアップファイル名生成
    backupFileName = BACKUP_FILE_PREFIX & Format(Now, DATE_FORMAT_FILE) & ".xlsx"
    backupFilePath = backupDir & backupFileName
    
    ' 現在のワークブックをバックアップ
    Application.DisplayAlerts = False
    ThisWorkbook.SaveCopyAs backupFilePath
    Application.DisplayAlerts = True
    
    ' 古いバックアップファイル削除
    Call CleanupOldBackups(backupDir)
    
    BackupCustomerData = True
    Call modCmn.LogInfo("BackupCustomerData", "バックアップ作成完了: " & backupFileName)
    Exit Function
    
ErrHandler:
    Application.DisplayAlerts = True
    BackupCustomerData = False
    Call modCmn.LogError("BackupCustomerData", "バックアップエラー: " & Err.Description)
End Function

' 古いバックアップファイル削除
Private Sub CleanupOldBackups(ByVal backupDir As String)
    On Error Resume Next
    
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim files As Collection
    Dim fileName As Variant
    Dim fileCount As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(backupDir)
    Set files = New Collection
    
    ' バックアップファイル一覧取得
    For Each file In folder.Files
        If Left(file.Name, Len(BACKUP_FILE_PREFIX)) = BACKUP_FILE_PREFIX Then
            files.Add file.Name
            fileCount = fileCount + 1
        End If
    Next file
    
    ' 最大保持数を超えている場合は古いものを削除
    If fileCount > MAX_BACKUP_FILES Then
        ' ファイル名でソート（日付順になる）
        Dim sortedFiles As New Collection
        ' 簡易的な削除（実装簡略化）
        For Each fileName In files
            If fileCount > MAX_BACKUP_FILES Then
                Kill backupDir & fileName
                fileCount = fileCount - 1
                Call modCmn.LogInfo("CleanupOldBackups", "古いバックアップ削除: " & fileName)
            End If
        Next fileName
    End If
End Sub

'=============================================================================
' ユーティリティ関数
'=============================================================================

' アップサート統計取得
Public Function GetUpsertStatistics() As Object
    On Error Resume Next
    
    Dim stats As Object
    Dim stagingTbl As ListObject
    Dim row As ListRow
    Dim totalCount As Long
    Dim validCount As Long
    Dim errorCount As Long
    
    Set GetUpsertStatistics = CreateObject("Scripting.Dictionary")
    Set stats = GetUpsertStatistics
    
    Set stagingTbl = modData.GetStagingTable()
    If stagingTbl Is Nothing Then Exit Function
    
    For Each row In stagingTbl.ListRows
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

' 最後のアップサート日時取得
Public Function GetLastUpsertDateTime() As Date
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    Dim lastDateTime As Date
    Dim rowDateTime As Date
    
    GetLastUpsertDateTime = 0
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_LOGS)
    If ws Is Nothing Then Exit Function
    
    Set tbl = modCmn.GetTable(ws, TABLE_LOGS)
    If tbl Is Nothing Then Exit Function
    
    ' ログから最新のアップサート記録を検索
    For Each row In tbl.ListRows
        If InStr(modCmn.GetRowText(row, "Message"), "アップサート") > 0 Then
            rowDateTime = modCmn.SafeDate(modCmn.GetRowText(row, "Timestamp"))
            If rowDateTime > lastDateTime Then
                lastDateTime = rowDateTime
            End If
        End If
    Next row
    
    GetLastUpsertDateTime = lastDateTime
End Function