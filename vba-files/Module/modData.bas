Attribute VB_Name = "modData"
'=============================================================================
' modData.bas - データアクセス・CSV処理モジュール
'=============================================================================
' 概要:
'   CSV取り込み、設定値管理、テーブル操作、ファイルI/O等のデータアクセス層
'   外部データソースとの連携、設定値の取得・保存機能を提供
'=============================================================================
Option Explicit

'=============================================================================
' 設定値管理
'=============================================================================

' 設定値取得
Public Function GetConfigValue(ByVal configKey As String) As String
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_CONFIG)
    If ws Is Nothing Then GoTo ErrHandler
    
    Set tbl = modCmn.GetTable(ws, TABLE_CONFIG)
    If tbl Is Nothing Then GoTo ErrHandler
    
    ' 設定テーブルから指定キーの値を検索
    For Each row In tbl.ListRows
        If modCmn.GetRowText(row, "ConfigKey") = configKey Then
            GetConfigValue = modCmn.GetRowText(row, "ConfigValue")
            Exit Function
        End If
    Next row
    
    ' 設定が見つからない場合はデフォルト値を返す
    GetConfigValue = GetDefaultConfigValue(configKey)
    Call modCmn.LogWarn("GetConfigValue", "設定キー未定義、デフォルト値使用: " & configKey)
    Exit Function
    
ErrHandler:
    GetConfigValue = GetDefaultConfigValue(configKey)
    Call modCmn.LogError("GetConfigValue", "設定値取得エラー: " & configKey & " - " & Err.Description)
End Function

' デフォルト設定値取得
Public Function GetDefaultConfigValue(ByVal key As String) As String
    On Error Resume Next
    
    Select Case UCase(key)
        Case CONFIG_CSV_DIR
            GetDefaultConfigValue = DEFAULT_CSV_DIR
        Case CONFIG_CSV_FILE
            GetDefaultConfigValue = DEFAULT_CSV_FILE
        Case CONFIG_PRIMARY_KEY
            GetDefaultConfigValue = DEFAULT_PRIMARY_KEY
        Case CONFIG_ALT_KEY
            GetDefaultConfigValue = DEFAULT_ALT_KEY
        Case CONFIG_REQUIRED
            GetDefaultConfigValue = DEFAULT_REQUIRED
        Case CONFIG_INACTIVATE_DAYS
            GetDefaultConfigValue = CStr(DEFAULT_INACTIVATE_DAYS)
        Case CONFIG_EMAIL_REGEX
            GetDefaultConfigValue = REGEX_EMAIL_STRICT
        Case CONFIG_ZIP_REGEX
            GetDefaultConfigValue = REGEX_ZIP_JAPAN
        Case CONFIG_PHONE_REGEX
            GetDefaultConfigValue = REGEX_PHONE_JAPAN
        Case CONFIG_BACKUP_DIR
            GetDefaultConfigValue = DEFAULT_BACKUP_DIR
        Case CONFIG_BACKUP_ENABLED
            GetDefaultConfigValue = "True"
        Case Else
            GetDefaultConfigValue = ""
    End Select
End Function

' 設定値保存
Public Sub SetConfigValue(ByVal configKey As String, ByVal configValue As String)
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    Dim found As Boolean
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_CONFIG)
    If ws Is Nothing Then GoTo ErrHandler
    
    Set tbl = modCmn.GetTable(ws, TABLE_CONFIG)
    If tbl Is Nothing Then GoTo ErrHandler
    
    ' 既存の設定を検索
    For Each row In tbl.ListRows
        If modCmn.GetRowText(row, "ConfigKey") = configKey Then
            Call modCmn.SetRowText(row, "ConfigValue", configValue)
            found = True
            Exit For
        End If
    Next row
    
    ' 新しい設定の場合は追加
    If Not found Then
        Set row = tbl.ListRows.Add
        Call modCmn.SetRowText(row, "ConfigKey", configKey)
        Call modCmn.SetRowText(row, "ConfigValue", configValue)
    End If
    
    Call modCmn.LogInfo("SetConfigValue", "設定保存完了: " & configKey & " = " & configValue)
    Exit Sub
    
ErrHandler:
    Call modCmn.LogError("SetConfigValue", "設定値保存エラー: " & configKey & " - " & Err.Description)
End Sub

'=============================================================================
' CSV処理関数
'=============================================================================

' CSVファイル一覧取得
Public Function GetCsvFileList() As Collection
    On Error GoTo ErrHandler
    
    Dim csvDir As String
    Dim filePattern As String
    Dim fileName As String
    
    Set GetCsvFileList = New Collection
    
    csvDir = GetConfigValue(CONFIG_CSV_DIR)
    filePattern = GetConfigValue(CONFIG_CSV_FILE)
    
    ' ディレクトリ存在チェック
    If Not modCmn.DirectoryExists(csvDir) Then
        Call modCmn.LogWarn("GetCsvFileList", "CSVディレクトリが存在しません: " & csvDir)
        Exit Function
    End If
    
    ' ワイルドカードパターンをDir関数用に変換
    fileName = Dir(csvDir & filePattern)
    Do While fileName <> ""
        GetCsvFileList.Add csvDir & fileName
        fileName = Dir()
    Loop
    
    Call modCmn.LogInfo("GetCsvFileList", "CSVファイル " & GetCsvFileList.Count & " 件発見")
    Exit Function
    
ErrHandler:
    Set GetCsvFileList = New Collection
    Call modCmn.LogError("GetCsvFileList", "CSVファイル一覧取得エラー: " & Err.Description)
End Function

' CSV→Staging取り込み
Public Sub ImportCsvToStaging()
    On Error GoTo ErrHandler
    
    Dim csvFiles As Collection
    Dim filePath As Variant
    Dim totalRecords As Long
    Dim startTime As Double
    
    startTime = Timer
    Call modCmn.ShowProgressStart(MSG_IMPORT_STARTED)
    
    ' Stagingクリア
    Call ClearStagingData
    
    ' CSVファイル一覧取得
    Set csvFiles = GetCsvFileList()
    If csvFiles.Count = 0 Then
        Call modCmn.LogWarn("ImportCsvToStaging", ERR_CSV_NOT_FOUND)
        MsgBox ERR_CSV_NOT_FOUND, vbExclamation
        GoTo ExitHandler
    End If
    
    ' 各CSVファイルを処理
    For Each filePath In csvFiles
        Call modCmn.UpdateProgress("CSV処理中: " & Dir(CStr(filePath)))
        totalRecords = totalRecords + ImportSingleCsvFile(CStr(filePath))
    Next filePath
    
    ' Stagingデータ正規化
    Call NormalizeStagingData
    
    ' ログ記録
    Call LogImportOperation("CSV取り込み完了", totalRecords, Timer - startTime, "")
    
    MsgBox MSG_IMPORT_COMPLETED & vbCrLf & _
           "処理件数: " & Format(totalRecords, NUMBER_FORMAT_COUNT) & " 件" & vbCrLf & _
           "処理時間: " & Format(Timer - startTime, "0.0") & " 秒", vbInformation
    
ExitHandler:
    Call modCmn.HideProgress
    Exit Sub
    
ErrHandler:
    Call modCmn.HideProgress
    Call modCmn.LogError("ImportCsvToStaging", "CSV取り込みエラー: " & Err.Description)
End Sub

' 単一CSVファイル取り込み
Private Function ImportSingleCsvFile(ByVal filePath As String) As Long
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fileNum As Integer
    Dim lineData As String
    Dim fields As Variant
    Dim recordCount As Long
    Dim lineNumber As Long
    Dim row As ListRow
    Dim fileName As String
    
    ImportSingleCsvFile = 0
    fileName = Dir(filePath)
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_STAGING)
    If ws Is Nothing Then GoTo ErrHandler
    
    Set tbl = modCmn.GetTable(ws, TABLE_STAGING)
    If tbl Is Nothing Then GoTo ErrHandler
    
    ' ファイル読み込み開始
    fileNum = FreeFile
    Open filePath For Input As fileNum
    
    ' ヘッダー行スキップ
    Line Input #fileNum, lineData
    lineNumber = 1
    
    ' データ行処理
    Do Until EOF(fileNum)
        Line Input #fileNum, lineData
        lineNumber = lineNumber + 1
        
        If Len(Trim(lineData)) > 0 Then
            fields = ParseCsvLine(lineData)
            If IsArray(fields) And UBound(fields) >= 8 Then ' 最低必要な列数チェック
                Set row = tbl.ListRows.Add
                Call SetStagingRowFromCsv(row, fields, fileName)
                recordCount = recordCount + 1
                
                ' バッチ処理でプログレス更新
                If recordCount Mod BATCH_SIZE_CSV_IMPORT = 0 Then
                    Call modCmn.UpdateProgress("CSV処理中: " & fileName & " (" & recordCount & " 件)")
                End If
            Else
                Call modCmn.LogWarn("ImportSingleCsvFile", "不正なCSV行をスキップ: " & fileName & " 行" & lineNumber)
            End If
        End If
    Loop
    
    Close fileNum
    ImportSingleCsvFile = recordCount
    
    Call modCmn.LogInfo("ImportSingleCsvFile", fileName & " 取り込み完了: " & recordCount & " 件")
    Exit Function
    
ErrHandler:
    If fileNum > 0 Then Close fileNum
    ImportSingleCsvFile = 0
    Call modCmn.LogError("ImportSingleCsvFile", "ファイル取り込みエラー: " & filePath & " - " & Err.Description)
End Function

' CSV行パース（簡易版）
Private Function ParseCsvLine(ByVal lineData As String) As Variant
    On Error Resume Next
    
    ' カンマ区切りで分割（引用符内のカンマは考慮しない簡易版）
    ParseCsvLine = Split(lineData, CSV_DELIMITER)
End Function

' StagingテーブルにCSVデータ設定
Private Sub SetStagingRowFromCsv(ByVal row As ListRow, ByVal fields As Variant, ByVal fileName As String)
    On Error Resume Next
    
    If Not IsArray(fields) Then Exit Sub
    
    ' CSV列をStagingテーブル列にマッピング
    Call modCmn.SetRowText(row, COL_CUSTOMER_ID, GetFieldValue(fields, 0))
    Call modCmn.SetRowText(row, COL_CUSTOMER_NAME, GetFieldValue(fields, 1))
    Call modCmn.SetRowText(row, COL_EMAIL, GetFieldValue(fields, 2))
    Call modCmn.SetRowText(row, COL_PHONE, GetFieldValue(fields, 3))
    Call modCmn.SetRowText(row, COL_ZIP, GetFieldValue(fields, 4))
    Call modCmn.SetRowText(row, COL_ADDRESS1, GetFieldValue(fields, 5))
    Call modCmn.SetRowText(row, COL_ADDRESS2, GetFieldValue(fields, 6))
    Call modCmn.SetRowText(row, COL_CATEGORY, GetFieldValue(fields, 7))
    Call modCmn.SetRowText(row, COL_STATUS, GetFieldValue(fields, 8))
    Call modCmn.SetRowText(row, COL_SOURCE_FILE, fileName)
End Sub

' 配列から安全にフィールド値取得
Private Function GetFieldValue(ByVal fields As Variant, ByVal index As Integer) As String
    On Error Resume Next
    
    If IsArray(fields) And index <= UBound(fields) Then
        GetFieldValue = modCmn.TrimAll(CStr(fields(index)))
    Else
        GetFieldValue = ""
    End If
End Function

'=============================================================================
' データ正規化
'=============================================================================

' Stagingデータ正規化
Public Sub NormalizeStagingData()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    Dim recordCount As Long
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_STAGING)
    If ws Is Nothing Then GoTo ErrHandler
    
    Set tbl = modCmn.GetTable(ws, TABLE_STAGING)
    If tbl Is Nothing Then GoTo ErrHandler
    
    Call modCmn.ShowProgressStart(MSG_VALIDATION_STARTED)
    
    For Each row In tbl.ListRows
        Call NormalizeStagingRow(row)
        recordCount = recordCount + 1
        
        If recordCount Mod BATCH_SIZE_VALIDATION = 0 Then
            Call modCmn.UpdateProgress("正規化処理中: " & recordCount & " 件")
        End If
    Next row
    
    Call modCmn.LogInfo("NormalizeStagingData", "データ正規化完了: " & recordCount & " 件")
    Call modCmn.HideProgress
    Exit Sub
    
ErrHandler:
    Call modCmn.HideProgress
    Call modCmn.LogError("NormalizeStagingData", "データ正規化エラー: " & Err.Description)
End Sub

' Staging行の正規化
Private Sub NormalizeStagingRow(ByVal row As ListRow)
    On Error Resume Next
    
    Dim email As String
    Dim phone As String
    Dim zip As String
    Dim customerName As String
    Dim keyCandidate As String
    
    ' 顧客名正規化
    customerName = modCmn.TrimAll(modCmn.GetRowText(row, COL_CUSTOMER_NAME))
    Call modCmn.SetRowText(row, COL_CUSTOMER_NAME, customerName)
    
    ' メール正規化
    email = modCmn.NormalizeEmail(modCmn.GetRowText(row, COL_EMAIL))
    Call modCmn.SetRowText(row, COL_EMAIL_NORM, email)
    
    ' 電話番号正規化
    phone = modCmn.NormalizePhone(modCmn.GetRowText(row, COL_PHONE))
    Call modCmn.SetRowText(row, COL_PHONE_NORM, phone)
    
    ' 郵便番号正規化
    zip = modCmn.NormalizeZip(modCmn.GetRowText(row, COL_ZIP))
    Call modCmn.SetRowText(row, COL_ZIP_NORM, zip)
    
    ' 代替キー生成
    keyCandidate = BuildAlternateKey(row)
    Call modCmn.SetRowText(row, COL_KEY_CANDIDATE, keyCandidate)
End Sub

' 代替キー構築（Email + CustomerName）
Private Function BuildAlternateKey(ByVal row As ListRow) As String
    On Error Resume Next
    
    Dim email As String
    Dim customerName As String
    
    email = modCmn.GetRowText(row, COL_EMAIL_NORM)
    customerName = modCmn.GetRowText(row, COL_CUSTOMER_NAME)
    
    If Len(email) > 0 And Len(customerName) > 0 Then
        BuildAlternateKey = email & "+" & customerName
    ElseIf Len(customerName) > 0 Then
        BuildAlternateKey = customerName
    Else
        BuildAlternateKey = ""
    End If
End Function

'=============================================================================
' テーブル管理
'=============================================================================

' Stagingデータクリア
Public Sub ClearStagingData()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_STAGING)
    If ws Is Nothing Then GoTo ErrHandler
    
    Call modCmn.SafeClearSheet(ws, keepFormats:=True)
    Call EnsureStagingTableStructure
    
    Call modCmn.LogInfo("ClearStagingData", "Stagingデータクリア完了")
    Exit Sub
    
ErrHandler:
    Call modCmn.LogError("ClearStagingData", "Stagingクリアエラー: " & Err.Description)
End Sub

' Stagingテーブル構造確保
Private Sub EnsureStagingTableStructure()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_STAGING)
    If ws Is Nothing Then GoTo ErrHandler
    
    ' 既存テーブル削除
    If modCmn.TableExists(ws, TABLE_STAGING) Then
        ws.ListObjects(TABLE_STAGING).Delete
    End If
    
    ' ヘッダー設定
    Call SetTableHeaders(ws, "A1", STAGING_HEADERS)
    
    ' テーブル作成
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes)
    tbl.Name = TABLE_STAGING
    
    ' フォーマット適用
    Call modCmn.ApplyStandardTableFormat(tbl)
    
    Exit Sub
    
ErrHandler:
    Call modCmn.LogError("EnsureStagingTableStructure", "Stagingテーブル構造エラー: " & Err.Description)
End Sub

' テーブルヘッダー設定
Private Sub SetTableHeaders(ByVal ws As Worksheet, ByVal startCell As String, ByVal headers As String)
    On Error Resume Next
    
    Dim headerArray As Variant
    Dim i As Integer
    Dim startRange As Range
    
    headerArray = Split(headers, ",")
    Set startRange = ws.Range(startCell)
    
    For i = 0 To UBound(headerArray)
        startRange.Offset(0, i).Value = Trim(headerArray(i))
    Next i
End Sub

' 顧客テーブル取得
Public Function GetCustomersTable() As ListObject
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_CUSTOMERS)
    If ws Is Nothing Then GoTo ErrHandler
    
    Set GetCustomersTable = modCmn.GetTable(ws, TABLE_CUSTOMERS)
    Exit Function
    
ErrHandler:
    Set GetCustomersTable = Nothing
    Call modCmn.LogError("GetCustomersTable", "顧客テーブル取得エラー: " & Err.Description)
End Function

' Stagingテーブル取得
Public Function GetStagingTable() As ListObject
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_STAGING)
    If ws Is Nothing Then GoTo ErrHandler
    
    Set GetStagingTable = modCmn.GetTable(ws, TABLE_STAGING)
    Exit Function
    
ErrHandler:
    Set GetStagingTable = Nothing
    Call modCmn.LogError("GetStagingTable", "Stagingテーブル取得エラー: " & Err.Description)
End Function

'=============================================================================
' ログ記録
'=============================================================================

' インポート操作ログ記録
Public Sub LogImportOperation(ByVal message As String, ByVal recordCount As Long, _
                             ByVal processTime As Double, ByVal sourceFile As String)
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_LOGS)
    If ws Is Nothing Then Exit Sub
    
    Set tbl = modCmn.GetTable(ws, TABLE_LOGS)
    If tbl Is Nothing Then Exit Sub
    
    Set row = tbl.ListRows.Add
    Call modCmn.SetRowText(row, "Timestamp", modCmn.GetCurrentDateTimeString())
    Call modCmn.SetRowText(row, "Level", LOG_LEVEL_INFO)
    Call modCmn.SetRowText(row, "Module", "modData")
    Call modCmn.SetRowText(row, "Message", message)
    Call modCmn.SetRowText(row, "RecordCount", CStr(recordCount))
    Call modCmn.SetRowText(row, "ProcessTime", Format(processTime, "0.00") & "秒")
    Call modCmn.SetRowText(row, "SourceFile", sourceFile)
    
    Exit Sub
    
ErrHandler:
    Call modCmn.LogError("LogImportOperation", "ログ記録エラー: " & Err.Description)
End Sub

' エラーログ記録
Public Sub LogErrorOperation(ByVal message As String, ByVal details As String, ByVal sourceFile As String)
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_LOGS)
    If ws Is Nothing Then Exit Sub
    
    Set tbl = modCmn.GetTable(ws, TABLE_LOGS)
    If tbl Is Nothing Then Exit Sub
    
    Set row = tbl.ListRows.Add
    Call modCmn.SetRowText(row, "Timestamp", modCmn.GetCurrentDateTimeString())
    Call modCmn.SetRowText(row, "Level", LOG_LEVEL_ERROR)
    Call modCmn.SetRowText(row, "Module", "modData")
    Call modCmn.SetRowText(row, "Message", message)
    Call modCmn.SetRowText(row, "Details", details)
    Call modCmn.SetRowText(row, "SourceFile", sourceFile)
    
    Exit Sub
    
ErrHandler:
    ' ログエラーの場合は外部ログのみ出力
    Call modCmn.LogError("LogErrorOperation", "ログ記録エラー: " & Err.Description)
End Sub