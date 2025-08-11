Attribute VB_Name = "modUtils"
'=============================================================================
' modUtils.bas - ユーティリティ・ヘルパー関数モジュール
'=============================================================================
' 概要:
'   システム全体で使用されるユーティリティ関数群
'   ファイル操作、テーブル初期化、データ変換、システム管理等の補助機能
'=============================================================================
Option Explicit

'=============================================================================
' システム初期化ユーティリティ
'=============================================================================

' システム全体初期化
Public Sub InitializeCustomerSystem()
    On Error GoTo ErrHandler
    
    Call modCmn.LogInfo("InitializeCustomerSystem", "顧客管理システム初期化開始")
    
    ' 各シートの初期化
    Call InitializeAllSheets()
    
    ' テーブル構造確認・作成
    Call EnsureAllTableStructures()
    
    ' デフォルト設定値設定
    Call SetupDefaultConfiguration()
    
    ' サンプルデータ作成（初回のみ）
    Call CreateInitialSampleData()
    
    ' ダッシュボード初期化
    Call modDashboard.InitializeDashboard()
    
    Call modCmn.LogInfo("InitializeCustomerSystem", "顧客管理システム初期化完了")
    
    MsgBox SYSTEM_NAME & " の初期化が完了しました。" & vbCrLf & _
           "バージョン: " & SYSTEM_VERSION, vbInformation, "初期化完了"
    
    Exit Sub
    
ErrHandler:
    Call modCmn.LogError("InitializeCustomerSystem", "システム初期化エラー: " & Err.Description)
End Sub

' 全シート初期化
Private Sub InitializeAllSheets()
    On Error Resume Next
    
    ' シート名変更
    Call RenameSystemSheets()
    
    ' 各シートのフォント統一
    Dim i As Integer
    For i = 1 To 6
        Dim ws As Worksheet
        Set ws = modCmn.GetWorksheetByIndex(i)
        If Not ws Is Nothing Then
            Call modCmn.ApplySheetFont(ws)
        End If
    Next i
End Sub

' システムシート名変更
Private Sub RenameSystemSheets()
    On Error Resume Next
    
    Dim ws As Worksheet
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_DASHBOARD)
    If Not ws Is Nothing Then ws.Name = SHEET_DASHBOARD
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_CUSTOMERS)
    If Not ws Is Nothing Then ws.Name = SHEET_CUSTOMERS
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_STAGING)
    If Not ws Is Nothing Then ws.Name = SHEET_STAGING
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_CONFIG)
    If Not ws Is Nothing Then ws.Name = SHEET_CONFIG
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_LOGS)
    If Not ws Is Nothing Then ws.Name = SHEET_LOGS
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_CODEBOOK)
    If Not ws Is Nothing Then ws.Name = SHEET_CODEBOOK
End Sub

'=============================================================================
' テーブル構造管理
'=============================================================================

' 全テーブル構造確認・作成
Public Sub EnsureAllTableStructures()
    On Error Resume Next
    
    Call EnsureCustomersTableStructure()
    Call EnsureStagingTableStructure()
    Call EnsureConfigTableStructure()
    Call EnsureLogsTableStructure()
    Call EnsureCodebookTableStructure()
End Sub

' Customersテーブル構造確認
Private Sub EnsureCustomersTableStructure()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_CUSTOMERS)
    If ws Is Nothing Then Exit Sub
    
    ' 既存テーブルチェック
    If Not modCmn.TableExists(ws, TABLE_CUSTOMERS) Then
        ' テーブル新規作成
        Call SetTableHeaders(ws, "A1", CUSTOMERS_HEADERS)
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes)
        tbl.Name = TABLE_CUSTOMERS
        
        ' フォーマット適用
        Call modCmn.ApplyStandardTableFormat(tbl)
        
        Call modCmn.LogInfo("EnsureCustomersTableStructure", "Customersテーブル作成完了")
    End If
    
    Exit Sub
    
ErrHandler:
    Call modCmn.LogError("EnsureCustomersTableStructure", "Customersテーブル構造エラー: " & Err.Description)
End Sub

' Stagingテーブル構造確認
Private Sub EnsureStagingTableStructure()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_STAGING)
    If ws Is Nothing Then Exit Sub
    
    ' 既存テーブル削除（Stagingは都度再作成）
    If modCmn.TableExists(ws, TABLE_STAGING) Then
        ws.ListObjects(TABLE_STAGING).Delete
    End If
    
    ' シートクリア
    Call modCmn.SafeClearSheet(ws, keepFormats:=False)
    
    ' テーブル新規作成
    Call SetTableHeaders(ws, "A1", STAGING_HEADERS)
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes)
    tbl.Name = TABLE_STAGING
    
    ' フォーマット適用
    Call modCmn.ApplyStandardTableFormat(tbl)
    
    Call modCmn.LogInfo("EnsureStagingTableStructure", "Stagingテーブル作成完了")
    Exit Sub
    
ErrHandler:
    Call modCmn.LogError("EnsureStagingTableStructure", "Stagingテーブル構造エラー: " & Err.Description)
End Sub

' Configテーブル構造確認
Private Sub EnsureConfigTableStructure()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_CONFIG)
    If ws Is Nothing Then Exit Sub
    
    ' 既存テーブルチェック
    If Not modCmn.TableExists(ws, TABLE_CONFIG) Then
        Call SetTableHeaders(ws, "A1", CONFIG_HEADERS)
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes)
        tbl.Name = TABLE_CONFIG
        
        ' フォーマット適用
        Call modCmn.ApplyStandardTableFormat(tbl)
        
        Call modCmn.LogInfo("EnsureConfigTableStructure", "Configテーブル作成完了")
    End If
    
    Exit Sub
    
ErrHandler:
    Call modCmn.LogError("EnsureConfigTableStructure", "Configテーブル構造エラー: " & Err.Description)
End Sub

' Logsテーブル構造確認
Private Sub EnsureLogsTableStructure()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_LOGS)
    If ws Is Nothing Then Exit Sub
    
    ' 既存テーブルチェック
    If Not modCmn.TableExists(ws, TABLE_LOGS) Then
        Call SetTableHeaders(ws, "A1", LOGS_HEADERS)
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes)
        tbl.Name = TABLE_LOGS
        
        ' フォーマット適用
        Call modCmn.ApplyStandardTableFormat(tbl)
        
        Call modCmn.LogInfo("EnsureLogsTableStructure", "Logsテーブル作成完了")
    End If
    
    Exit Sub
    
ErrHandler:
    Call modCmn.LogError("EnsureLogsTableStructure", "Logsテーブル構造エラー: " & Err.Description)
End Sub

' Codebookテーブル構造確認
Private Sub EnsureCodebookTableStructure()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_CODEBOOK)
    If ws Is Nothing Then Exit Sub
    
    ' 既存テーブルチェック
    If Not modCmn.TableExists(ws, TABLE_CODEBOOK) Then
        Call SetTableHeaders(ws, "A1", CODEBOOK_HEADERS)
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes)
        tbl.Name = TABLE_CODEBOOK
        
        ' フォーマット適用
        Call modCmn.ApplyStandardTableFormat(tbl)
        
        ' デフォルトマッピング設定
        Call SetupDefaultColumnMappings(tbl)
        
        Call modCmn.LogInfo("EnsureCodebookTableStructure", "Codebookテーブル作成完了")
    End If
    
    Exit Sub
    
ErrHandler:
    Call modCmn.LogError("EnsureCodebookTableStructure", "Codebookテーブル構造エラー: " & Err.Description)
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
        With startRange.Offset(0, i)
            .Value = Trim(headerArray(i))
            Call modCmn.ApplyHeaderFont(.Cells)
        End With
    Next i
End Sub

'=============================================================================
' デフォルトデータ設定
'=============================================================================

' デフォルト設定値設定
Public Sub SetupDefaultConfiguration()
    On Error Resume Next
    
    ' 基本設定値
    Call modData.SetConfigValue(CONFIG_CSV_DIR, DEFAULT_CSV_DIR)
    Call modData.SetConfigValue(CONFIG_CSV_FILE, DEFAULT_CSV_FILE)
    Call modData.SetConfigValue(CONFIG_PRIMARY_KEY, DEFAULT_PRIMARY_KEY)
    Call modData.SetConfigValue(CONFIG_ALT_KEY, DEFAULT_ALT_KEY)
    Call modData.SetConfigValue(CONFIG_REQUIRED, DEFAULT_REQUIRED)
    Call modData.SetConfigValue(CONFIG_INACTIVATE_DAYS, CStr(DEFAULT_INACTIVATE_DAYS))
    Call modData.SetConfigValue(CONFIG_EMAIL_REGEX, REGEX_EMAIL_STRICT)
    Call modData.SetConfigValue(CONFIG_ZIP_REGEX, REGEX_ZIP_JAPAN)
    Call modData.SetConfigValue(CONFIG_PHONE_REGEX, REGEX_PHONE_JAPAN)
    Call modData.SetConfigValue(CONFIG_BACKUP_ENABLED, "True")
    Call modData.SetConfigValue(CONFIG_BACKUP_DIR, DEFAULT_BACKUP_DIR)
    
    Call modCmn.LogInfo("SetupDefaultConfiguration", "デフォルト設定値設定完了")
End Sub

' デフォルト列マッピング設定
Private Sub SetupDefaultColumnMappings(ByVal codebookTbl As ListObject)
    On Error Resume Next
    
    Dim mappings As Variant
    Dim i As Integer
    
    ' CSVヘッダーと内部列のマッピング定義
    mappings = Array( _
        Array("customer_id", COL_CUSTOMER_ID, "文字列", "", "", "True", "顧客識別ID"), _
        Array("customer_name", COL_CUSTOMER_NAME, "文字列", "", "TrimAll", "True", "顧客名"), _
        Array("email", COL_EMAIL, "文字列", REGEX_EMAIL_STRICT, "NormalizeEmail", "False", "メールアドレス"), _
        Array("phone", COL_PHONE, "文字列", REGEX_PHONE_JAPAN, "NormalizePhone", "False", "電話番号"), _
        Array("zip", COL_ZIP, "文字列", REGEX_ZIP_JAPAN, "NormalizeZip", "False", "郵便番号"), _
        Array("address1", COL_ADDRESS1, "文字列", "", "TrimAll", "False", "住所1"), _
        Array("address2", COL_ADDRESS2, "文字列", "", "TrimAll", "False", "住所2"), _
        Array("category", COL_CATEGORY, "文字列", "", "", "False", "顧客カテゴリ"), _
        Array("status", COL_STATUS, "文字列", "", "", "True", "顧客ステータス") _
    )
    
    ' マッピングデータ追加
    For i = 0 To UBound(mappings)
        Dim row As ListRow
        Set row = codebookTbl.ListRows.Add
        
        Call modCmn.SetRowText(row, "ExternalColumnName", CStr(mappings(i)(0)))
        Call modCmn.SetRowText(row, "InternalColumnName", CStr(mappings(i)(1)))
        Call modCmn.SetRowText(row, "DataType", CStr(mappings(i)(2)))
        Call modCmn.SetRowText(row, "ValidationRule", CStr(mappings(i)(3)))
        Call modCmn.SetRowText(row, "NormalizationRule", CStr(mappings(i)(4)))
        Call modCmn.SetRowText(row, "Required", CStr(mappings(i)(5)))
        Call modCmn.SetRowText(row, "Description", CStr(mappings(i)(6)))
    Next i
End Sub

' 初期サンプルデータ作成
Public Sub CreateInitialSampleData()
    On Error Resume Next
    
    Dim customerTbl As ListObject
    Set customerTbl = modData.GetCustomersTable()
    
    ' 既にデータがある場合はスキップ
    If Not customerTbl Is Nothing Then
        If customerTbl.ListRows.Count > 0 Then Exit Sub
    End If
    
    ' サンプル顧客データ
    Dim sampleData As Variant
    sampleData = Array( _
        Array("SAMPLE001", "サンプル株式会社", "sample@example.com", "03-1234-5678", "100-0001", "東京都千代田区", "千代田1-1-1", CATEGORY_B2B, STATUS_ACTIVE, "システム初期データ"), _
        Array("SAMPLE002", "テスト商事", "test@business.co.jp", "06-9876-5432", "530-0001", "大阪府大阪市北区", "梅田2-2-2", CATEGORY_PARTNER, STATUS_ACTIVE, "システム初期データ") _
    )
    
    ' サンプルデータ追加
    Dim i As Integer
    For i = 0 To UBound(sampleData)
        Dim row As ListRow
        Set row = customerTbl.ListRows.Add
        
        Call modCmn.SetRowText(row, COL_CUSTOMER_ID, CStr(sampleData(i)(0)))
        Call modCmn.SetRowText(row, COL_CUSTOMER_NAME, CStr(sampleData(i)(1)))
        Call modCmn.SetRowText(row, COL_EMAIL, CStr(sampleData(i)(2)))
        Call modCmn.SetRowText(row, COL_PHONE, CStr(sampleData(i)(3)))
        Call modCmn.SetRowText(row, COL_ZIP, CStr(sampleData(i)(4)))
        Call modCmn.SetRowText(row, COL_ADDRESS1, CStr(sampleData(i)(5)))
        Call modCmn.SetRowText(row, COL_ADDRESS2, CStr(sampleData(i)(6)))
        Call modCmn.SetRowText(row, COL_CATEGORY, CStr(sampleData(i)(7)))
        Call modCmn.SetRowText(row, COL_STATUS, CStr(sampleData(i)(8)))
        Call modCmn.SetRowDate(row, COL_CREATED_AT, Now)
        Call modCmn.SetRowDate(row, COL_UPDATED_AT, Now)
        Call modCmn.SetRowText(row, COL_SOURCE_FILE, "初期データ")
        Call modCmn.SetRowText(row, COL_NOTES, CStr(sampleData(i)(9)))
    Next i
    
    Call modCmn.LogInfo("CreateInitialSampleData", "サンプルデータ作成完了")
End Sub

'=============================================================================
' データ変換ユーティリティ
'=============================================================================

' 文字列を安全に配列に変換
Public Function SafeSplitString(ByVal inputString As String, ByVal delimiter As String) As Variant
    On Error Resume Next
    
    If Len(inputString) = 0 Then
        SafeSplitString = Array()
    Else
        SafeSplitString = Split(inputString, delimiter)
    End If
End Function

' 配列を安全に文字列に変換
Public Function SafeJoinArray(ByVal inputArray As Variant, ByVal delimiter As String) As String
    On Error Resume Next
    
    If IsArray(inputArray) Then
        SafeJoinArray = Join(inputArray, delimiter)
    Else
        SafeJoinArray = ""
    End If
End Function

' CSV行安全パース
Public Function SafeParseCsvLine(ByVal csvLine As String) As Variant
    On Error Resume Next
    
    Dim fields As Variant
    Dim i As Integer
    
    ' 基本的なCSV分割（引用符処理なし）
    fields = Split(csvLine, CSV_DELIMITER)
    
    ' 各フィールドをトリム
    For i = 0 To UBound(fields)
        fields(i) = modCmn.TrimAll(CStr(fields(i)))
        ' 引用符除去
        If Left(fields(i), 1) = CSV_QUOTE_CHAR And Right(fields(i), 1) = CSV_QUOTE_CHAR Then
            If Len(fields(i)) > 1 Then
                fields(i) = Mid(fields(i), 2, Len(fields(i)) - 2)
            Else
                fields(i) = ""
            End If
        End If
    Next i
    
    SafeParseCsvLine = fields
End Function

'=============================================================================
' パフォーマンス最適化ユーティリティ
'=============================================================================

' パフォーマンス最適化開始
Public Sub StartPerformanceOptimization()
    On Error Resume Next
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
End Sub

' パフォーマンス最適化終了
Public Sub EndPerformanceOptimization()
    On Error Resume Next
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub

' メモリクリーンアップ
Public Sub CleanupMemory()
    On Error Resume Next
    
    ' ガベージコレクション実行（可能な場合）
    DoEvents
    
    ' 一時変数クリア
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "Temp*" Then ws.Delete
    Next ws
    
    Call modCmn.LogInfo("CleanupMemory", "メモリクリーンアップ実行")
End Sub

'=============================================================================
' システム診断ユーティリティ
'=============================================================================

' システム健全性チェック
Public Function PerformSystemHealthCheck() As Object
    On Error Resume Next
    
    Dim result As Object
    Set PerformSystemHealthCheck = CreateObject("Scripting.Dictionary")
    Set result = PerformSystemHealthCheck
    
    result("OverallHealth") = "Healthy"
    result("Issues") = New Collection
    result("Warnings") = New Collection
    
    ' シート存在チェック
    Dim i As Integer
    For i = 1 To 6
        If modCmn.GetWorksheetByIndex(i) Is Nothing Then
            result("Issues").Add "必要なシートが見つかりません: インデックス " & i
            result("OverallHealth") = "Critical"
        End If
    Next i
    
    ' テーブル存在チェック
    If modData.GetCustomersTable() Is Nothing Then
        result("Issues").Add "Customersテーブルが見つかりません"
        result("OverallHealth") = "Critical"
    End If
    
    ' 設定値チェック
    If Len(modData.GetConfigValue(CONFIG_CSV_DIR)) = 0 Then
        result("Warnings").Add "CSVディレクトリが設定されていません"
        If result("OverallHealth") = "Healthy" Then result("OverallHealth") = "Warning"
    End If
    
    ' ディスク容量チェック（簡易）
    If modCmn.DirectoryExists(modData.GetConfigValue(CONFIG_CSV_DIR)) = False Then
        result("Warnings").Add "CSVディレクトリにアクセスできません"
        If result("OverallHealth") = "Healthy" Then result("OverallHealth") = "Warning"
    End If
End Function

' システム情報取得
Public Function GetSystemInformation() As Object
    On Error Resume Next
    
    Dim info As Object
    Set GetSystemInformation = CreateObject("Scripting.Dictionary")
    Set info = GetSystemInformation
    
    info("SystemName") = SYSTEM_NAME
    info("Version") = SYSTEM_VERSION
    info("Author") = SYSTEM_AUTHOR
    info("CurrentTime") = modCmn.GetCurrentDateTimeString()
    info("WorkbookPath") = ThisWorkbook.FullName
    info("WorkbookName") = ThisWorkbook.Name
    info("ExcelVersion") = Application.Version
    info("SheetCount") = ThisWorkbook.Worksheets.Count
    
    ' テーブル統計
    Dim customerTbl As ListObject
    Set customerTbl = modData.GetCustomersTable()
    If Not customerTbl Is Nothing Then
        info("CustomerCount") = customerTbl.ListRows.Count
    Else
        info("CustomerCount") = 0
    End If
    
    Dim stagingTbl As ListObject
    Set stagingTbl = modData.GetStagingTable()
    If Not stagingTbl Is Nothing Then
        info("StagingCount") = stagingTbl.ListRows.Count
    Else
        info("StagingCount") = 0
    End If
End Function

'=============================================================================
' デバッグ・開発支援ユーティリティ
'=============================================================================

' システム状態ダンプ（デバッグ用）
Public Sub DumpSystemState()
    On Error Resume Next
    
    Debug.Print "=== システム状態ダンプ ==="
    Debug.Print "時刻: " & modCmn.GetCurrentDateTimeString()
    
    ' システム情報
    Dim info As Object
    Set info = GetSystemInformation()
    
    Dim key As Variant
    For Each key In info.Keys
        Debug.Print key & ": " & info(key)
    Next key
    
    ' 健全性チェック
    Dim health As Object
    Set health = PerformSystemHealthCheck()
    Debug.Print "システム健全性: " & health("OverallHealth")
    
    Debug.Print "=== ダンプ終了 ==="
End Sub

' テストデータ生成（開発用）
Public Sub GenerateTestData()
    On Error Resume Next
    
    Dim customerTbl As ListObject
    Set customerTbl = modData.GetCustomersTable()
    If customerTbl Is Nothing Then Exit Sub
    
    ' テストデータ100件生成
    Dim i As Integer
    For i = 1 To 100
        Dim row As ListRow
        Set row = customerTbl.ListRows.Add
        
        Call modCmn.SetRowText(row, COL_CUSTOMER_ID, "TEST" & Format(i, "000"))
        Call modCmn.SetRowText(row, COL_CUSTOMER_NAME, "テスト顧客" & i)
        Call modCmn.SetRowText(row, COL_EMAIL, "test" & i & "@example.com")
        Call modCmn.SetRowText(row, COL_PHONE, "03-" & Format(i, "0000") & "-0000")
        Call modCmn.SetRowText(row, COL_ZIP, "100-000" & (i Mod 10))
        Call modCmn.SetRowText(row, COL_ADDRESS1, "東京都千代田区")
        Call modCmn.SetRowText(row, COL_ADDRESS2, "テスト" & i & "-1-1")
        Call modCmn.SetRowText(row, COL_CATEGORY, IIf(i Mod 2 = 0, CATEGORY_B2B, CATEGORY_B2C))
        Call modCmn.SetRowText(row, COL_STATUS, IIf(i Mod 10 = 0, STATUS_INACTIVE, STATUS_ACTIVE))
        Call modCmn.SetRowDate(row, COL_CREATED_AT, Now - i)
        Call modCmn.SetRowDate(row, COL_UPDATED_AT, Now)
        Call modCmn.SetRowText(row, COL_SOURCE_FILE, "テストデータ")
        Call modCmn.SetRowText(row, COL_NOTES, "自動生成テストデータ")
    Next i
    
    Call modCmn.LogInfo("GenerateTestData", "テストデータ100件生成完了")
End Sub