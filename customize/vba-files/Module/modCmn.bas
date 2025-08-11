Attribute VB_Name = "modCmn"
'=============================================================================
' modCmn.bas - 共通ユーティリティ・汎用機能
'=============================================================================
' 概要:
'   業務システム共通で使用する汎用機能を集約
'   データアクセス、文字列処理、ログ、フォーマット、検証など再利用可能な機能群
'=============================================================================
Option Explicit

'=============================================================================
' 独立動作のための定数定義
'=============================================================================
' エラーメッセージ定数
Private Const ERR_SHEET_NOT_FOUND As String = "シートが見つかりません: "
Private Const ERR_TABLE_NOT_FOUND As String = "テーブルが見つかりません: "
Private Const ERR_COLUMN_NOT_FOUND As String = "列が見つかりません: "

' フォント設定定数
Private Const FONT_NAME As String = "Yu Gothic UI"
Private Const FONT_SIZE_NORMAL As Integer = 10
Private Const FONT_SIZE_HEADER As Integer = 12
Private Const FONT_SIZE_BUTTON As Integer = 11
Private Const FONT_COLOR_NORMAL As Long = 0
Private Const FONT_COLOR_HEADER As Long = 16777215

' 色設定定数
Private Const BG_COLOR_HEADER As Long = 5287936
Private Const BG_COLOR_ALTERNATE As Long = 15921906
Private Const BORDER_COLOR_DEFAULT As Long = 8421504

' 正規表現パターン定数
Private Const REGEX_EMAIL As String = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
Private Const REGEX_PHONE As String = "^[0-9]{2,4}-[0-9]{3,4}-[0-9]{4}$"
Private Const REGEX_ZIP As String = "^\d{3}-\d{4}$"
Private Const REGEX_CUSTOMERID As String = "^[A-Za-z0-9]{3,20}$"

' デフォルト設定値定数
Private Const DEFAULT_LOG_DIR As String = "C:\git\xvba-mock-creator\logs\"

'=============================================================================
' データアクセス・操作系関数
'=============================================================================

' 安全なワークシート取得（シート名）
Public Function GetWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheet = ThisWorkbook.Worksheets(sheetName)
    If Err.Number <> 0 Then
        Call LogError("GetWorksheet", ERR_SHEET_NOT_FOUND & sheetName)
        Set GetWorksheet = Nothing
    End If
End Function

' 安全なワークシート取得（インデックス番号）
Public Function GetWorksheetByIndex(ByVal sheetIndex As Integer) As Worksheet
    On Error GoTo ErrHandler
    
    If sheetIndex <= 0 Or sheetIndex > ThisWorkbook.Worksheets.Count Then
        Call LogError("GetWorksheetByIndex", "無効なシートインデックス: " & sheetIndex)
        Set GetWorksheetByIndex = Nothing
        Exit Function
    End If
    
    Set GetWorksheetByIndex = ThisWorkbook.Worksheets(sheetIndex)
    Exit Function
    
ErrHandler:
    Call LogError("GetWorksheetByIndex", "シートインデックス " & sheetIndex & " でエラー: " & Err.Description)
    Set GetWorksheetByIndex = Nothing
End Function

' 安全なテーブル取得
Public Function GetTable(ByVal ws As Worksheet, ByVal tableName As String) As ListObject
    On Error Resume Next
    Set GetTable = ws.ListObjects(tableName)
    If Err.Number <> 0 Then
        Call LogError("GetTable", ERR_TABLE_NOT_FOUND & tableName)
        Set GetTable = Nothing
    End If
End Function

' テーブル存在チェック
Public Function TableExists(ByVal ws As Worksheet, ByVal tableName As String) As Boolean
    On Error Resume Next
    TableExists = Not (ws.ListObjects(tableName) Is Nothing)
    If Err.Number <> 0 Then TableExists = False
End Function

' 列インデックス取得（0チェック必須）
Public Function GetColumnIndex(ByVal tbl As ListObject, ByVal columnName As String) As Integer
    On Error GoTo ErrHandler
    
    Dim col As ListColumn
    For Each col In tbl.ListColumns
        If col.Name = columnName Then
            GetColumnIndex = col.Index
            Exit Function
        End If
    Next col
    
    GetColumnIndex = 0
    Call LogError("GetColumnIndex", ERR_COLUMN_NOT_FOUND & columnName)
    Exit Function
    
ErrHandler:
    GetColumnIndex = 0
    Call LogError("GetColumnIndex", Err.Description & " (Column: " & columnName & ")")
End Function

' 安全なシートクリア
Public Sub SafeClearSheet(ByVal ws As Worksheet, Optional ByVal keepFormats As Boolean = False)
    On Error GoTo ErrHandler
    
    If ws Is Nothing Then
        Call LogError("SafeClearSheet", "ワークシートがNullです")
        Exit Sub
    End If
    
    Dim isProt As Boolean
    isProt = ws.ProtectContents
    If isProt Then ws.Unprotect
    
    ' 使用範囲を特定
    Dim found As Range
    Set found = ws.Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If Not found Is Nothing Then
        With ws.Range("A1", ws.Cells(found.Row, found.Column))
            If keepFormats Then
                .ClearContents
            Else
                .Clear
            End If
        End With
    End If
    
    If isProt Then ws.Protect UserInterfaceOnly:=True
    Exit Sub
    
ErrHandler:
    If isProt Then ws.Protect UserInterfaceOnly:=True
    Call LogError("SafeClearSheet", Err.Description)
End Sub

'=============================================================================
' データ取得・設定関数（テーブル行操作）
'=============================================================================

' テーブル行から値を取得
Public Function GetRowValue(ByVal row As ListRow, ByVal columnName As String) As Variant
    On Error GoTo ErrHandler
    
    Dim colIndex As Integer
    colIndex = GetColumnIndex(row.Parent, columnName)
    If colIndex = 0 Then
        GetRowValue = Empty
        Exit Function
    End If
    
    GetRowValue = row.Range.Cells(1, colIndex).Value
    Exit Function
    
ErrHandler:
    GetRowValue = Empty
    Call LogError("GetRowValue", Err.Description & " (Column: " & columnName & ")")
End Function

' テーブル行に値を設定
Public Sub SetRowValue(ByVal row As ListRow, ByVal columnName As String, ByVal value As Variant)
    On Error GoTo ErrHandler
    
    Dim colIndex As Integer
    colIndex = GetColumnIndex(row.Parent, columnName)
    If colIndex = 0 Then Exit Sub
    
    row.Range.Cells(1, colIndex).Value = value
    Exit Sub
    
ErrHandler:
    Call LogError("SetRowValue", Err.Description & " (Column: " & columnName & ")")
End Sub

' テーブル行から文字列取得
Public Function GetRowText(ByVal row As ListRow, ByVal columnName As String) As String
    On Error Resume Next
    GetRowText = CStr(GetRowValue(row, columnName))
    If Err.Number <> 0 Then GetRowText = ""
End Function

' テーブル行に文字列設定
Public Sub SetRowText(ByVal row As ListRow, ByVal columnName As String, ByVal text As String)
    Call SetRowValue(row, columnName, text)
End Sub

' テーブル行から日付取得
Public Function GetRowDate(ByVal row As ListRow, ByVal columnName As String) As Date
    On Error Resume Next
    GetRowDate = CDate(GetRowValue(row, columnName))
    If Err.Number <> 0 Then GetRowDate = 0
End Function

' テーブル行に日付設定
Public Sub SetRowDate(ByVal row As ListRow, ByVal columnName As String, ByVal dateValue As Date)
    Call SetRowValue(row, columnName, dateValue)
End Sub

'=============================================================================
' 文字列処理ユーティリティ
'=============================================================================

' 文字列トリム（全角スペースも対応）
Public Function TrimAll(ByVal text As String) As String
    On Error Resume Next
    ' 前後の半角・全角スペース、タブ、改行を除去
    TrimAll = text
    TrimAll = Replace(TrimAll, vbTab, " ")
    TrimAll = Replace(TrimAll, vbCrLf, " ")
    TrimAll = Replace(TrimAll, vbCr, " ")
    TrimAll = Replace(TrimAll, vbLf, " ")
    TrimAll = Replace(TrimAll, "　", " ")  ' 全角スペース→半角スペース
    
    ' 連続スペースを単一化
    Do While InStr(TrimAll, "  ") > 0
        TrimAll = Replace(TrimAll, "  ", " ")
    Loop
    
    TrimAll = Trim(TrimAll)
End Function

' 電話番号正規化
Public Function NormalizePhone(ByVal phone As String) As String
    On Error Resume Next
    
    NormalizePhone = TrimAll(phone)
    ' 全角数字・ハイフンを半角に変換
    NormalizePhone = StrConv(NormalizePhone, vbNarrow)
    
    ' 不要文字削除
    NormalizePhone = Replace(NormalizePhone, "(", "")
    NormalizePhone = Replace(NormalizePhone, ")", "")
    NormalizePhone = Replace(NormalizePhone, " ", "")
    NormalizePhone = Replace(NormalizePhone, "　", "")
    
    ' ハイフンの正規化（03-1234-5678形式）
    If Len(NormalizePhone) = 10 Or Len(NormalizePhone) = 11 Then
        ' 一度ハイフンを全て削除
        NormalizePhone = Replace(NormalizePhone, "-", "")
        
        ' 適切な位置にハイフンを挿入
        If Len(NormalizePhone) = 10 Then
            ' 03-XXXX-XXXX または 06-XXXX-XXXX
            If Left(NormalizePhone, 2) = "03" Or Left(NormalizePhone, 2) = "06" Then
                NormalizePhone = Left(NormalizePhone, 2) & "-" & Mid(NormalizePhone, 3, 4) & "-" & Right(NormalizePhone, 4)
            Else
                ' 0XX-XXX-XXXX
                NormalizePhone = Left(NormalizePhone, 3) & "-" & Mid(NormalizePhone, 4, 3) & "-" & Right(NormalizePhone, 4)
            End If
        ElseIf Len(NormalizePhone) = 11 Then
            ' 090-XXXX-XXXX
            NormalizePhone = Left(NormalizePhone, 3) & "-" & Mid(NormalizePhone, 4, 4) & "-" & Right(NormalizePhone, 4)
        End If
    End If
End Function

' 郵便番号正規化
Public Function NormalizeZip(ByVal zip As String) As String
    On Error Resume Next
    
    NormalizeZip = TrimAll(zip)
    ' 全角数字・ハイフンを半角に変換
    NormalizeZip = StrConv(NormalizeZip, vbNarrow)
    
    ' 不要文字削除
    NormalizeZip = Replace(NormalizeZip, " ", "")
    NormalizeZip = Replace(NormalizeZip, "　", "")
    NormalizeZip = Replace(NormalizeZip, "〒", "")
    
    ' ハイフンがない場合は追加（1234567 → 123-4567）
    If Len(NormalizeZip) = 7 And InStr(NormalizeZip, "-") = 0 Then
        NormalizeZip = Left(NormalizeZip, 3) & "-" & Right(NormalizeZip, 4)
    End If
End Function

' メールアドレス正規化
Public Function NormalizeEmail(ByVal email As String) As String
    On Error Resume Next
    
    NormalizeEmail = TrimAll(email)
    ' 小文字に統一
    NormalizeEmail = LCase(NormalizeEmail)
    
    ' 全角英数字を半角に変換
    NormalizeEmail = StrConv(NormalizeEmail, vbNarrow)
End Function

' 文字列をコレクションに分割
Public Function SplitToCollection(ByVal str As String, Optional ByVal delimiter As String = ",") As Collection
    On Error Resume Next
    
    Set SplitToCollection = New Collection
    
    If Len(str) = 0 Then Exit Function
    
    Dim parts As Variant
    parts = Split(str, delimiter)
    
    Dim i As Integer
    For i = LBound(parts) To UBound(parts)
        Dim part As String
        part = TrimAll(CStr(parts(i)))
        If part <> "" Then
            SplitToCollection.Add part
        End If
    Next i
End Function

'=============================================================================
' ログ管理関数
'=============================================================================

' 外部ログファイル出力
Public Sub LogError(ByVal functionName As String, ByVal errorMessage As String)
    On Error Resume Next
    
    Dim logMessage As String
    logMessage = Format(Now, "yyyy-mm-dd hh:nn:ss") & " [ERROR] " & functionName & ": " & errorMessage
    
    ' デバッグ出力（開発時のみ）
    Debug.Print logMessage
    
    ' 外部ログファイルに記録（エラー処理なし）
    Call WriteExternalLogSafe(logMessage)
    
    ' エラーダイアログを表示して強制終了
    MsgBox "システムエラーが発生しました。" & vbCrLf & vbCrLf & _
           "関数: " & functionName & vbCrLf & _
           "エラー: " & errorMessage & vbCrLf & vbCrLf & _
           "詳細はログファイルをご確認ください。" & vbCrLf & _
           "アプリケーションを終了します。", vbCritical, "システムエラー"
    
    ' 強制終了（保存確認なし）
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    ThisWorkbook.Saved = True  ' 保存済みとしてマーク
    Application.Quit
    
    ' 上記で終了しない場合の最終手段
    End
End Sub

' 情報ログ記録
Public Sub LogInfo(ByVal functionName As String, ByVal message As String)
    On Error Resume Next
    
    Dim logMessage As String
    logMessage = Format(Now, "yyyy-mm-dd hh:nn:ss") & " [INFO] " & functionName & ": " & message
    
    
    ' デバッグ出力（開発時のみ）
    Debug.Print logMessage
End Sub

' 警告ログ記録
Public Sub LogWarn(ByVal functionName As String, ByVal message As String)
    On Error Resume Next
    
    Dim logMessage As String
    logMessage = Format(Now, "yyyy-mm-dd hh:nn:ss") & " [WARN] " & functionName & ": " & message
    
    
    ' デバッグ出力（開発時のみ）
    Debug.Print logMessage
End Sub


' 外部ログファイル出力（安全版）
Private Sub WriteExternalLogSafe(ByVal logMessage As String)
    On Error Resume Next
    
    Dim logDir As String
    Dim logFilePath As String
    Dim fileNum As Integer
    
    logDir = DEFAULT_LOG_DIR
    
    ' ディレクトリが存在しない場合は作成（エラー無視）
    If Dir(logDir, vbDirectory) = "" Then
        MkDir logDir
    End If
    
    ' 日付別ログファイル
    logFilePath = logDir & "system_" & Format(Now, "yyyymmdd") & ".log"
    
    fileNum = FreeFile
    Open logFilePath For Append As fileNum
    Print #fileNum, logMessage
    Close fileNum
End Sub

' 外部ログファイル出力（従来版）
Private Sub WriteExternalLog(ByVal logMessage As String)
    On Error Resume Next
    Call WriteExternalLogSafe(logMessage)
End Sub

'=============================================================================
' フォント・色設定関数
'=============================================================================

' システム標準フォント適用
Public Sub ApplySystemFont(ByVal targetRange As Range)
    On Error GoTo ErrHandler
    
    If targetRange Is Nothing Then Exit Sub
    
    With targetRange.Font
        .Name = FONT_NAME
        .Size = FONT_SIZE_NORMAL
        .Color = FONT_COLOR_NORMAL
        .Bold = False
    End With
    
    Exit Sub
    
ErrHandler:
    Call LogError("ApplySystemFont", Err.Description)
End Sub

' ヘッダーフォント適用
Public Sub ApplyHeaderFont(ByVal targetRange As Range)
    On Error GoTo ErrHandler
    
    If targetRange Is Nothing Then Exit Sub
    
    With targetRange
        With .Font
            .Name = FONT_NAME
            .Size = FONT_SIZE_HEADER
            .Color = FONT_COLOR_HEADER
            .Bold = True
        End With
        .Interior.Color = BG_COLOR_HEADER
    End With
    
    Exit Sub
    
ErrHandler:
    Call LogError("ApplyHeaderFont", Err.Description)
End Sub

' ボタンフォント適用
Public Sub ApplyButtonFont(ByVal targetRange As Range)
    On Error GoTo ErrHandler
    
    If targetRange Is Nothing Then Exit Sub
    
    With targetRange.Font
        .Name = FONT_BUTTON
        .Size = FONT_SIZE_BUTTON
        .Color = FONT_COLOR_NORMAL
        .Bold = True
    End With
    
    Exit Sub
    
ErrHandler:
    Call LogError("ApplyButtonFont", Err.Description)
End Sub

' シート全体フォント統一
Public Sub ApplySheetFont(ByVal ws As Worksheet)
    On Error GoTo ErrHandler
    
    If ws Is Nothing Then Exit Sub
    
    ' シート全体に標準フォントを適用
    With ws.Cells.Font
        .Name = FONT_NAME
        .Size = FONT_SIZE_NORMAL
        .Color = FONT_COLOR_NORMAL
    End With
    
    Exit Sub
    
ErrHandler:
    Call LogError("ApplySheetFont", Err.Description)
End Sub

' テーブル標準フォーマット適用
Public Sub ApplyStandardTableFormat(ByVal tbl As ListObject)
    On Error GoTo ErrHandler
    
    If tbl Is Nothing Then Exit Sub
    
    ' ヘッダー行フォーマット
    Call ApplyHeaderFont(tbl.HeaderRowRange)
    
    ' データ行フォーマット
    If Not tbl.DataBodyRange Is Nothing Then
        Call ApplySystemFont(tbl.DataBodyRange)
        
        ' 交互行の背景色設定（ゼブラ縞）
        Dim i As Long
        For i = 1 To tbl.ListRows.Count Step 2
            tbl.ListRows(i).Range.Interior.Color = BG_COLOR_ALTERNATE
        Next i
    End If
    
    ' 枠線設定
    With tbl.Range.Borders
        .LineStyle = xlContinuous
        .Color = BORDER_COLOR_DEFAULT
        .Weight = xlThin
    End With
    
    Exit Sub
    
ErrHandler:
    Call LogError("ApplyStandardTableFormat", Err.Description)
End Sub

'=============================================================================
' 進捗表示管理
'=============================================================================

' 進捗表示開始
Public Sub ShowProgressStart(ByVal message As String)
    On Error Resume Next
    Application.StatusBar = message
    Application.ScreenUpdating = False
End Sub

' 進捗更新
Public Sub UpdateProgress(ByVal message As String)
    On Error Resume Next
    Application.StatusBar = message
End Sub

' 進捗表示終了
Public Sub HideProgress()
    On Error Resume Next
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub

'=============================================================================
' 検証ヘルパー関数
'=============================================================================

' 空値チェック
Public Function IsEmpty(ByVal value As Variant) As Boolean
    IsEmpty = (VarType(value) = vbEmpty Or VarType(value) = vbNull Or Len(Trim(CStr(value))) = 0)
End Function

' 日付安全変換
Public Function SafeDate(ByVal value As Variant) As Date
    On Error Resume Next
    SafeDate = CDate(value)
    If Err.Number <> 0 Then SafeDate = 0
End Function

' 数値安全変換
Public Function SafeLong(ByVal value As Variant) As Long
    On Error Resume Next
    SafeLong = CLng(value)
    If Err.Number <> 0 Then SafeLong = 0
End Function

'=============================================================================
' 正規表現検証関数
'=============================================================================

' メールアドレス形式検証
Public Function IsValidEmail(ByVal email As String) As Boolean
    On Error Resume Next
    
    If Len(email) < 5 Or InStr(email, "@") = 0 Then
        IsValidEmail = False
        Exit Function
    End If
    
    ' 正規表現パターンチェック
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = REGEX_EMAIL
    regex.IgnoreCase = True
    
    IsValidEmail = regex.Test(email)
End Function

' 電話番号形式検証
Public Function IsValidPhone(ByVal phone As String) As Boolean
    On Error Resume Next
    
    If Len(phone) < 10 Or Len(phone) > 15 Then
        IsValidPhone = False
        Exit Function
    End If
    
    ' 正規表現パターンチェック
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = REGEX_PHONE
    
    IsValidPhone = regex.Test(phone)
End Function

' 郵便番号形式検証
Public Function IsValidZip(ByVal zip As String) As Boolean
    On Error Resume Next
    
    If Len(zip) <> 8 And Len(zip) <> 7 Then  ' 123-4567 または 1234567
        IsValidZip = False
        Exit Function
    End If
    
    ' 正規表現パターンチェック
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = REGEX_ZIP
    
    IsValidZip = regex.Test(zip)
End Function

' 顧客ID形式検証
Public Function IsValidCustomerId(ByVal customerId As String) As Boolean
    On Error Resume Next
    
    If Len(customerId) < 3 Or Len(customerId) > 20 Then
        IsValidCustomerId = False
        Exit Function
    End If
    
    ' 正規表現パターンチェック（英数字のみ）
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = REGEX_CUSTOMERID
    regex.IgnoreCase = True
    
    IsValidCustomerId = regex.Test(customerId)
End Function

'=============================================================================
' 追加のヘルパー関数（設計仕様対応）
'=============================================================================

' 安全な数値変換
Public Function SafeInteger(ByVal value As Variant) As Integer
    On Error Resume Next
    SafeInteger = CInt(value)
    If Err.Number <> 0 Then SafeInteger = 0
End Function

' 安全なDouble変換
Public Function SafeDouble(ByVal value As Variant) As Double
    On Error Resume Next
    SafeDouble = CDbl(value)
    If Err.Number <> 0 Then SafeDouble = 0
End Function

' 空白または空文字チェック
Public Function IsNullOrEmpty(ByVal value As Variant) As Boolean
    IsNullOrEmpty = (IsNull(value) Or VarType(value) = vbEmpty Or Len(Trim(CStr(value & ""))) = 0)
End Function

' 文字列の最大長制限
Public Function LimitLength(ByVal text As String, ByVal maxLength As Integer) As String
    If Len(text) > maxLength Then
        LimitLength = Left(text, maxLength)
    Else
        LimitLength = text
    End If
End Function

' 日本の郵便番号パターン詳細検証
Public Function IsValidJapaneseZip(ByVal zip As String) As Boolean
    On Error Resume Next
    
    IsValidJapaneseZip = False
    
    If Len(zip) <> 8 And Len(zip) <> 7 Then Exit Function
    
    ' ハイフンありの場合 (123-4567)
    If Len(zip) = 8 Then
        If Mid(zip, 4, 1) <> "-" Then Exit Function
        If Not IsNumeric(Left(zip, 3)) Then Exit Function
        If Not IsNumeric(Right(zip, 4)) Then Exit Function
        IsValidJapaneseZip = True
    ' ハイフンなしの場合 (1234567)
    ElseIf Len(zip) = 7 Then
        If IsNumeric(zip) Then IsValidJapaneseZip = True
    End If
End Function

' 日本の電話番号パターン検証
Public Function IsValidJapanesePhone(ByVal phone As String) As Boolean
    On Error Resume Next
    
    IsValidJapanesePhone = False
    
    ' 基本長さチェック
    If Len(phone) < 10 Or Len(phone) > 15 Then Exit Function
    
    ' パターンチェック（固定電話、携帯電話）
    If Left(phone, 1) = "0" Then
        ' 固定電話: 0X-XXXX-XXXX または 0XX-XXX-XXXX
        ' 携帯電話: 090-XXXX-XXXX, 080-XXXX-XXXX, 070-XXXX-XXXX
        Dim parts As Variant
        parts = Split(phone, "-")
        
        If UBound(parts) = 2 Then ' 3つの部分に分かれている
            If Left(phone, 3) = "090" Or Left(phone, 3) = "080" Or Left(phone, 3) = "070" Then
                ' 携帯電話パターン
                If Len(parts(0)) = 3 And Len(parts(1)) = 4 And Len(parts(2)) = 4 Then
                    IsValidJapanesePhone = True
                End If
            ElseIf Left(phone, 2) = "03" Or Left(phone, 2) = "06" Then
                ' 主要都市固定電話パターン
                If Len(parts(0)) = 2 And Len(parts(1)) = 4 And Len(parts(2)) = 4 Then
                    IsValidJapanesePhone = True
                End If
            Else
                ' その他固定電話パターン
                If Len(parts(0)) = 3 And Len(parts(1)) = 3 And Len(parts(2)) = 4 Then
                    IsValidJapanesePhone = True
                End If
            End If
        End If
    End If
End Function

' CSV行のフィールド数検証
Public Function ValidateCsvFieldCount(ByVal fields As Variant, ByVal expectedCount As Integer) As Boolean
    On Error Resume Next
    
    If IsArray(fields) Then
        ValidateCsvFieldCount = (UBound(fields) + 1 = expectedCount)
    Else
        ValidateCsvFieldCount = False
    End If
End Function

' 設定値のデフォルト取得
Public Function GetDefaultConfigValue(ByVal key As String) As String
    On Error Resume Next
    
    Select Case UCase(key)
        Case "CSV_DIR"
            GetDefaultConfigValue = "C:\Data\Import\"
        Case "CSV_FILE"
            GetDefaultConfigValue = "customers_*.csv"
        Case "PRIMARY_KEY"
            GetDefaultConfigValue = "CustomerID"
        Case "ALT_KEY"
            GetDefaultConfigValue = "Email+CustomerName"
        Case "REQUIRED"
            GetDefaultConfigValue = "CustomerID,CustomerName,Status"
        Case "INACTIVATE_DAYS"
            GetDefaultConfigValue = "180"
        Case "EMAIL_REGEX"
            GetDefaultConfigValue = REGEX_EMAIL
        Case "ZIP_REGEX"
            GetDefaultConfigValue = REGEX_ZIP
        Case "PHONE_REGEX"
            GetDefaultConfigValue = REGEX_PHONE
        Case "CUSTOMERID_REGEX"
            GetDefaultConfigValue = REGEX_CUSTOMERID
        Case Else
            GetDefaultConfigValue = ""
    End Select
End Function

' コレクションから配列への変換
Public Function CollectionToArray(ByVal col As Collection) As Variant
    On Error Resume Next
    
    If col.Count = 0 Then
        CollectionToArray = Array()
        Exit Function
    End If
    
    Dim arr() As Variant
    ReDim arr(col.Count - 1)
    
    Dim i As Integer
    For i = 1 To col.Count
        arr(i - 1) = col(i)
    Next i
    
    CollectionToArray = arr
End Function

' 配列からコレクションへの変換
Public Function ArrayToCollection(ByVal arr As Variant) As Collection
    On Error Resume Next
    
    Set ArrayToCollection = New Collection
    
    If Not IsArray(arr) Then Exit Function
    
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        ArrayToCollection.Add arr(i)
    Next i
End Function

' ファイル存在チェック
Public Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(filePath) <> "")
End Function

' ディレクトリ存在チェック
Public Function DirectoryExists(ByVal dirPath As String) As Boolean
    On Error Resume Next
    DirectoryExists = (Dir(dirPath, vbDirectory) <> "")
End Function

' 安全なディレクトリ作成
Public Function CreateDirectoryIfNotExists(ByVal dirPath As String) As Boolean
    On Error Resume Next
    
    CreateDirectoryIfNotExists = False
    
    If DirectoryExists(dirPath) Then
        CreateDirectoryIfNotExists = True
        Exit Function
    End If
    
    MkDir dirPath
    If Err.Number = 0 Then
        CreateDirectoryIfNotExists = True
    End If
End Function

' 現在の日時を文字列として取得
Public Function GetCurrentDateTimeString() As String
    GetCurrentDateTimeString = Format(Now, "yyyy-mm-dd hh:nn:ss")
End Function

' 現在の日付を文字列として取得
Public Function GetCurrentDateString() As String
    GetCurrentDateString = Format(Date, "yyyy-mm-dd")
End Function

' ミリ秒を含む現在時刻取得
Public Function GetCurrentTimeStamp() As String
    GetCurrentTimeStamp = Format(Now, "yyyy-mm-dd hh:nn:ss") & "." & Format(Timer Mod 1 * 1000, "000")
End Function