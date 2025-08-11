Attribute VB_Name = "modValidation"
'=============================================================================
' modValidation.bas - データ検証・重複検出モジュール
'=============================================================================
' 概要:
'   Stagingデータの検証、重複検出、形式チェック、必須項目チェック等
'   検証ルールの管理、エラー情報の記録機能を提供
'=============================================================================
Option Explicit

'=============================================================================
' メイン検証処理
'=============================================================================

' Stagingデータ検証実行
Public Function ValidateStagingData() As Long
    On Error GoTo ErrHandler
    
    Dim tbl As ListObject
    Dim row As ListRow
    Dim errorCount As Long
    Dim totalCount As Long
    Dim startTime As Double
    
    startTime = Timer
    errorCount = 0
    totalCount = 0
    
    Set tbl = modData.GetStagingTable()
    If tbl Is Nothing Then
        ValidateStagingData = -1
        Exit Function
    End If
    
    Call modCmn.ShowProgressStart(MSG_VALIDATION_STARTED)
    
    ' 各行の検証
    For Each row In tbl.ListRows
        totalCount = totalCount + 1
        
        If ValidateStagingRow(row) > 0 Then
            errorCount = errorCount + 1
        End If
        
        ' プログレス更新
        If totalCount Mod BATCH_SIZE_VALIDATION = 0 Then
            Call modCmn.UpdateProgress("検証処理中: " & totalCount & " 件 (エラー: " & errorCount & " 件)")
        End If
    Next row
    
    ' 重複検出（Staging内）
    errorCount = errorCount + DetectStagingDuplicates()
    
    ' 既存顧客との重複検出
    errorCount = errorCount + DetectCustomerDuplicates()
    
    Call modData.LogImportOperation("データ検証完了", totalCount, Timer - startTime, _
                                   "エラー件数: " & errorCount)
    
    ValidateStagingData = errorCount
    Call modCmn.HideProgress
    Exit Function
    
ErrHandler:
    Call modCmn.HideProgress
    ValidateStagingData = -1
    Call modCmn.LogError("ValidateStagingData", "検証処理エラー: " & Err.Description)
End Function

' 単一行検証
Private Function ValidateStagingRow(ByVal row As ListRow) As Long
    On Error Resume Next
    
    Dim errorMessages As Collection
    Dim errorMsg As Variant
    Dim fullErrorMessage As String
    
    Set errorMessages = New Collection
    
    ' 必須フィールドチェック
    Call CheckRequiredFields(row, errorMessages)
    
    ' 形式チェック
    Call CheckDataFormats(row, errorMessages)
    
    ' ビジネスルールチェック
    Call CheckBusinessRules(row, errorMessages)
    
    ' エラーメッセージ統合
    If errorMessages.Count > 0 Then
        For Each errorMsg In errorMessages
            If Len(fullErrorMessage) > 0 Then fullErrorMessage = fullErrorMessage & "; "
            fullErrorMessage = fullErrorMessage & CStr(errorMsg)
        Next errorMsg
        
        Call modCmn.SetRowValue(row, COL_IS_VALID, False)
        Call modCmn.SetRowText(row, COL_ERROR_MESSAGE, fullErrorMessage)
        ValidateStagingRow = 1
    Else
        Call modCmn.SetRowValue(row, COL_IS_VALID, True)
        Call modCmn.SetRowText(row, COL_ERROR_MESSAGE, "")
        ValidateStagingRow = 0
    End If
End Function

'=============================================================================
' 個別検証ルール
'=============================================================================

' 必須フィールドチェック
Private Sub CheckRequiredFields(ByVal row As ListRow, ByVal errorMessages As Collection)
    On Error Resume Next
    
    Dim requiredFields As Collection
    Dim fieldName As Variant
    Dim fieldValue As String
    
    Set requiredFields = modCmn.SplitToCollection(modData.GetConfigValue(CONFIG_REQUIRED))
    
    For Each fieldName In requiredFields
        fieldValue = modCmn.TrimAll(modCmn.GetRowText(row, CStr(fieldName)))
        If Len(fieldValue) = 0 Then
            errorMessages.Add "必須フィールドが空です: " & fieldName
        End If
    Next fieldName
End Sub

' データ形式チェック
Private Sub CheckDataFormats(ByVal row As ListRow, ByVal errorMessages As Collection)
    On Error Resume Next
    
    Dim emailNorm As String
    Dim phoneNorm As String
    Dim zipNorm As String
    Dim customerId As String
    
    ' メールアドレス形式チェック
    emailNorm = modCmn.GetRowText(row, COL_EMAIL_NORM)
    If Len(emailNorm) > 0 And Not IsValidEmailFormat(emailNorm) Then
        errorMessages.Add ERR_INVALID_EMAIL_FORMAT & ": " & emailNorm
    End If
    
    ' 電話番号形式チェック
    phoneNorm = modCmn.GetRowText(row, COL_PHONE_NORM)
    If Len(phoneNorm) > 0 And Not IsValidPhoneFormat(phoneNorm) Then
        errorMessages.Add ERR_INVALID_PHONE_FORMAT & ": " & phoneNorm
    End If
    
    ' 郵便番号形式チェック
    zipNorm = modCmn.GetRowText(row, COL_ZIP_NORM)
    If Len(zipNorm) > 0 And Not IsValidZipFormat(zipNorm) Then
        errorMessages.Add ERR_INVALID_ZIP_FORMAT & ": " & zipNorm
    End If
    
    ' 顧客ID形式チェック
    customerId = modCmn.GetRowText(row, COL_CUSTOMER_ID)
    If Len(customerId) > 0 And Not IsValidCustomerIdFormat(customerId) Then
        errorMessages.Add "無効な顧客ID形式: " & customerId
    End If
End Sub

' ビジネスルールチェック
Private Sub CheckBusinessRules(ByVal row As ListRow, ByVal errorMessages As Collection)
    On Error Resume Next
    
    Dim status As String
    Dim category As String
    Dim customerName As String
    
    ' ステータス値チェック
    status = modCmn.GetRowText(row, COL_STATUS)
    If Not IsValidStatus(status) Then
        errorMessages.Add "無効なステータス値: " & status
    End If
    
    ' カテゴリ値チェック
    category = modCmn.GetRowText(row, COL_CATEGORY)
    If Len(category) > 0 And Not IsValidCategory(category) Then
        errorMessages.Add "無効なカテゴリ値: " & category
    End If
    
    ' 顧客名長さチェック
    customerName = modCmn.GetRowText(row, COL_CUSTOMER_NAME)
    If Len(customerName) > 100 Then
        errorMessages.Add "顧客名が長すぎます（100文字以内）: " & Left(customerName, 20) & "..."
    End If
End Sub

'=============================================================================
' 重複検出
'=============================================================================

' Staging内重複検出
Private Function DetectStagingDuplicates() As Long
    On Error GoTo ErrHandler
    
    Dim tbl As ListObject
    Dim keyDict As Object
    Dim row As ListRow
    Dim primaryKey As String
    Dim altKey As String
    Dim duplicateCount As Long
    
    DetectStagingDuplicates = 0
    
    Set tbl = modData.GetStagingTable()
    If tbl Is Nothing Then Exit Function
    
    Set keyDict = CreateObject("Scripting.Dictionary")
    
    ' 1回目：キー収集
    For Each row In tbl.ListRows
        primaryKey = modCmn.GetRowText(row, COL_CUSTOMER_ID)
        altKey = modCmn.GetRowText(row, COL_KEY_CANDIDATE)
        
        ' 主キーチェック
        If Len(primaryKey) > 0 Then
            If keyDict.Exists("PK:" & primaryKey) Then
                keyDict("PK:" & primaryKey) = keyDict("PK:" & primaryKey) + 1
            Else
                keyDict("PK:" & primaryKey) = 1
            End If
        End If
        
        ' 代替キーチェック
        If Len(altKey) > 0 Then
            If keyDict.Exists("AK:" & altKey) Then
                keyDict("AK:" & altKey) = keyDict("AK:" & altKey) + 1
            Else
                keyDict("AK:" & altKey) = 1
            End If
        End If
    Next row
    
    ' 2回目：重複マーク
    For Each row In tbl.ListRows
        If modCmn.GetRowValue(row, COL_IS_VALID) = True Then ' 他のエラーがない場合のみ
            primaryKey = modCmn.GetRowText(row, COL_CUSTOMER_ID)
            altKey = modCmn.GetRowText(row, COL_KEY_CANDIDATE)
            
            Dim isDuplicate As Boolean
            Dim dupMessage As String
            
            ' 主キー重複チェック
            If Len(primaryKey) > 0 And keyDict.Exists("PK:" & primaryKey) Then
                If keyDict("PK:" & primaryKey) > 1 Then
                    isDuplicate = True
                    dupMessage = "Staging内で顧客IDが重複: " & primaryKey
                End If
            End If
            
            ' 代替キー重複チェック（主キーがない場合）
            If Not isDuplicate And Len(primaryKey) = 0 And Len(altKey) > 0 Then
                If keyDict.Exists("AK:" & altKey) And keyDict("AK:" & altKey) > 1 Then
                    isDuplicate = True
                    dupMessage = "Staging内で代替キーが重複: " & altKey
                End If
            End If
            
            ' 重複マーク
            If isDuplicate Then
                Call modCmn.SetRowValue(row, COL_IS_VALID, False)
                Dim existingMsg As String
                existingMsg = modCmn.GetRowText(row, COL_ERROR_MESSAGE)
                If Len(existingMsg) > 0 Then dupMessage = existingMsg & "; " & dupMessage
                Call modCmn.SetRowText(row, COL_ERROR_MESSAGE, dupMessage)
                duplicateCount = duplicateCount + 1
            End If
        End If
    Next row
    
    DetectStagingDuplicates = duplicateCount
    
    If duplicateCount > 0 Then
        Call modCmn.LogWarn("DetectStagingDuplicates", "Staging内重複検出: " & duplicateCount & " 件")
    End If
    
    Exit Function
    
ErrHandler:
    DetectStagingDuplicates = 0
    Call modCmn.LogError("DetectStagingDuplicates", "Staging重複検出エラー: " & Err.Description)
End Function

' 既存顧客との重複検出
Private Function DetectCustomerDuplicates() As Long
    On Error GoTo ErrHandler
    
    Dim stagingTbl As ListObject
    Dim customerTbl As ListObject
    Dim stagingRow As ListRow
    Dim customerDict As Object
    Dim duplicateCount As Long
    
    DetectCustomerDuplicates = 0
    
    Set stagingTbl = modData.GetStagingTable()
    Set customerTbl = modData.GetCustomersTable()
    If stagingTbl Is Nothing Or customerTbl Is Nothing Then Exit Function
    
    ' 既存顧客のキー辞書作成
    Set customerDict = CreateCustomerKeyDictionary(customerTbl)
    
    ' Stagingデータとの重複チェック
    For Each stagingRow In stagingTbl.ListRows
        If modCmn.GetRowValue(stagingRow, COL_IS_VALID) = True Then ' 他のエラーがない場合のみ
            If IsCustomerDuplicate(stagingRow, customerDict) Then
                Call modCmn.SetRowValue(stagingRow, COL_IS_VALID, False)
                Dim existingMsg As String
                existingMsg = modCmn.GetRowText(stagingRow, COL_ERROR_MESSAGE)
                Dim dupMsg As String
                dupMsg = "既存顧客データとの重複が検出されました"
                If Len(existingMsg) > 0 Then dupMsg = existingMsg & "; " & dupMsg
                Call modCmn.SetRowText(stagingRow, COL_ERROR_MESSAGE, dupMsg)
                duplicateCount = duplicateCount + 1
            End If
        End If
    Next stagingRow
    
    DetectCustomerDuplicates = duplicateCount
    
    If duplicateCount > 0 Then
        Call modCmn.LogWarn("DetectCustomerDuplicates", "既存顧客重複検出: " & duplicateCount & " 件")
    End If
    
    Exit Function
    
ErrHandler:
    DetectCustomerDuplicates = 0
    Call modCmn.LogError("DetectCustomerDuplicates", "顧客重複検出エラー: " & Err.Description)
End Function

' 顧客キー辞書作成
Private Function CreateCustomerKeyDictionary(ByVal customerTbl As ListObject) As Object
    On Error Resume Next
    
    Dim customerDict As Object
    Dim row As ListRow
    Dim primaryKey As String
    Dim email As String
    Dim customerName As String
    Dim altKey As String
    
    Set CreateCustomerKeyDictionary = CreateObject("Scripting.Dictionary")
    Set customerDict = CreateCustomerKeyDictionary
    
    For Each row In customerTbl.ListRows
        ' 主キー（顧客ID）
        primaryKey = modCmn.GetRowText(row, COL_CUSTOMER_ID)
        If Len(primaryKey) > 0 Then
            customerDict("PK:" & primaryKey) = True
        End If
        
        ' 代替キー（Email + CustomerName）
        email = modCmn.NormalizeEmail(modCmn.GetRowText(row, COL_EMAIL))
        customerName = modCmn.GetRowText(row, COL_CUSTOMER_NAME)
        If Len(email) > 0 And Len(customerName) > 0 Then
            altKey = email & "+" & customerName
            customerDict("AK:" & altKey) = True
        End If
    Next row
End Function

' 顧客重複判定
Private Function IsCustomerDuplicate(ByVal stagingRow As ListRow, ByVal customerDict As Object) As Boolean
    On Error Resume Next
    
    Dim primaryKey As String
    Dim altKey As String
    
    IsCustomerDuplicate = False
    
    ' 主キーによる重複チェック
    primaryKey = modCmn.GetRowText(stagingRow, COL_CUSTOMER_ID)
    If Len(primaryKey) > 0 Then
        If customerDict.Exists("PK:" & primaryKey) Then
            IsCustomerDuplicate = True
            Exit Function
        End If
    End If
    
    ' 代替キーによる重複チェック（主キーがない場合）
    If Len(primaryKey) = 0 Then
        altKey = modCmn.GetRowText(stagingRow, COL_KEY_CANDIDATE)
        If Len(altKey) > 0 Then
            If customerDict.Exists("AK:" & altKey) Then
                IsCustomerDuplicate = True
            End If
        End If
    End If
End Function

'=============================================================================
' 形式検証ヘルパー関数
'=============================================================================

' メールアドレス形式検証
Private Function IsValidEmailFormat(ByVal email As String) As Boolean
    On Error Resume Next
    
    If Len(email) < 5 Or InStr(email, "@") = 0 Then
        IsValidEmailFormat = False
        Exit Function
    End If
    
    ' 基本的な正規表現チェック
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = modData.GetConfigValue(CONFIG_EMAIL_REGEX)
    regex.IgnoreCase = True
    
    IsValidEmailFormat = regex.Test(email)
End Function

' 電話番号形式検証
Private Function IsValidPhoneFormat(ByVal phone As String) As Boolean
    On Error Resume Next
    
    IsValidPhoneFormat = modCmn.IsValidJapanesePhone(phone)
End Function

' 郵便番号形式検証
Private Function IsValidZipFormat(ByVal zip As String) As Boolean
    On Error Resume Next
    
    IsValidZipFormat = modCmn.IsValidJapaneseZip(zip)
End Function

' 顧客ID形式検証
Private Function IsValidCustomerIdFormat(ByVal customerId As String) As Boolean
    On Error Resume Next
    
    IsValidCustomerIdFormat = modCmn.IsValidCustomerId(customerId)
End Function

' ステータス値検証
Private Function IsValidStatus(ByVal status As String) As Boolean
    On Error Resume Next
    
    Select Case status
        Case STATUS_ACTIVE, STATUS_INACTIVE, STATUS_SUSPENDED
            IsValidStatus = True
        Case Else
            IsValidStatus = False
    End Select
End Function

' カテゴリ値検証
Private Function IsValidCategory(ByVal category As String) As Boolean
    On Error Resume Next
    
    Select Case category
        Case CATEGORY_B2B, CATEGORY_B2C, CATEGORY_PARTNER, CATEGORY_RESELLER
            IsValidCategory = True
        Case Else
            IsValidCategory = False
    End Select
End Function

'=============================================================================
' 検証レポート生成
'=============================================================================

' 検証エラーレポート生成
Public Function GenerateValidationReport() As String
    On Error GoTo ErrHandler
    
    Dim tbl As ListObject
    Dim row As ListRow
    Dim report As String
    Dim totalCount As Long
    Dim errorCount As Long
    Dim validCount As Long
    
    Set tbl = modData.GetStagingTable()
    If tbl Is Nothing Then
        GenerateValidationReport = "Stagingテーブルが取得できませんでした。"
        Exit Function
    End If
    
    ' 統計収集
    For Each row In tbl.ListRows
        totalCount = totalCount + 1
        If modCmn.GetRowValue(row, COL_IS_VALID) = True Then
            validCount = validCount + 1
        Else
            errorCount = errorCount + 1
        End If
    Next row
    
    ' レポート生成
    report = "=== データ検証レポート ===" & vbCrLf
    report = report & "検証日時: " & modCmn.GetCurrentDateTimeString() & vbCrLf
    report = report & "総件数: " & Format(totalCount, NUMBER_FORMAT_COUNT) & " 件" & vbCrLf
    report = report & "有効件数: " & Format(validCount, NUMBER_FORMAT_COUNT) & " 件" & vbCrLf
    report = report & "エラー件数: " & Format(errorCount, NUMBER_FORMAT_COUNT) & " 件" & vbCrLf
    
    If totalCount > 0 Then
        report = report & "有効率: " & Format(validCount / totalCount * 100, "0.0") & "%" & vbCrLf
    End If
    
    GenerateValidationReport = report
    Exit Function
    
ErrHandler:
    GenerateValidationReport = "検証レポート生成エラー: " & Err.Description
    Call modCmn.LogError("GenerateValidationReport", Err.Description)
End Function

' 有効なStagingレコード数取得
Public Function GetValidStagingCount() As Long
    On Error Resume Next
    
    Dim tbl As ListObject
    Dim row As ListRow
    
    GetValidStagingCount = 0
    
    Set tbl = modData.GetStagingTable()
    If tbl Is Nothing Then Exit Function
    
    For Each row In tbl.ListRows
        If modCmn.GetRowValue(row, COL_IS_VALID) = True Then
            GetValidStagingCount = GetValidStagingCount + 1
        End If
    Next row
End Function

' エラーレコード数取得
Public Function GetErrorStagingCount() As Long
    On Error Resume Next
    
    Dim tbl As ListObject
    Dim row As ListRow
    
    GetErrorStagingCount = 0
    
    Set tbl = modData.GetStagingTable()
    If tbl Is Nothing Then Exit Function
    
    For Each row In tbl.ListRows
        If modCmn.GetRowValue(row, COL_IS_VALID) <> True Then
            GetErrorStagingCount = GetErrorStagingCount + 1
        End If
    Next row
End Function