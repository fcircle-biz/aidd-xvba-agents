Attribute VB_Name = "modLending"
Option Explicit

' =============================================================================
' 備品貸出管理システム - 貸出・返却処理ビジネスロジック
' =============================================================================

' 貸出登録メイン処理
Public Sub RegisterLending()
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    
    ' 入力値の取得と検証
    Dim inputValues As Variant
    inputValues = GetAndValidateInput()
    
    If IsEmpty(inputValues) Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    Dim itemID As Long, borrower As String, lendDate As Date, lendingDays As Long
    itemID = inputValues(0)
    borrower = inputValues(1)
    lendDate = inputValues(2)
    lendingDays = inputValues(3)
    
    ' 在庫チェック
    If Not CheckStockAvailable(itemID) Then
        MsgBox MSG_INSUFFICIENT_STOCK, vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' 貸出記録を追加
    Call AddLendingRecord(itemID, borrower, lendDate, lendingDays)
    
    ' 入力フィールドクリア
    Call ClearInputFields()
    
    ' ダッシュボード更新
    Call UpdateDashboard
    
    ' 監査ログ記録
    Call LogAudit("貸出登録", "ItemID:" & itemID & ", Borrower:" & borrower & ", Days:" & lendingDays)
    
    MsgBox "貸出を登録しました。", vbInformation
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Call LogError("RegisterLending", Err.Number, Err.Description)
    MsgBox "貸出登録中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' 返却登録メイン処理
Public Sub RegisterReturn()
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    
    ' 入力値の取得と検証
    Dim ws As Worksheet
    Set ws = GetWorksheet(SHEET_INPUT)
    If ws Is Nothing Then
        MsgBox "入力シートが見つかりません。", vbCritical
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' 必須入力チェック
    Dim itemID As Long, borrower As String, returnDate As Date
    
    If Not IsNumeric(ws.Range(INPUT_ITEM_ID).Value) Then
        MsgBox MSG_INVALID_ITEM_ID, vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If
    itemID = CLng(ws.Range(INPUT_ITEM_ID).Value)
    
    borrower = Trim(ws.Range(INPUT_BORROWER).Value)
    If borrower = "" Then
        MsgBox MSG_REQUIRED_FIELD & "(借用者)", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' 返却日（省略時は今日）
    If IsDate(ws.Range(INPUT_RETURN_DATE).Value) Then
        returnDate = CDate(ws.Range(INPUT_RETURN_DATE).Value)
    Else
        returnDate = Date
    End If
    
    ' 貸出記録検索
    Dim recordRow As Long
    recordRow = FindLendingRecord(itemID, borrower)
    
    If recordRow = 0 Then
        MsgBox MSG_NO_LENDING_RECORD, vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' 返却処理実行
    Call ProcessReturn(recordRow, returnDate)
    
    ' 入力フィールドクリア
    Call ClearInputFields()
    
    ' ダッシュボード更新
    Call UpdateDashboard
    
    ' 監査ログ記録
    Call LogAudit("返却登録", "ItemID:" & itemID & ", Borrower:" & borrower)
    
    MsgBox "返却を登録しました。", vbInformation
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Call LogError("RegisterReturn", Err.Number, Err.Description)
    MsgBox "返却登録中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' 入力値取得・検証（内部関数）
Private Function GetAndValidateInput() As Variant
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = GetWorksheet(SHEET_INPUT)
    If ws Is Nothing Then
        MsgBox "入力シートが見つかりません。", vbCritical
        GetAndValidateInput = Empty
        Exit Function
    End If
    
    ' 必須入力チェック
    Dim itemID As Long, borrower As String, lendDate As Date, lendingDays As Long
    
    If Not IsNumeric(ws.Range(INPUT_ITEM_ID).Value) Then
        MsgBox MSG_INVALID_ITEM_ID, vbExclamation
        GetAndValidateInput = Empty
        Exit Function
    End If
    itemID = CLng(ws.Range(INPUT_ITEM_ID).Value)
    
    borrower = Trim(ws.Range(INPUT_BORROWER).Value)
    If borrower = "" Then
        MsgBox MSG_REQUIRED_FIELD & "(借用者)", vbExclamation
        GetAndValidateInput = Empty
        Exit Function
    End If
    
    ' 貸出日（省略時は今日）
    If IsDate(ws.Range(INPUT_LEND_DATE).Value) Then
        lendDate = CDate(ws.Range(INPUT_LEND_DATE).Value)
    Else
        lendDate = Date
    End If
    
    ' 貸出期間（省略時はデフォルト値）
    If IsNumeric(ws.Range(INPUT_LENDING_DAYS).Value) Then
        lendingDays = CLng(ws.Range(INPUT_LENDING_DAYS).Value)
    Else
        lendingDays = DEFAULT_LENDING_DAYS
    End If
    
    ' 期間上限チェック
    If lendingDays > MAX_LENDING_DAYS Then
        MsgBox "貸出期間は" & MAX_LENDING_DAYS & "日以内で指定してください。", vbExclamation
        GetAndValidateInput = Empty
        Exit Function
    End If
    
    ' 備品存在チェック
    If Not ItemExists(itemID) Then
        MsgBox MSG_ITEM_NOT_FOUND, vbExclamation
        GetAndValidateInput = Empty
        Exit Function
    End If
    
    ' 戻り値として配列で返す
    GetAndValidateInput = Array(itemID, borrower, lendDate, lendingDays)
    
    Exit Function
    
ErrHandler:
    Call LogError("GetAndValidateInput", Err.Number, Err.Description)
    GetAndValidateInput = Empty
End Function

' 在庫可用性チェック（内部関数）
Private Function CheckStockAvailable(itemID As Long) As Boolean
    On Error GoTo ErrHandler
    
    CheckStockAvailable = False
    
    Dim availableQty As Long
    availableQty = GetAvailableQuantity(itemID)
    
    If availableQty > 0 Then
        CheckStockAvailable = True
    End If
    
    Exit Function
    
ErrHandler:
    Call LogError("CheckStockAvailable", Err.Number, Err.Description)
    CheckStockAvailable = False
End Function

' 貸出記録追加（内部関数）
Private Sub AddLendingRecord(itemID As Long, borrower As String, lendDate As Date, lendingDays As Long)
    On Error GoTo ErrHandler
    
    Dim tbl As ListObject
    Set tbl = GetLendingTable()
    
    If tbl Is Nothing Then
        Call LogError("AddLendingRecord", 9, "Lending table not found")
        Exit Sub
    End If
    
    ' 新しい行を追加
    Dim newRow As ListRow
    Set newRow = tbl.ListRows.Add()
    
    ' 各列に値を設定
    newRow.Range.Cells(1, GetColumnIndex(tbl, COL_RECORD_ID)).Value = GetNextRecordID()
    newRow.Range.Cells(1, GetColumnIndex(tbl, COL_LENDING_ITEM_ID)).Value = itemID
    newRow.Range.Cells(1, GetColumnIndex(tbl, COL_LENDING_ITEM_NAME)).Value = GetItemName(itemID)
    newRow.Range.Cells(1, GetColumnIndex(tbl, COL_BORROWER)).Value = borrower
    newRow.Range.Cells(1, GetColumnIndex(tbl, COL_LEND_DATE)).Value = lendDate
    newRow.Range.Cells(1, GetColumnIndex(tbl, COL_DUE_DATE)).Value = lendDate + lendingDays
    newRow.Range.Cells(1, GetColumnIndex(tbl, COL_RETURN_DATE)).Value = ""
    newRow.Range.Cells(1, GetColumnIndex(tbl, COL_STATUS)).Value = STATUS_LENDING
    newRow.Range.Cells(1, GetColumnIndex(tbl, COL_REMARKS)).Value = ""
    
    Exit Sub
    
ErrHandler:
    Call LogError("AddLendingRecord", Err.Number, Err.Description)
End Sub

' 返却処理実行（内部関数）
Private Sub ProcessReturn(recordRow As Long, returnDate As Date)
    On Error GoTo ErrHandler
    
    Dim tbl As ListObject
    Set tbl = GetLendingTable()
    
    If tbl Is Nothing Then
        Call LogError("ProcessReturn", 9, "Lending table not found")
        Exit Sub
    End If
    
    Dim returnDateCol As Long, statusCol As Long
    returnDateCol = GetColumnIndex(tbl, COL_RETURN_DATE)
    statusCol = GetColumnIndex(tbl, COL_STATUS)
    
    If returnDateCol = 0 Or statusCol = 0 Then
        Call LogError("ProcessReturn", 9, "Required columns not found")
        Exit Sub
    End If
    
    ' ステータス確認
    If tbl.DataBodyRange.Cells(recordRow, statusCol).Value = STATUS_RETURNED Then
        MsgBox MSG_ALREADY_RETURNED, vbExclamation
        Exit Sub
    End If
    
    ' 返却日とステータスを更新
    tbl.DataBodyRange.Cells(recordRow, returnDateCol).Value = returnDate
    tbl.DataBodyRange.Cells(recordRow, statusCol).Value = STATUS_RETURNED
    
    Exit Sub
    
ErrHandler:
    Call LogError("ProcessReturn", Err.Number, Err.Description)
End Sub

' 入力フィールドクリア（内部関数）
Private Sub ClearInputFields()
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = GetWorksheet(SHEET_INPUT)
    If Not ws Is Nothing Then
        ws.Range(INPUT_ITEM_ID).Value = ""
        ws.Range(INPUT_BORROWER).Value = ""
        ws.Range(INPUT_LEND_DATE).Value = ""
        ws.Range(INPUT_LENDING_DAYS).Value = ""
        ws.Range(INPUT_RETURN_DATE).Value = ""
    End If
    
    On Error GoTo 0
End Sub

' 期限超過件数取得
Public Function GetOverdueCount() As Long
    On Error GoTo ErrHandler
    
    GetOverdueCount = 0
    Dim tbl As ListObject
    Set tbl = GetLendingTable()
    
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    
    Dim dueDateCol As Long, statusCol As Long
    dueDateCol = GetColumnIndex(tbl, COL_DUE_DATE)
    statusCol = GetColumnIndex(tbl, COL_STATUS)
    
    If dueDateCol = 0 Or statusCol = 0 Then Exit Function
    
    Dim i As Long, count As Long
    count = 0
    For i = 1 To tbl.DataBodyRange.Rows.Count
        If tbl.DataBodyRange.Cells(i, statusCol).Value = STATUS_LENDING Then
            If IsDate(tbl.DataBodyRange.Cells(i, dueDateCol).Value) Then
                If CDate(tbl.DataBodyRange.Cells(i, dueDateCol).Value) < Date Then
                    count = count + 1
                End If
            End If
        End If
    Next i
    
    GetOverdueCount = count
    
    Exit Function
    
ErrHandler:
    Call LogError("GetOverdueCount", Err.Number, Err.Description)
    GetOverdueCount = 0
End Function

' 総貸出中件数取得
Public Function GetTotalLendingCount() As Long
    On Error GoTo ErrHandler
    
    GetTotalLendingCount = 0
    Dim tbl As ListObject
    Set tbl = GetLendingTable()
    
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    
    Dim statusCol As Long
    statusCol = GetColumnIndex(tbl, COL_STATUS)
    
    If statusCol = 0 Then Exit Function
    
    Dim i As Long, count As Long
    count = 0
    For i = 1 To tbl.DataBodyRange.Rows.Count
        If tbl.DataBodyRange.Cells(i, statusCol).Value = STATUS_LENDING Then
            count = count + 1
        End If
    Next i
    
    GetTotalLendingCount = count
    
    Exit Function
    
ErrHandler:
    Call LogError("GetTotalLendingCount", Err.Number, Err.Description)
    GetTotalLendingCount = 0
End Function