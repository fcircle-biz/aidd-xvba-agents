Attribute VB_Name = "modData"
Option Explicit

' =============================================================================
' 備品貸出管理システム - データアクセス層
' =============================================================================

' ワークシート取得関数
Public Function GetWorksheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

' テーブル取得関数（存在チェック付き）
Public Function GetTable(sheetName As String, tableName As String) As ListObject
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = GetWorksheet(sheetName)
    If Not ws Is Nothing Then
        Set GetTable = ws.ListObjects(tableName)
    End If
    On Error GoTo 0
End Function

' 備品テーブル取得
Public Function GetItemsTable() As ListObject
    On Error Resume Next
    Set GetItemsTable = GetTable(SHEET_ITEMS, TABLE_ITEMS)
    On Error GoTo 0
End Function

' 貸出テーブル取得
Public Function GetLendingTable() As ListObject
    On Error Resume Next
    Set GetLendingTable = GetTable(SHEET_LENDING, TABLE_LENDING)
    On Error GoTo 0
End Function

' 列インデックス取得関数
Public Function GetColumnIndex(tbl As ListObject, columnName As String) As Long
    On Error Resume Next
    Dim i As Long
    GetColumnIndex = 0
    
    If tbl Is Nothing Then Exit Function
    If tbl.HeaderRowRange Is Nothing Then Exit Function
    
    For i = 1 To tbl.ListColumns.Count
        If tbl.HeaderRowRange.Cells(1, i).Value = columnName Then
            GetColumnIndex = i
            Exit Function
        End If
    Next i
    On Error GoTo 0
End Function

' エラーログ出力関数
Public Sub LogError(procedureName As String, errorNumber As Long, errorDescription As String)
    On Error Resume Next
    
    ' Debug.Print出力（開発時）
    Debug.Print "Error in " & procedureName & ": " & errorNumber & " - " & errorDescription
    
    ' 外部ログファイル出力（本番運用）
    Dim logPath As String, fileNum As Integer
    logPath = ThisWorkbook.Path & "\" & Replace(ThisWorkbook.Name, ".xlsm", "_error.log")
    fileNum = FreeFile
    
    Open logPath For Append As fileNum
    Print #fileNum, Format(Now, "yyyy-mm-dd hh:mm:ss") & " | ERROR | " & procedureName & " | " & errorNumber & " | " & errorDescription
    Close fileNum
End Sub

' 監査ログ出力関数
Public Sub LogAudit(actionName As String, details As String, Optional userName As String = "")
    On Error Resume Next
    
    If userName = "" Then userName = Application.UserName
    
    ' Debug.Print出力（開発時）
    Debug.Print "Audit: " & actionName & " by " & userName & " - " & details
    
    ' 外部ログファイル出力（本番運用）
    Dim logPath As String, fileNum As Integer
    logPath = ThisWorkbook.Path & "\" & Replace(ThisWorkbook.Name, ".xlsm", "_audit.log")
    fileNum = FreeFile
    
    Open logPath For Append As fileNum
    Print #fileNum, Format(Now, "yyyy-mm-dd hh:mm:ss") & " | AUDIT | " & actionName & " | " & userName & " | " & details
    Close fileNum
End Sub

' 備品ID存在チェック
Public Function ItemExists(itemID As Long) As Boolean
    On Error GoTo ErrHandler
    
    ItemExists = False
    Dim tbl As ListObject
    Set tbl = GetItemsTable()
    
    If tbl Is Nothing Then
        Call LogError("ItemExists", 9, "Items table not found")
        Exit Function
    End If
    
    If tbl.DataBodyRange Is Nothing Then Exit Function
    
    Dim itemIDCol As Long
    itemIDCol = GetColumnIndex(tbl, COL_ITEM_ID)
    
    If itemIDCol = 0 Then
        Call LogError("ItemExists", 9, "Column not found: " & COL_ITEM_ID)
        Exit Function
    End If
    
    Dim i As Long
    For i = 1 To tbl.DataBodyRange.Rows.Count
        If tbl.DataBodyRange.Cells(i, itemIDCol).Value = itemID Then
            ItemExists = True
            Exit Function
        End If
    Next i
    
    Exit Function
    
ErrHandler:
    Call LogError("ItemExists", Err.Number, Err.Description)
    ItemExists = False
End Function

' 備品名取得
Public Function GetItemName(itemID As Long) As String
    On Error GoTo ErrHandler
    
    GetItemName = ""
    Dim tbl As ListObject
    Set tbl = GetItemsTable()
    
    If tbl Is Nothing Then
        Call LogError("GetItemName", 9, "Items table not found")
        Exit Function
    End If
    
    If tbl.DataBodyRange Is Nothing Then Exit Function
    
    Dim itemIDCol As Long, itemNameCol As Long
    itemIDCol = GetColumnIndex(tbl, COL_ITEM_ID)
    itemNameCol = GetColumnIndex(tbl, COL_ITEM_NAME)
    
    If itemIDCol = 0 Or itemNameCol = 0 Then
        Call LogError("GetItemName", 9, "Required columns not found")
        Exit Function
    End If
    
    Dim i As Long
    For i = 1 To tbl.DataBodyRange.Rows.Count
        If tbl.DataBodyRange.Cells(i, itemIDCol).Value = itemID Then
            GetItemName = tbl.DataBodyRange.Cells(i, itemNameCol).Value
            Exit Function
        End If
    Next i
    
    Exit Function
    
ErrHandler:
    Call LogError("GetItemName", Err.Number, Err.Description)
    GetItemName = ""
End Function

' 貸出中件数取得
Public Function GetLendingCount(itemID As Long) As Long
    On Error GoTo ErrHandler
    
    GetLendingCount = 0
    Dim tbl As ListObject
    Set tbl = GetLendingTable()
    
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    
    Dim itemIDCol As Long, statusCol As Long
    itemIDCol = GetColumnIndex(tbl, COL_LENDING_ITEM_ID)
    statusCol = GetColumnIndex(tbl, COL_STATUS)
    
    If itemIDCol = 0 Or statusCol = 0 Then Exit Function
    
    Dim i As Long, count As Long
    count = 0
    For i = 1 To tbl.DataBodyRange.Rows.Count
        If tbl.DataBodyRange.Cells(i, itemIDCol).Value = itemID And _
           tbl.DataBodyRange.Cells(i, statusCol).Value = STATUS_LENDING Then
            count = count + 1
        End If
    Next i
    
    GetLendingCount = count
    
    Exit Function
    
ErrHandler:
    Call LogError("GetLendingCount", Err.Number, Err.Description)
    GetLendingCount = 0
End Function

' 総在庫数取得
Public Function GetTotalQuantity(itemID As Long) As Long
    On Error GoTo ErrHandler
    
    GetTotalQuantity = 0
    Dim tbl As ListObject
    Set tbl = GetItemsTable()
    
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    
    Dim itemIDCol As Long, quantityCol As Long
    itemIDCol = GetColumnIndex(tbl, COL_ITEM_ID)
    quantityCol = GetColumnIndex(tbl, COL_QUANTITY)
    
    If itemIDCol = 0 Or quantityCol = 0 Then Exit Function
    
    Dim i As Long
    For i = 1 To tbl.DataBodyRange.Rows.Count
        If tbl.DataBodyRange.Cells(i, itemIDCol).Value = itemID Then
            GetTotalQuantity = tbl.DataBodyRange.Cells(i, quantityCol).Value
            Exit Function
        End If
    Next i
    
    Exit Function
    
ErrHandler:
    Call LogError("GetTotalQuantity", Err.Number, Err.Description)
    GetTotalQuantity = 0
End Function

' 利用可能在庫数取得
Public Function GetAvailableQuantity(itemID As Long) As Long
    On Error GoTo ErrHandler
    
    Dim totalQty As Long, lendingQty As Long
    totalQty = GetTotalQuantity(itemID)
    lendingQty = GetLendingCount(itemID)
    
    GetAvailableQuantity = totalQty - lendingQty
    If GetAvailableQuantity < 0 Then GetAvailableQuantity = 0
    
    Exit Function
    
ErrHandler:
    Call LogError("GetAvailableQuantity", Err.Number, Err.Description)
    GetAvailableQuantity = 0
End Function

' 次のレコードID取得
Public Function GetNextRecordID() As Long
    On Error GoTo ErrHandler
    
    GetNextRecordID = 1
    Dim tbl As ListObject
    Set tbl = GetLendingTable()
    
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    
    Dim recordIDCol As Long
    recordIDCol = GetColumnIndex(tbl, COL_RECORD_ID)
    
    If recordIDCol = 0 Then Exit Function
    
    Dim i As Long, maxID As Long
    maxID = 0
    For i = 1 To tbl.DataBodyRange.Rows.Count
        If IsNumeric(tbl.DataBodyRange.Cells(i, recordIDCol).Value) Then
            If tbl.DataBodyRange.Cells(i, recordIDCol).Value > maxID Then
                maxID = tbl.DataBodyRange.Cells(i, recordIDCol).Value
            End If
        End If
    Next i
    
    GetNextRecordID = maxID + 1
    
    Exit Function
    
ErrHandler:
    Call LogError("GetNextRecordID", Err.Number, Err.Description)
    GetNextRecordID = 1
End Function

' 貸出記録検索
Public Function FindLendingRecord(itemID As Long, borrower As String) As Long
    On Error GoTo ErrHandler
    
    FindLendingRecord = 0
    Dim tbl As ListObject
    Set tbl = GetLendingTable()
    
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    
    Dim itemIDCol As Long, borrowerCol As Long, statusCol As Long
    itemIDCol = GetColumnIndex(tbl, COL_LENDING_ITEM_ID)
    borrowerCol = GetColumnIndex(tbl, COL_BORROWER)
    statusCol = GetColumnIndex(tbl, COL_STATUS)
    
    If itemIDCol = 0 Or borrowerCol = 0 Or statusCol = 0 Then Exit Function
    
    Dim i As Long
    For i = 1 To tbl.DataBodyRange.Rows.Count
        If tbl.DataBodyRange.Cells(i, itemIDCol).Value = itemID And _
           tbl.DataBodyRange.Cells(i, borrowerCol).Value = borrower And _
           tbl.DataBodyRange.Cells(i, statusCol).Value = STATUS_LENDING Then
            FindLendingRecord = i
            Exit Function
        End If
    Next i
    
    Exit Function
    
ErrHandler:
    Call LogError("FindLendingRecord", Err.Number, Err.Description)
    FindLendingRecord = 0
End Function