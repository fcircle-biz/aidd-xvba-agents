Attribute VB_Name = "modDashboard"
Option Explicit

' =============================================================================
' 備品貸出管理システム - ダッシュボード更新・集計処理
' =============================================================================

' ダッシュボード全体更新メイン処理
Public Sub UpdateDashboard()
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    
    ' KPI更新
    Call UpdateKPISummary
    
    ' 貸出中一覧更新
    Call UpdateCurrentLendingList
    
    ' 在庫状況更新
    Call UpdateStockStatus
    
    ' 期限超過一覧更新
    Call UpdateOverdueList
    
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Call LogError("UpdateDashboard", Err.Number, Err.Description)
End Sub

' KPIサマリー更新
Public Sub UpdateKPISummary()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = GetWorksheet(SHEET_DASHBOARD)
    If ws Is Nothing Then
        Call LogError("UpdateKPISummary", 9, "Dashboard sheet not found")
        Exit Sub
    End If
    
    ' 総備品数
    Dim totalItems As Long
    totalItems = GetTotalItemsCount()
    ws.Range(RANGE_TOTAL_ITEMS).Value = totalItems
    
    ' 貸出中件数
    Dim lendingCount As Long
    lendingCount = GetTotalLendingCount()
    ws.Range(RANGE_LENDING_COUNT).Value = lendingCount
    
    ' 期限超過件数
    Dim overdueCount As Long
    overdueCount = GetOverdueCount()
    ws.Range(RANGE_OVERDUE_COUNT).Value = overdueCount
    
    ' 利用可能件数
    Dim availableCount As Long
    availableCount = GetTotalAvailableCount()
    ws.Range(RANGE_AVAILABLE_COUNT).Value = availableCount
    
    ' KPI値の色分け
    Call ApplyKPIColorFormatting(ws, overdueCount, lendingCount)
    
    Exit Sub
    
ErrHandler:
    Call LogError("UpdateKPISummary", Err.Number, Err.Description)
End Sub

' 貸出中一覧更新
Public Sub UpdateCurrentLendingList()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = GetWorksheet(SHEET_DASHBOARD)
    If ws Is Nothing Then Exit Sub
    
    ' 貸出中一覧を作成（A8:F20範囲）
    Dim startRange As Range
    Set startRange = ws.Range("A8")
    
    ' ヘッダー作成
    Call CreateLendingListHeader(startRange)
    
    ' データ作成
    Call PopulateLendingListData(startRange)
    
    Exit Sub
    
ErrHandler:
    Call LogError("UpdateCurrentLendingList", Err.Number, Err.Description)
End Sub

' 在庫状況更新
Public Sub UpdateStockStatus()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = GetWorksheet(SHEET_DASHBOARD)
    If ws Is Nothing Then Exit Sub
    
    ' 在庫状況一覧を作成（H8:L20範囲）
    Dim startRange As Range
    Set startRange = ws.Range("H8")
    
    ' ヘッダー作成
    Call CreateStockStatusHeader(startRange)
    
    ' データ作成
    Call PopulateStockStatusData(startRange)
    
    Exit Sub
    
ErrHandler:
    Call LogError("UpdateStockStatus", Err.Number, Err.Description)
End Sub

' 期限超過一覧更新
Public Sub UpdateOverdueList()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = GetWorksheet(SHEET_DASHBOARD)
    If ws Is Nothing Then Exit Sub
    
    ' 期限超過一覧を作成（A22:F35範囲）
    Dim startRange As Range
    Set startRange = ws.Range("A22")
    
    ' ヘッダー作成
    Call CreateOverdueListHeader(startRange)
    
    ' データ作成
    Call PopulateOverdueListData(startRange)
    
    Exit Sub
    
ErrHandler:
    Call LogError("UpdateOverdueList", Err.Number, Err.Description)
End Sub

' 総備品数取得（内部関数）
Private Function GetTotalItemsCount() As Long
    On Error GoTo ErrHandler
    
    GetTotalItemsCount = 0
    Dim tbl As ListObject
    Set tbl = GetItemsTable()
    
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    
    Dim quantityCol As Long
    quantityCol = GetColumnIndex(tbl, COL_QUANTITY)
    
    If quantityCol = 0 Then Exit Function
    
    Dim i As Long, total As Long
    total = 0
    For i = 1 To tbl.DataBodyRange.Rows.Count
        If IsNumeric(tbl.DataBodyRange.Cells(i, quantityCol).Value) Then
            total = total + tbl.DataBodyRange.Cells(i, quantityCol).Value
        End If
    Next i
    
    GetTotalItemsCount = total
    
    Exit Function
    
ErrHandler:
    Call LogError("GetTotalItemsCount", Err.Number, Err.Description)
    GetTotalItemsCount = 0
End Function

' 総利用可能件数取得（内部関数）
Private Function GetTotalAvailableCount() As Long
    On Error GoTo ErrHandler
    
    GetTotalAvailableCount = 0
    Dim tbl As ListObject
    Set tbl = GetItemsTable()
    
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    
    Dim itemIDCol As Long
    itemIDCol = GetColumnIndex(tbl, COL_ITEM_ID)
    
    If itemIDCol = 0 Then Exit Function
    
    Dim i As Long, total As Long
    total = 0
    For i = 1 To tbl.DataBodyRange.Rows.Count
        If IsNumeric(tbl.DataBodyRange.Cells(i, itemIDCol).Value) Then
            Dim itemID As Long
            itemID = tbl.DataBodyRange.Cells(i, itemIDCol).Value
            total = total + GetAvailableQuantity(itemID)
        End If
    Next i
    
    GetTotalAvailableCount = total
    
    Exit Function
    
ErrHandler:
    Call LogError("GetTotalAvailableCount", Err.Number, Err.Description)
    GetTotalAvailableCount = 0
End Function

' KPI色分けフォーマット適用（内部関数）
Private Sub ApplyKPIColorFormatting(ws As Worksheet, overdueCount As Long, lendingCount As Long)
    On Error Resume Next
    
    ' 期限超過件数の色分け
    If overdueCount > 0 Then
        With ws.Range(RANGE_OVERDUE_COUNT)
            .Interior.Color = COLOR_OVERDUE
            .Font.Color = vbWhite
            .Font.Bold = True
        End With
    Else
        With ws.Range(RANGE_OVERDUE_COUNT)
            .Interior.Color = COLOR_SUCCESS
            .Font.Color = vbWhite
        End With
    End If
    
    ' 貸出中件数の色分け
    If lendingCount > 10 Then ' 警告閾値
        With ws.Range(RANGE_LENDING_COUNT)
            .Interior.Color = COLOR_WARNING
            .Font.Color = vbBlack
        End With
    Else
        With ws.Range(RANGE_LENDING_COUNT)
            .Interior.Color = COLOR_NORMAL
            .Font.Color = vbBlack
        End With
    End If
    
    On Error GoTo 0
End Sub

' 貸出中一覧ヘッダー作成（内部関数）
Private Sub CreateLendingListHeader(startRange As Range)
    On Error Resume Next
    
    Dim headerRange As Range
    Set headerRange = startRange.Resize(1, 6)
    
    ' ヘッダー設定
    headerRange.Cells(1, 1).Value = "備品ID"
    headerRange.Cells(1, 2).Value = "備品名"
    headerRange.Cells(1, 3).Value = "借用者"
    headerRange.Cells(1, 4).Value = "貸出日"
    headerRange.Cells(1, 5).Value = "返却期限"
    headerRange.Cells(1, 6).Value = "超過日数"
    
    ' ヘッダーフォーマット
    With headerRange
        .Font.Bold = True
        .Interior.Color = COLOR_HEADER
        .Font.Color = vbWhite
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    On Error GoTo 0
End Sub

' 貸出中一覧データ作成（内部関数）
Private Sub PopulateLendingListData(startRange As Range)
    On Error GoTo ErrHandler
    
    ' データ範囲クリア
    Dim dataRange As Range
    Set dataRange = startRange.Offset(1, 0).Resize(12, 6)
    dataRange.ClearContents
    dataRange.Interior.Color = COLOR_NORMAL
    
    Dim tbl As ListObject
    Set tbl = GetLendingTable()
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    
    ' 列インデックス取得
    Dim itemIDCol As Long, itemNameCol As Long, borrowerCol As Long
    Dim lendDateCol As Long, dueDateCol As Long, statusCol As Long
    itemIDCol = GetColumnIndex(tbl, COL_LENDING_ITEM_ID)
    itemNameCol = GetColumnIndex(tbl, COL_LENDING_ITEM_NAME)
    borrowerCol = GetColumnIndex(tbl, COL_BORROWER)
    lendDateCol = GetColumnIndex(tbl, COL_LEND_DATE)
    dueDateCol = GetColumnIndex(tbl, COL_DUE_DATE)
    statusCol = GetColumnIndex(tbl, COL_STATUS)
    
    If itemIDCol = 0 Or itemNameCol = 0 Or borrowerCol = 0 Or _
       lendDateCol = 0 Or dueDateCol = 0 Or statusCol = 0 Then Exit Sub
    
    Dim i As Long, row As Long
    row = 1
    For i = 1 To tbl.DataBodyRange.Rows.Count
        If tbl.DataBodyRange.Cells(i, statusCol).Value = STATUS_LENDING And row <= 12 Then
            ' データ設定
            dataRange.Cells(row, 1).Value = tbl.DataBodyRange.Cells(i, itemIDCol).Value
            dataRange.Cells(row, 2).Value = tbl.DataBodyRange.Cells(i, itemNameCol).Value
            dataRange.Cells(row, 3).Value = tbl.DataBodyRange.Cells(i, borrowerCol).Value
            dataRange.Cells(row, 4).Value = tbl.DataBodyRange.Cells(i, lendDateCol).Value
            dataRange.Cells(row, 5).Value = tbl.DataBodyRange.Cells(i, dueDateCol).Value
            
            ' 超過日数計算と色分け
            If IsDate(tbl.DataBodyRange.Cells(i, dueDateCol).Value) Then
                Dim dueDate As Date, overdueDays As Long
                dueDate = CDate(tbl.DataBodyRange.Cells(i, dueDateCol).Value)
                overdueDays = Date - dueDate
                
                If overdueDays > 0 Then
                    dataRange.Cells(row, 6).Value = overdueDays & "日超過"
                    ' 期限超過行を赤色で強調
                    dataRange.Rows(row).Interior.Color = COLOR_OVERDUE
                    dataRange.Rows(row).Font.Color = vbWhite
                ElseIf overdueDays >= -WARNING_DAYS_BEFORE Then
                    dataRange.Cells(row, 6).Value = "期限間近"
                    ' 期限間近行を黄色で警告
                    dataRange.Rows(row).Interior.Color = COLOR_WARNING
                    dataRange.Rows(row).Font.Color = vbBlack
                Else
                    dataRange.Cells(row, 6).Value = "正常"
                End If
            End If
            
            row = row + 1
        End If
    Next i
    
    Exit Sub
    
ErrHandler:
    Call LogError("PopulateLendingListData", Err.Number, Err.Description)
End Sub

' 在庫状況ヘッダー作成（内部関数）
Private Sub CreateStockStatusHeader(startRange As Range)
    On Error Resume Next
    
    Dim headerRange As Range
    Set headerRange = startRange.Resize(1, 5)
    
    ' ヘッダー設定
    headerRange.Cells(1, 1).Value = "備品ID"
    headerRange.Cells(1, 2).Value = "備品名"
    headerRange.Cells(1, 3).Value = "総在庫"
    headerRange.Cells(1, 4).Value = "貸出中"
    headerRange.Cells(1, 5).Value = "利用可能"
    
    ' ヘッダーフォーマット
    With headerRange
        .Font.Bold = True
        .Interior.Color = COLOR_HEADER
        .Font.Color = vbWhite
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    On Error GoTo 0
End Sub

' 在庫状況データ作成（内部関数）
Private Sub PopulateStockStatusData(startRange As Range)
    On Error GoTo ErrHandler
    
    ' データ範囲クリア
    Dim dataRange As Range
    Set dataRange = startRange.Offset(1, 0).Resize(12, 5)
    dataRange.ClearContents
    dataRange.Interior.Color = COLOR_NORMAL
    
    Dim tbl As ListObject
    Set tbl = GetItemsTable()
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    
    ' 列インデックス取得
    Dim itemIDCol As Long, itemNameCol As Long, quantityCol As Long
    itemIDCol = GetColumnIndex(tbl, COL_ITEM_ID)
    itemNameCol = GetColumnIndex(tbl, COL_ITEM_NAME)
    quantityCol = GetColumnIndex(tbl, COL_QUANTITY)
    
    If itemIDCol = 0 Or itemNameCol = 0 Or quantityCol = 0 Then Exit Sub
    
    Dim i As Long, row As Long
    row = 1
    For i = 1 To tbl.DataBodyRange.Rows.Count
        If row <= 12 Then
            Dim itemID As Long, totalQty As Long, lendingQty As Long, availableQty As Long
            
            itemID = tbl.DataBodyRange.Cells(i, itemIDCol).Value
            totalQty = tbl.DataBodyRange.Cells(i, quantityCol).Value
            lendingQty = GetLendingCount(itemID)
            availableQty = totalQty - lendingQty
            
            ' データ設定
            dataRange.Cells(row, 1).Value = itemID
            dataRange.Cells(row, 2).Value = tbl.DataBodyRange.Cells(i, itemNameCol).Value
            dataRange.Cells(row, 3).Value = totalQty
            dataRange.Cells(row, 4).Value = lendingQty
            dataRange.Cells(row, 5).Value = availableQty
            
            ' 在庫切れ警告色分け
            If availableQty = 0 Then
                dataRange.Rows(row).Interior.Color = COLOR_OVERDUE
                dataRange.Rows(row).Font.Color = vbWhite
            ElseIf availableQty <= 1 Then
                dataRange.Rows(row).Interior.Color = COLOR_WARNING
                dataRange.Rows(row).Font.Color = vbBlack
            End If
            
            row = row + 1
        End If
    Next i
    
    Exit Sub
    
ErrHandler:
    Call LogError("PopulateStockStatusData", Err.Number, Err.Description)
End Sub

' 期限超過一覧ヘッダー作成（内部関数）
Private Sub CreateOverdueListHeader(startRange As Range)
    On Error Resume Next
    
    Dim headerRange As Range
    Set headerRange = startRange.Resize(1, 6)
    
    ' ヘッダー設定
    headerRange.Cells(1, 1).Value = "備品ID"
    headerRange.Cells(1, 2).Value = "備品名"
    headerRange.Cells(1, 3).Value = "借用者"
    headerRange.Cells(1, 4).Value = "貸出日"
    headerRange.Cells(1, 5).Value = "返却期限"
    headerRange.Cells(1, 6).Value = "超過日数"
    
    ' ヘッダーフォーマット（期限超過は赤色）
    With headerRange
        .Font.Bold = True
        .Interior.Color = COLOR_OVERDUE
        .Font.Color = vbWhite
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    On Error GoTo 0
End Sub

' 期限超過一覧データ作成（内部関数）
Private Sub PopulateOverdueListData(startRange As Range)
    On Error GoTo ErrHandler
    
    ' データ範囲クリア
    Dim dataRange As Range
    Set dataRange = startRange.Offset(1, 0).Resize(13, 6)
    dataRange.ClearContents
    dataRange.Interior.Color = COLOR_NORMAL
    
    Dim tbl As ListObject
    Set tbl = GetLendingTable()
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    
    ' 列インデックス取得
    Dim itemIDCol As Long, itemNameCol As Long, borrowerCol As Long
    Dim lendDateCol As Long, dueDateCol As Long, statusCol As Long
    itemIDCol = GetColumnIndex(tbl, COL_LENDING_ITEM_ID)
    itemNameCol = GetColumnIndex(tbl, COL_LENDING_ITEM_NAME)
    borrowerCol = GetColumnIndex(tbl, COL_BORROWER)
    lendDateCol = GetColumnIndex(tbl, COL_LEND_DATE)
    dueDateCol = GetColumnIndex(tbl, COL_DUE_DATE)
    statusCol = GetColumnIndex(tbl, COL_STATUS)
    
    If itemIDCol = 0 Or itemNameCol = 0 Or borrowerCol = 0 Or _
       lendDateCol = 0 Or dueDateCol = 0 Or statusCol = 0 Then Exit Sub
    
    Dim i As Long, row As Long
    row = 1
    For i = 1 To tbl.DataBodyRange.Rows.Count
        If tbl.DataBodyRange.Cells(i, statusCol).Value = STATUS_LENDING And row <= 13 Then
            If IsDate(tbl.DataBodyRange.Cells(i, dueDateCol).Value) Then
                Dim dueDate As Date, overdueDays As Long
                dueDate = CDate(tbl.DataBodyRange.Cells(i, dueDateCol).Value)
                overdueDays = Date - dueDate
                
                If overdueDays > 0 Then ' 期限超過のみ表示
                    ' データ設定
                    dataRange.Cells(row, 1).Value = tbl.DataBodyRange.Cells(i, itemIDCol).Value
                    dataRange.Cells(row, 2).Value = tbl.DataBodyRange.Cells(i, itemNameCol).Value
                    dataRange.Cells(row, 3).Value = tbl.DataBodyRange.Cells(i, borrowerCol).Value
                    dataRange.Cells(row, 4).Value = tbl.DataBodyRange.Cells(i, lendDateCol).Value
                    dataRange.Cells(row, 5).Value = dueDate
                    dataRange.Cells(row, 6).Value = overdueDays & "日超過"
                    
                    ' 期限超過行を赤色で強調
                    dataRange.Rows(row).Interior.Color = COLOR_OVERDUE
                    dataRange.Rows(row).Font.Color = vbWhite
                    
                    row = row + 1
                End If
            End If
        End If
    Next i
    
    Exit Sub
    
ErrHandler:
    Call LogError("PopulateOverdueListData", Err.Number, Err.Description)
End Sub