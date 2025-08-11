Attribute VB_Name = "modDashboard"
Option Explicit

' =============================================================================
' ���i�ݏo�Ǘ��V�X�e�� - �_�b�V���{�[�h�X�V�E�W�v����
' =============================================================================

' �_�b�V���{�[�h�S�̍X�V���C������
Public Sub UpdateDashboard()
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    
    ' KPI�X�V
    Call UpdateKPISummary
    
    ' �ݏo���ꗗ�X�V
    Call UpdateCurrentLendingList
    
    ' �݌ɏ󋵍X�V
    Call UpdateStockStatus
    
    ' �������߈ꗗ�X�V
    Call UpdateOverdueList
    
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Call LogError("UpdateDashboard", Err.Number, Err.Description)
End Sub

' KPI�T�}���[�X�V
Public Sub UpdateKPISummary()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = GetWorksheet(SHEET_DASHBOARD)
    If ws Is Nothing Then
        Call LogError("UpdateKPISummary", 9, "Dashboard sheet not found")
        Exit Sub
    End If
    
    ' �����i��
    Dim totalItems As Long
    totalItems = GetTotalItemsCount()
    ws.Range(RANGE_TOTAL_ITEMS).Value = totalItems
    
    ' �ݏo������
    Dim lendingCount As Long
    lendingCount = GetTotalLendingCount()
    ws.Range(RANGE_LENDING_COUNT).Value = lendingCount
    
    ' �������ߌ���
    Dim overdueCount As Long
    overdueCount = GetOverdueCount()
    ws.Range(RANGE_OVERDUE_COUNT).Value = overdueCount
    
    ' ���p�\����
    Dim availableCount As Long
    availableCount = GetTotalAvailableCount()
    ws.Range(RANGE_AVAILABLE_COUNT).Value = availableCount
    
    ' KPI�l�̐F����
    Call ApplyKPIColorFormatting(ws, overdueCount, lendingCount)
    
    Exit Sub
    
ErrHandler:
    Call LogError("UpdateKPISummary", Err.Number, Err.Description)
End Sub

' �ݏo���ꗗ�X�V
Public Sub UpdateCurrentLendingList()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = GetWorksheet(SHEET_DASHBOARD)
    If ws Is Nothing Then Exit Sub
    
    ' �ݏo���ꗗ���쐬�iA8:F20�͈́j
    Dim startRange As Range
    Set startRange = ws.Range("A8")
    
    ' �w�b�_�[�쐬
    Call CreateLendingListHeader(startRange)
    
    ' �f�[�^�쐬
    Call PopulateLendingListData(startRange)
    
    Exit Sub
    
ErrHandler:
    Call LogError("UpdateCurrentLendingList", Err.Number, Err.Description)
End Sub

' �݌ɏ󋵍X�V
Public Sub UpdateStockStatus()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = GetWorksheet(SHEET_DASHBOARD)
    If ws Is Nothing Then Exit Sub
    
    ' �݌ɏ󋵈ꗗ���쐬�iH8:L20�͈́j
    Dim startRange As Range
    Set startRange = ws.Range("H8")
    
    ' �w�b�_�[�쐬
    Call CreateStockStatusHeader(startRange)
    
    ' �f�[�^�쐬
    Call PopulateStockStatusData(startRange)
    
    Exit Sub
    
ErrHandler:
    Call LogError("UpdateStockStatus", Err.Number, Err.Description)
End Sub

' �������߈ꗗ�X�V
Public Sub UpdateOverdueList()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = GetWorksheet(SHEET_DASHBOARD)
    If ws Is Nothing Then Exit Sub
    
    ' �������߈ꗗ���쐬�iA22:F35�͈́j
    Dim startRange As Range
    Set startRange = ws.Range("A22")
    
    ' �w�b�_�[�쐬
    Call CreateOverdueListHeader(startRange)
    
    ' �f�[�^�쐬
    Call PopulateOverdueListData(startRange)
    
    Exit Sub
    
ErrHandler:
    Call LogError("UpdateOverdueList", Err.Number, Err.Description)
End Sub

' �����i���擾�i�����֐��j
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

' �����p�\�����擾�i�����֐��j
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

' KPI�F�����t�H�[�}�b�g�K�p�i�����֐��j
Private Sub ApplyKPIColorFormatting(ws As Worksheet, overdueCount As Long, lendingCount As Long)
    On Error Resume Next
    
    ' �������ߌ����̐F����
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
    
    ' �ݏo�������̐F����
    If lendingCount > 10 Then ' �x��臒l
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

' �ݏo���ꗗ�w�b�_�[�쐬�i�����֐��j
Private Sub CreateLendingListHeader(startRange As Range)
    On Error Resume Next
    
    Dim headerRange As Range
    Set headerRange = startRange.Resize(1, 6)
    
    ' �w�b�_�[�ݒ�
    headerRange.Cells(1, 1).Value = "���iID"
    headerRange.Cells(1, 2).Value = "���i��"
    headerRange.Cells(1, 3).Value = "�ؗp��"
    headerRange.Cells(1, 4).Value = "�ݏo��"
    headerRange.Cells(1, 5).Value = "�ԋp����"
    headerRange.Cells(1, 6).Value = "���ߓ���"
    
    ' �w�b�_�[�t�H�[�}�b�g
    With headerRange
        .Font.Bold = True
        .Interior.Color = COLOR_HEADER
        .Font.Color = vbWhite
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    On Error GoTo 0
End Sub

' �ݏo���ꗗ�f�[�^�쐬�i�����֐��j
Private Sub PopulateLendingListData(startRange As Range)
    On Error GoTo ErrHandler
    
    ' �f�[�^�͈̓N���A
    Dim dataRange As Range
    Set dataRange = startRange.Offset(1, 0).Resize(12, 6)
    dataRange.ClearContents
    dataRange.Interior.Color = COLOR_NORMAL
    
    Dim tbl As ListObject
    Set tbl = GetLendingTable()
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    
    ' ��C���f�b�N�X�擾
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
            ' �f�[�^�ݒ�
            dataRange.Cells(row, 1).Value = tbl.DataBodyRange.Cells(i, itemIDCol).Value
            dataRange.Cells(row, 2).Value = tbl.DataBodyRange.Cells(i, itemNameCol).Value
            dataRange.Cells(row, 3).Value = tbl.DataBodyRange.Cells(i, borrowerCol).Value
            dataRange.Cells(row, 4).Value = tbl.DataBodyRange.Cells(i, lendDateCol).Value
            dataRange.Cells(row, 5).Value = tbl.DataBodyRange.Cells(i, dueDateCol).Value
            
            ' ���ߓ����v�Z�ƐF����
            If IsDate(tbl.DataBodyRange.Cells(i, dueDateCol).Value) Then
                Dim dueDate As Date, overdueDays As Long
                dueDate = CDate(tbl.DataBodyRange.Cells(i, dueDateCol).Value)
                overdueDays = Date - dueDate
                
                If overdueDays > 0 Then
                    dataRange.Cells(row, 6).Value = overdueDays & "������"
                    ' �������ߍs��ԐF�ŋ���
                    dataRange.Rows(row).Interior.Color = COLOR_OVERDUE
                    dataRange.Rows(row).Font.Color = vbWhite
                ElseIf overdueDays >= -WARNING_DAYS_BEFORE Then
                    dataRange.Cells(row, 6).Value = "�����ԋ�"
                    ' �����ԋߍs�����F�Ōx��
                    dataRange.Rows(row).Interior.Color = COLOR_WARNING
                    dataRange.Rows(row).Font.Color = vbBlack
                Else
                    dataRange.Cells(row, 6).Value = "����"
                End If
            End If
            
            row = row + 1
        End If
    Next i
    
    Exit Sub
    
ErrHandler:
    Call LogError("PopulateLendingListData", Err.Number, Err.Description)
End Sub

' �݌ɏ󋵃w�b�_�[�쐬�i�����֐��j
Private Sub CreateStockStatusHeader(startRange As Range)
    On Error Resume Next
    
    Dim headerRange As Range
    Set headerRange = startRange.Resize(1, 5)
    
    ' �w�b�_�[�ݒ�
    headerRange.Cells(1, 1).Value = "���iID"
    headerRange.Cells(1, 2).Value = "���i��"
    headerRange.Cells(1, 3).Value = "���݌�"
    headerRange.Cells(1, 4).Value = "�ݏo��"
    headerRange.Cells(1, 5).Value = "���p�\"
    
    ' �w�b�_�[�t�H�[�}�b�g
    With headerRange
        .Font.Bold = True
        .Interior.Color = COLOR_HEADER
        .Font.Color = vbWhite
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    On Error GoTo 0
End Sub

' �݌ɏ󋵃f�[�^�쐬�i�����֐��j
Private Sub PopulateStockStatusData(startRange As Range)
    On Error GoTo ErrHandler
    
    ' �f�[�^�͈̓N���A
    Dim dataRange As Range
    Set dataRange = startRange.Offset(1, 0).Resize(12, 5)
    dataRange.ClearContents
    dataRange.Interior.Color = COLOR_NORMAL
    
    Dim tbl As ListObject
    Set tbl = GetItemsTable()
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    
    ' ��C���f�b�N�X�擾
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
            
            ' �f�[�^�ݒ�
            dataRange.Cells(row, 1).Value = itemID
            dataRange.Cells(row, 2).Value = tbl.DataBodyRange.Cells(i, itemNameCol).Value
            dataRange.Cells(row, 3).Value = totalQty
            dataRange.Cells(row, 4).Value = lendingQty
            dataRange.Cells(row, 5).Value = availableQty
            
            ' �݌ɐ؂�x���F����
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

' �������߈ꗗ�w�b�_�[�쐬�i�����֐��j
Private Sub CreateOverdueListHeader(startRange As Range)
    On Error Resume Next
    
    Dim headerRange As Range
    Set headerRange = startRange.Resize(1, 6)
    
    ' �w�b�_�[�ݒ�
    headerRange.Cells(1, 1).Value = "���iID"
    headerRange.Cells(1, 2).Value = "���i��"
    headerRange.Cells(1, 3).Value = "�ؗp��"
    headerRange.Cells(1, 4).Value = "�ݏo��"
    headerRange.Cells(1, 5).Value = "�ԋp����"
    headerRange.Cells(1, 6).Value = "���ߓ���"
    
    ' �w�b�_�[�t�H�[�}�b�g�i�������߂͐ԐF�j
    With headerRange
        .Font.Bold = True
        .Interior.Color = COLOR_OVERDUE
        .Font.Color = vbWhite
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    On Error GoTo 0
End Sub

' �������߈ꗗ�f�[�^�쐬�i�����֐��j
Private Sub PopulateOverdueListData(startRange As Range)
    On Error GoTo ErrHandler
    
    ' �f�[�^�͈̓N���A
    Dim dataRange As Range
    Set dataRange = startRange.Offset(1, 0).Resize(13, 6)
    dataRange.ClearContents
    dataRange.Interior.Color = COLOR_NORMAL
    
    Dim tbl As ListObject
    Set tbl = GetLendingTable()
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    
    ' ��C���f�b�N�X�擾
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
                
                If overdueDays > 0 Then ' �������߂̂ݕ\��
                    ' �f�[�^�ݒ�
                    dataRange.Cells(row, 1).Value = tbl.DataBodyRange.Cells(i, itemIDCol).Value
                    dataRange.Cells(row, 2).Value = tbl.DataBodyRange.Cells(i, itemNameCol).Value
                    dataRange.Cells(row, 3).Value = tbl.DataBodyRange.Cells(i, borrowerCol).Value
                    dataRange.Cells(row, 4).Value = tbl.DataBodyRange.Cells(i, lendDateCol).Value
                    dataRange.Cells(row, 5).Value = dueDate
                    dataRange.Cells(row, 6).Value = overdueDays & "������"
                    
                    ' �������ߍs��ԐF�ŋ���
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