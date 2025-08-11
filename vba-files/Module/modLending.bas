Attribute VB_Name = "modLending"
Option Explicit

' =============================================================================
' ���i�ݏo�Ǘ��V�X�e�� - �ݏo�E�ԋp�����r�W�l�X���W�b�N
' =============================================================================

' �ݏo�o�^���C������
Public Sub RegisterLending()
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    
    ' ���͒l�̎擾�ƌ���
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
    
    ' �݌Ƀ`�F�b�N
    If Not CheckStockAvailable(itemID) Then
        MsgBox MSG_INSUFFICIENT_STOCK, vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' �ݏo�L�^��ǉ�
    Call AddLendingRecord(itemID, borrower, lendDate, lendingDays)
    
    ' ���̓t�B�[���h�N���A
    Call ClearInputFields()
    
    ' �_�b�V���{�[�h�X�V
    Call UpdateDashboard
    
    ' �č����O�L�^
    Call LogAudit("�ݏo�o�^", "ItemID:" & itemID & ", Borrower:" & borrower & ", Days:" & lendingDays)
    
    MsgBox "�ݏo��o�^���܂����B", vbInformation
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Call LogError("RegisterLending", Err.Number, Err.Description)
    MsgBox "�ݏo�o�^���ɃG���[���������܂���: " & Err.Description, vbCritical
End Sub

' �ԋp�o�^���C������
Public Sub RegisterReturn()
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    
    ' ���͒l�̎擾�ƌ���
    Dim ws As Worksheet
    Set ws = GetWorksheet(SHEET_INPUT)
    If ws Is Nothing Then
        MsgBox "���̓V�[�g��������܂���B", vbCritical
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' �K�{���̓`�F�b�N
    Dim itemID As Long, borrower As String, returnDate As Date
    
    If Not IsNumeric(ws.Range(INPUT_ITEM_ID).Value) Then
        MsgBox MSG_INVALID_ITEM_ID, vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If
    itemID = CLng(ws.Range(INPUT_ITEM_ID).Value)
    
    borrower = Trim(ws.Range(INPUT_BORROWER).Value)
    If borrower = "" Then
        MsgBox MSG_REQUIRED_FIELD & "(�ؗp��)", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' �ԋp���i�ȗ����͍����j
    If IsDate(ws.Range(INPUT_RETURN_DATE).Value) Then
        returnDate = CDate(ws.Range(INPUT_RETURN_DATE).Value)
    Else
        returnDate = Date
    End If
    
    ' �ݏo�L�^����
    Dim recordRow As Long
    recordRow = FindLendingRecord(itemID, borrower)
    
    If recordRow = 0 Then
        MsgBox MSG_NO_LENDING_RECORD, vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' �ԋp�������s
    Call ProcessReturn(recordRow, returnDate)
    
    ' ���̓t�B�[���h�N���A
    Call ClearInputFields()
    
    ' �_�b�V���{�[�h�X�V
    Call UpdateDashboard
    
    ' �č����O�L�^
    Call LogAudit("�ԋp�o�^", "ItemID:" & itemID & ", Borrower:" & borrower)
    
    MsgBox "�ԋp��o�^���܂����B", vbInformation
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Call LogError("RegisterReturn", Err.Number, Err.Description)
    MsgBox "�ԋp�o�^���ɃG���[���������܂���: " & Err.Description, vbCritical
End Sub

' ���͒l�擾�E���؁i�����֐��j
Private Function GetAndValidateInput() As Variant
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = GetWorksheet(SHEET_INPUT)
    If ws Is Nothing Then
        MsgBox "���̓V�[�g��������܂���B", vbCritical
        GetAndValidateInput = Empty
        Exit Function
    End If
    
    ' �K�{���̓`�F�b�N
    Dim itemID As Long, borrower As String, lendDate As Date, lendingDays As Long
    
    If Not IsNumeric(ws.Range(INPUT_ITEM_ID).Value) Then
        MsgBox MSG_INVALID_ITEM_ID, vbExclamation
        GetAndValidateInput = Empty
        Exit Function
    End If
    itemID = CLng(ws.Range(INPUT_ITEM_ID).Value)
    
    borrower = Trim(ws.Range(INPUT_BORROWER).Value)
    If borrower = "" Then
        MsgBox MSG_REQUIRED_FIELD & "(�ؗp��)", vbExclamation
        GetAndValidateInput = Empty
        Exit Function
    End If
    
    ' �ݏo���i�ȗ����͍����j
    If IsDate(ws.Range(INPUT_LEND_DATE).Value) Then
        lendDate = CDate(ws.Range(INPUT_LEND_DATE).Value)
    Else
        lendDate = Date
    End If
    
    ' �ݏo���ԁi�ȗ����̓f�t�H���g�l�j
    If IsNumeric(ws.Range(INPUT_LENDING_DAYS).Value) Then
        lendingDays = CLng(ws.Range(INPUT_LENDING_DAYS).Value)
    Else
        lendingDays = DEFAULT_LENDING_DAYS
    End If
    
    ' ���ԏ���`�F�b�N
    If lendingDays > MAX_LENDING_DAYS Then
        MsgBox "�ݏo���Ԃ�" & MAX_LENDING_DAYS & "���ȓ��Ŏw�肵�Ă��������B", vbExclamation
        GetAndValidateInput = Empty
        Exit Function
    End If
    
    ' ���i���݃`�F�b�N
    If Not ItemExists(itemID) Then
        MsgBox MSG_ITEM_NOT_FOUND, vbExclamation
        GetAndValidateInput = Empty
        Exit Function
    End If
    
    ' �߂�l�Ƃ��Ĕz��ŕԂ�
    GetAndValidateInput = Array(itemID, borrower, lendDate, lendingDays)
    
    Exit Function
    
ErrHandler:
    Call LogError("GetAndValidateInput", Err.Number, Err.Description)
    GetAndValidateInput = Empty
End Function

' �݌ɉp���`�F�b�N�i�����֐��j
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

' �ݏo�L�^�ǉ��i�����֐��j
Private Sub AddLendingRecord(itemID As Long, borrower As String, lendDate As Date, lendingDays As Long)
    On Error GoTo ErrHandler
    
    Dim tbl As ListObject
    Set tbl = GetLendingTable()
    
    If tbl Is Nothing Then
        Call LogError("AddLendingRecord", 9, "Lending table not found")
        Exit Sub
    End If
    
    ' �V�����s��ǉ�
    Dim newRow As ListRow
    Set newRow = tbl.ListRows.Add()
    
    ' �e��ɒl��ݒ�
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

' �ԋp�������s�i�����֐��j
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
    
    ' �X�e�[�^�X�m�F
    If tbl.DataBodyRange.Cells(recordRow, statusCol).Value = STATUS_RETURNED Then
        MsgBox MSG_ALREADY_RETURNED, vbExclamation
        Exit Sub
    End If
    
    ' �ԋp���ƃX�e�[�^�X���X�V
    tbl.DataBodyRange.Cells(recordRow, returnDateCol).Value = returnDate
    tbl.DataBodyRange.Cells(recordRow, statusCol).Value = STATUS_RETURNED
    
    Exit Sub
    
ErrHandler:
    Call LogError("ProcessReturn", Err.Number, Err.Description)
End Sub

' ���̓t�B�[���h�N���A�i�����֐��j
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

' �������ߌ����擾
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

' ���ݏo�������擾
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