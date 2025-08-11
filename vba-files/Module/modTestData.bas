Attribute VB_Name = "modTestData"
Option Explicit

' =============================================================================
' ���i�ݏo�Ǘ��V�X�e�� - �e�X�g�f�[�^�������W���[��
' =============================================================================

' �S�e�X�g�f�[�^�쐬���C������
Public Sub CreateAllTestData()
    On Error GoTo ErrHandler
    
    Dim result As VbMsgBoxResult
    result = MsgBox("�e�X�g�f�[�^���쐬���܂��B�����̃f�[�^�͍폜����܂��B���s���܂����H", vbQuestion + vbYesNo)
    
    If result = vbNo Then Exit Sub
    
    Application.ScreenUpdating = False
    
    ' �����f�[�^�N���A
    Call ClearAllData
    
    ' ���i�}�X�^�e�X�g�f�[�^�쐬
    Call CreateItemsTestData
    
    ' �ݏo�����e�X�g�f�[�^�쐬
    Call CreateLendingTestData
    
    ' �_�b�V���{�[�h�X�V
    Call UpdateDashboard
    
    Application.ScreenUpdating = True
    
    ' �č����O�L�^
    Call LogAudit("�e�X�g�f�[�^�쐬", "All test data created successfully")
    
    MsgBox "�e�X�g�f�[�^�̍쐬���������܂����B", vbInformation
    
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Call LogError("CreateAllTestData", Err.Number, Err.Description)
    MsgBox "�e�X�g�f�[�^�쐬���ɃG���[���������܂���: " & Err.Description, vbCritical
End Sub

' �����f�[�^�N���A
Private Sub ClearAllData()
    On Error Resume Next
    
    ' ���i�e�[�u���N���A
    Dim itemsTbl As ListObject
    Set itemsTbl = GetItemsTable()
    If Not itemsTbl Is Nothing Then
        If Not itemsTbl.DataBodyRange Is Nothing Then
            itemsTbl.DataBodyRange.Delete
        End If
    End If
    
    ' �ݏo�e�[�u���N���A
    Dim lendingTbl As ListObject
    Set lendingTbl = GetLendingTable()
    If Not lendingTbl Is Nothing Then
        If Not lendingTbl.DataBodyRange Is Nothing Then
            lendingTbl.DataBodyRange.Delete
        End If
    End If
    
    On Error GoTo 0
End Sub

' ���i�}�X�^�e�X�g�f�[�^�쐬
Private Sub CreateItemsTestData()
    On Error GoTo ErrHandler
    
    Dim tbl As ListObject
    Set tbl = GetItemsTable()
    
    If tbl Is Nothing Then
        Call LogError("CreateItemsTestData", 9, "Items table not found")
        Exit Sub
    End If
    
    ' �e�X�g�f�[�^�z��
    Dim testItems As Variant
    testItems = Array( _
        Array(1001, "�m�[�gPC - ThinkPad", CATEGORY_PC, LOCATION_OFFICE_1F, 5), _
        Array(1002, "�m�[�gPC - MacBook Air", CATEGORY_PC, LOCATION_OFFICE_1F, 3), _
        Array(1003, "�f�X�N�g�b�vPC - iMac", CATEGORY_PC, LOCATION_OFFICE_2F, 2), _
        Array(2001, "�v���W�F�N�^�[ - EPSON", CATEGORY_AV, LOCATION_MEETING_ROOM, 4), _
        Array(2002, "���j�^�[ 24inch", CATEGORY_AV, LOCATION_OFFICE_2F, 6), _
        Array(2003, "Web�J���� - Logitech", CATEGORY_AV, LOCATION_OFFICE_1F, 8), _
        Array(3001, "�d�� - CASIO", CATEGORY_STATIONERY, LOCATION_OFFICE_1F, 10), _
        Array(3002, "USB������ 32GB", CATEGORY_STATIONERY, LOCATION_WAREHOUSE, 15), _
        Array(3003, "�}�E�X - ���C�����X", CATEGORY_STATIONERY, LOCATION_OFFICE_1F, 12), _
        Array(4001, "�e�X�^�[ - �f�W�^��", CATEGORY_TOOL, LOCATION_WAREHOUSE, 3), _
        Array(4002, "�h���C�o�[�Z�b�g", CATEGORY_TOOL, LOCATION_WAREHOUSE, 5), _
        Array(4003, "LAN �P�[�u���e�X�^�[", CATEGORY_TOOL, LOCATION_OFFICE_2F, 2), _
        Array(5001, "�����R�[�h 10m", CATEGORY_OTHER, LOCATION_WAREHOUSE, 8), _
        Array(5002, "�z���C�g�{�[�h�p�}�[�J�[", CATEGORY_OTHER, LOCATION_MEETING_ROOM, 20), _
        Array(5003, "���ރo�C���_�[", CATEGORY_OTHER, LOCATION_OFFICE_1F, 25) _
    )
    
    ' �f�[�^�ǉ�
    Call AddItemsData(tbl, testItems)
    
    Exit Sub
    
ErrHandler:
    Call LogError("CreateItemsTestData", Err.Number, Err.Description)
End Sub

' ���i�f�[�^�ǉ��i�����֐��j
Private Sub AddItemsData(tbl As ListObject, dataArray As Variant)
    On Error GoTo ErrHandler
    
    Dim i As Long
    For i = 0 To UBound(dataArray)
        Dim newRow As ListRow
        Set newRow = tbl.ListRows.Add()
        
        ' ��C���f�b�N�X�擾
        Dim itemIDCol As Long, itemNameCol As Long, categoryCol As Long
        Dim locationCol As Long, quantityCol As Long
        itemIDCol = GetColumnIndex(tbl, COL_ITEM_ID)
        itemNameCol = GetColumnIndex(tbl, COL_ITEM_NAME)
        categoryCol = GetColumnIndex(tbl, COL_CATEGORY)
        locationCol = GetColumnIndex(tbl, COL_LOCATION)
        quantityCol = GetColumnIndex(tbl, COL_QUANTITY)
        
        If itemIDCol > 0 Then newRow.Range.Cells(1, itemIDCol).Value = dataArray(i)(0)
        If itemNameCol > 0 Then newRow.Range.Cells(1, itemNameCol).Value = dataArray(i)(1)
        If categoryCol > 0 Then newRow.Range.Cells(1, categoryCol).Value = dataArray(i)(2)
        If locationCol > 0 Then newRow.Range.Cells(1, locationCol).Value = dataArray(i)(3)
        If quantityCol > 0 Then newRow.Range.Cells(1, quantityCol).Value = dataArray(i)(4)
    Next i
    
    ' �e�[�u���t�H�[�}�b�g�K�p
    Call ApplyStandardTableFormat(tbl)
    
    Exit Sub
    
ErrHandler:
    Call LogError("AddItemsData", Err.Number, Err.Description)
End Sub

' �ݏo�����e�X�g�f�[�^�쐬
Private Sub CreateLendingTestData()
    On Error GoTo ErrHandler
    
    Dim tbl As ListObject
    Set tbl = GetLendingTable()
    
    If tbl Is Nothing Then
        Call LogError("CreateLendingTestData", 9, "Lending table not found")
        Exit Sub
    End If
    
    ' ���݂̓��t����Ƀe�X�g�f�[�^���쐬
    Dim baseDate As Date
    baseDate = Date
    
    ' �e�X�g�f�[�^�z��i�������߁E�����ԋ߁E����ȑݏo���܂ށj
    Dim testLendings As Variant
    testLendings = Array( _
        Array(1001, "�c�����Y", baseDate - 10, baseDate - 3, "", STATUS_LENDING), _
        Array(2001, "�����Ԏq", baseDate - 8, baseDate - 1, "", STATUS_LENDING), _
        Array(1002, "��؈�Y", baseDate - 5, baseDate + 2, "", STATUS_LENDING), _
        Array(3001, "��������", baseDate - 3, baseDate + 4, "", STATUS_LENDING), _
        Array(2002, "�R�c���Y", baseDate - 12, baseDate - 5, baseDate - 2, STATUS_RETURNED), _
        Array(4001, "�n�ӎO�Y", baseDate - 7, baseDate, "", STATUS_LENDING), _
        Array(1003, "�ɓ��l�Y", baseDate - 4, baseDate + 3, "", STATUS_LENDING), _
        Array(3002, "���ьܘY", baseDate - 15, baseDate - 8, baseDate - 6, STATUS_RETURNED), _
        Array(2003, "�����Z�Y", baseDate - 2, baseDate + 5, "", STATUS_LENDING), _
        Array(5001, "���{���q", baseDate - 9, baseDate - 2, baseDate - 1, STATUS_RETURNED) _
    )
    
    ' �f�[�^�ǉ�
    Call AddLendingData(tbl, testLendings)
    
    Exit Sub
    
ErrHandler:
    Call LogError("CreateLendingTestData", Err.Number, Err.Description)
End Sub

' �ݏo�f�[�^�ǉ��i�����֐��j
Private Sub AddLendingData(tbl As ListObject, dataArray As Variant)
    On Error GoTo ErrHandler
    
    Dim i As Long, recordID As Long
    recordID = 1
    
    For i = 0 To UBound(dataArray)
        Dim newRow As ListRow
        Set newRow = tbl.ListRows.Add()
        
        ' ��C���f�b�N�X�擾
        Dim recordIDCol As Long, itemIDCol As Long, itemNameCol As Long
        Dim borrowerCol As Long, lendDateCol As Long, dueDateCol As Long
        Dim returnDateCol As Long, statusCol As Long, remarksCol As Long
        
        recordIDCol = GetColumnIndex(tbl, COL_RECORD_ID)
        itemIDCol = GetColumnIndex(tbl, COL_LENDING_ITEM_ID)
        itemNameCol = GetColumnIndex(tbl, COL_LENDING_ITEM_NAME)
        borrowerCol = GetColumnIndex(tbl, COL_BORROWER)
        lendDateCol = GetColumnIndex(tbl, COL_LEND_DATE)
        dueDateCol = GetColumnIndex(tbl, COL_DUE_DATE)
        returnDateCol = GetColumnIndex(tbl, COL_RETURN_DATE)
        statusCol = GetColumnIndex(tbl, COL_STATUS)
        remarksCol = GetColumnIndex(tbl, COL_REMARKS)
        
        ' �f�[�^�ݒ�
        If recordIDCol > 0 Then newRow.Range.Cells(1, recordIDCol).Value = recordID
        ' �ꎞ�ϐ��ɒl���R�s�[����ByRef�������
        Dim tempItemID As Long
        tempItemID = dataArray(i)(0)
        
        If itemIDCol > 0 Then newRow.Range.Cells(1, itemIDCol).Value = tempItemID
        If itemNameCol > 0 Then
            ' ���i���������擾
            Dim itemName As String
            itemName = GetItemName(tempItemID)
            newRow.Range.Cells(1, itemNameCol).Value = itemName
        End If
        If borrowerCol > 0 Then newRow.Range.Cells(1, borrowerCol).Value = dataArray(i)(1)
        If lendDateCol > 0 Then newRow.Range.Cells(1, lendDateCol).Value = dataArray(i)(2)
        If dueDateCol > 0 Then newRow.Range.Cells(1, dueDateCol).Value = dataArray(i)(3)
        If returnDateCol > 0 Then newRow.Range.Cells(1, returnDateCol).Value = dataArray(i)(4)
        If statusCol > 0 Then newRow.Range.Cells(1, statusCol).Value = dataArray(i)(5)
        If remarksCol > 0 Then
            ' �X�e�[�^�X�ɉ����Ĕ��l�������ݒ�i�ꎞ�ϐ���ByRef������j
            Dim tempStatus As String, tempDueDate As Date
            tempStatus = dataArray(i)(5)
            tempDueDate = dataArray(i)(3)
            
            If tempStatus = STATUS_LENDING Then
                If tempDueDate < Date Then
                    newRow.Range.Cells(1, remarksCol).Value = "�������ߒ�"
                ElseIf tempDueDate <= Date + WARNING_DAYS_BEFORE Then
                    newRow.Range.Cells(1, remarksCol).Value = "�����ԋ�"
                Else
                    newRow.Range.Cells(1, remarksCol).Value = "����ݏo��"
                End If
            Else
                newRow.Range.Cells(1, remarksCol).Value = "����ԋp����"
            End If
        End If
        
        recordID = recordID + 1
    Next i
    
    ' �e�[�u���t�H�[�}�b�g�K�p
    Call ApplyStandardTableFormat(tbl)
    
    Exit Sub
    
ErrHandler:
    Call LogError("AddLendingData", Err.Number, Err.Description)
End Sub

' �T���v�����i�ǉ�
Public Sub AddSampleItem()
    On Error GoTo ErrHandler
    
    Dim tbl As ListObject
    Set tbl = GetItemsTable()
    
    If tbl Is Nothing Then
        MsgBox "���i�e�[�u����������܂���B", vbExclamation
        Exit Sub
    End If
    
    ' ����ID���v�Z
    Dim nextID As Long
    nextID = 6001 ' �T���v���pID�x�[�X
    
    If Not tbl.DataBodyRange Is Nothing Then
        Dim itemIDCol As Long
        itemIDCol = GetColumnIndex(tbl, COL_ITEM_ID)
        If itemIDCol > 0 Then
            Dim i As Long, maxID As Long
            maxID = 0
            For i = 1 To tbl.DataBodyRange.Rows.Count
                If IsNumeric(tbl.DataBodyRange.Cells(i, itemIDCol).Value) Then
                    If tbl.DataBodyRange.Cells(i, itemIDCol).Value > maxID Then
                        maxID = tbl.DataBodyRange.Cells(i, itemIDCol).Value
                    End If
                End If
            Next i
            nextID = maxID + 1
        End If
    End If
    
    ' �V�������i��ǉ�
    Dim newRow As ListRow
    Set newRow = tbl.ListRows.Add()
    
    newRow.Range.Cells(1, GetColumnIndex(tbl, COL_ITEM_ID)).Value = nextID
    newRow.Range.Cells(1, GetColumnIndex(tbl, COL_ITEM_NAME)).Value = "�T���v�����i - " & Format(Now, "hhmmss")
    newRow.Range.Cells(1, GetColumnIndex(tbl, COL_CATEGORY)).Value = CATEGORY_OTHER
    newRow.Range.Cells(1, GetColumnIndex(tbl, COL_LOCATION)).Value = LOCATION_OFFICE_1F
    newRow.Range.Cells(1, GetColumnIndex(tbl, COL_QUANTITY)).Value = 1
    
    ' �t�H�[�}�b�g�K�p
    Call ApplyStandardTableFormat(tbl)
    
    ' �_�b�V���{�[�h�X�V
    Call UpdateDashboard
    
    ' ���O�L�^
    Call LogAudit("�T���v�����i�ǉ�", "ItemID: " & nextID)
    
    MsgBox "�T���v�����i�iID: " & nextID & "�j��ǉ����܂����B", vbInformation
    
    Exit Sub
    
ErrHandler:
    Call LogError("AddSampleItem", Err.Number, Err.Description)
    MsgBox "�T���v�����i�ǉ����ɃG���[���������܂���: " & Err.Description, vbCritical
End Sub

' �e�X�g�p�ݏo�쐬
Public Sub CreateTestLending()
    On Error GoTo ErrHandler
    
    ' ���p�\�Ȕ��i������
    Dim tbl As ListObject
    Set tbl = GetItemsTable()
    
    If tbl Is Nothing Then
        MsgBox "���i�e�[�u����������܂���B", vbExclamation
        Exit Sub
    End If
    
    Dim availableItemID As Long
    availableItemID = 0
    
    ' �݌ɂ̂�����i������
    If Not tbl.DataBodyRange Is Nothing Then
        Dim itemIDCol As Long
        itemIDCol = GetColumnIndex(tbl, COL_ITEM_ID)
        If itemIDCol > 0 Then
            Dim i As Long
            For i = 1 To tbl.DataBodyRange.Rows.Count
                Dim itemID As Long
                itemID = tbl.DataBodyRange.Cells(i, itemIDCol).Value
                If GetAvailableQuantity(itemID) > 0 Then
                    availableItemID = itemID
                    Exit For
                End If
            Next i
        End If
    End If
    
    If availableItemID = 0 Then
        MsgBox "�ݏo�\�Ȕ��i������܂���B", vbExclamation
        Exit Sub
    End If
    
    ' �ݏo�L�^���쐬
    Dim lendingTbl As ListObject
    Set lendingTbl = GetLendingTable()
    
    If lendingTbl Is Nothing Then
        MsgBox "�ݏo�e�[�u����������܂���B", vbExclamation
        Exit Sub
    End If
    
    Dim newRow As ListRow
    Set newRow = lendingTbl.ListRows.Add()
    
    newRow.Range.Cells(1, GetColumnIndex(lendingTbl, COL_RECORD_ID)).Value = GetNextRecordID()
    newRow.Range.Cells(1, GetColumnIndex(lendingTbl, COL_LENDING_ITEM_ID)).Value = availableItemID
    newRow.Range.Cells(1, GetColumnIndex(lendingTbl, COL_LENDING_ITEM_NAME)).Value = GetItemName(availableItemID)
    newRow.Range.Cells(1, GetColumnIndex(lendingTbl, COL_BORROWER)).Value = "�e�X�g���[�U�[ " & Format(Now, "mmss")
    newRow.Range.Cells(1, GetColumnIndex(lendingTbl, COL_LEND_DATE)).Value = Date
    newRow.Range.Cells(1, GetColumnIndex(lendingTbl, COL_DUE_DATE)).Value = Date + DEFAULT_LENDING_DAYS
    newRow.Range.Cells(1, GetColumnIndex(lendingTbl, COL_RETURN_DATE)).Value = ""
    newRow.Range.Cells(1, GetColumnIndex(lendingTbl, COL_STATUS)).Value = STATUS_LENDING
    newRow.Range.Cells(1, GetColumnIndex(lendingTbl, COL_REMARKS)).Value = "�e�X�g�ݏo�f�[�^"
    
    ' �t�H�[�}�b�g�K�p
    Call ApplyStandardTableFormat(lendingTbl)
    
    ' �_�b�V���{�[�h�X�V
    Call UpdateDashboard
    
    ' ���O�L�^
    Call LogAudit("�e�X�g�ݏo�쐬", "ItemID: " & availableItemID)
    
    MsgBox "�e�X�g�ݏo�i���iID: " & availableItemID & "�j���쐬���܂����B", vbInformation
    
    Exit Sub
    
ErrHandler:
    Call LogError("CreateTestLending", Err.Number, Err.Description)
    MsgBox "�e�X�g�ݏo�쐬���ɃG���[���������܂���: " & Err.Description, vbCritical
End Sub

' �f�o�b�O���\��
Public Sub ShowDebugInfo()
    On Error GoTo ErrHandler
    
    Dim info As String
    info = "=== �V�X�e����ԃf�o�b�O��� ===" & vbCrLf & vbCrLf
    
    ' �e�[�u�����
    Dim itemsTbl As ListObject, lendingTbl As ListObject
    Set itemsTbl = GetItemsTable()
    Set lendingTbl = GetLendingTable()
    
    info = info & "���i�e�[�u��: "
    If itemsTbl Is Nothing Then
        info = info & "�Ȃ�" & vbCrLf
    Else
        info = info & "����i�s��: "
        If itemsTbl.DataBodyRange Is Nothing Then
            info = info & "0"
        Else
            info = info & itemsTbl.DataBodyRange.Rows.Count
        End If
        info = info & "�j" & vbCrLf
    End If
    
    info = info & "�ݏo�e�[�u��: "
    If lendingTbl Is Nothing Then
        info = info & "�Ȃ�" & vbCrLf
    Else
        info = info & "����i�s��: "
        If lendingTbl.DataBodyRange Is Nothing Then
            info = info & "0"
        Else
            info = info & lendingTbl.DataBodyRange.Rows.Count
        End If
        info = info & "�j" & vbCrLf
    End If
    
    ' ���v���
    info = info & vbCrLf & "���v���:" & vbCrLf
    info = info & "�����i��: " & GetTotalItemsCount() & vbCrLf
    info = info & "�ݏo������: " & GetTotalLendingCount() & vbCrLf
    info = info & "�������ߌ���: " & GetOverdueCount() & vbCrLf
    
    ' �V�X�e�����
    info = info & vbCrLf & "�V�X�e�����:" & vbCrLf
    info = info & "���ݓ���: " & Format(Now, "yyyy/mm/dd hh:mm:ss") & vbCrLf
    info = info & "���[�U�[: " & Application.UserName & vbCrLf
    info = info & "���[�N�u�b�N: " & ThisWorkbook.Name & vbCrLf
    
    MsgBox info, vbInformation, "�V�X�e���f�o�b�O���"
    
    Exit Sub
    
ErrHandler:
    Call LogError("ShowDebugInfo", Err.Number, Err.Description)
    MsgBox "�f�o�b�O���擾���ɃG���[���������܂���: " & Err.Description, vbCritical
End Sub