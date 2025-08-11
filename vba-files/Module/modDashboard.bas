Attribute VB_Name = "modDashboard"
'=============================================================================
' modDashboard.bas - �_�b�V���{�[�hUI�EKPI�Ǘ����W���[��
'=============================================================================
' �T�v:
'   Dashboard�V�[�g�� UI����AKPI�\���A�{�^���C�x���g����
'   ���v���̎��W�E�\���A���|�[�g�����@�\���
'=============================================================================
Option Explicit

'=============================================================================
' �_�b�V���{�[�h������
'=============================================================================

' �_�b�V���{�[�h�V�[�g������
Public Sub InitializeDashboard()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_DASHBOARD)
    If ws Is Nothing Then
        Call modCmn.LogError("InitializeDashboard", "Dashboard�V�[�g��������܂���")
        Exit Sub
    End If
    
    ' �V�[�g���ݒ�
    ws.Name = SHEET_DASHBOARD
    
    ' �t�H���g�K�p
    Call modCmn.ApplySheetFont(ws)
    
    ' UI�v�f�쐬
    Call CreateDashboardLayout(ws)
    Call CreateDashboardButtons(ws)
    
    ' ����KPI�\��
    Call RefreshKPI()
    
    Call modCmn.LogInfo("InitializeDashboard", "�_�b�V���{�[�h����������")
    Exit Sub
    
ErrHandler:
    Call modCmn.LogError("InitializeDashboard", "�_�b�V���{�[�h�������G���[: " & Err.Description)
End Sub

' �_�b�V���{�[�h���C�A�E�g�쐬
Private Sub CreateDashboardLayout(ByVal ws As Worksheet)
    On Error Resume Next
    
    ' �w�b�_�[
    With ws.Range("A1")
        .Value = SYSTEM_NAME
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(0, 100, 200)
    End With
    
    With ws.Range("A2")
        .Value = "�o�[�W����: " & SYSTEM_VERSION
        .Font.Size = 10
        .Font.Color = RGB(100, 100, 100)
    End With
    
    ' KPI�Z�N�V�����w�b�_�[
    With ws.Range("A4")
        .Value = "=== �V�X�e���� ==="
        .Font.Size = 12
        .Font.Bold = True
    End With
    
    ' KPI����
    ws.Range("B5").Value = "���ڋq��:"
    ws.Range("B6").Value = "�ǉ�����:"
    ws.Range("B7").Value = "�X�V����:"
    ws.Range("B8").Value = "�d�����o:"
    ws.Range("B9").Value = "�G���[����:"
    ws.Range("B10").Value = "����������:"
    ws.Range("B11").Value = "�ŏI�捞����:"
    ws.Range("B12").Value = "��������:"
    
    ' KPI�l�Z���i�E�����j
    ws.Range("D5:D12").HorizontalAlignment = xlRight
    
    ' �{�^���Z�N�V�����w�b�_�[
    With ws.Range("A14")
        .Value = "=== ���상�j���[ ==="
        .Font.Size = 12
        .Font.Bold = True
    End With
    
    ' �񕝒���
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 15
    ws.Columns("C").ColumnWidth = 2
    ws.Columns("D").ColumnWidth = 15
    ws.Columns("E").ColumnWidth = 20
End Sub

' �_�b�V���{�[�h�{�^���쐬
Private Sub CreateDashboardButtons(ByVal ws As Worksheet)
    On Error GoTo ErrHandler
    
    Dim btn As Button
    
    ' ���C���{�^��: CSV��荞�݁����`�����؁����f
    Set btn = ws.Buttons.Add(30, 240, 200, 30) ' B15�Z���ʒu
    With btn
        .Caption = "CSV�ꊇ�捞�E�X�V���s"
        .OnAction = "Sheet1.ExecuteFullImportProcess"
        .Font.Bold = True
    End With
    
    ' �T�u�{�^��1: Staging�N���A
    Set btn = ws.Buttons.Add(30, 280, 150, 25) ' B16�Z���ʒu
    With btn
        .Caption = "Staging�f�[�^�N���A"
        .OnAction = "Sheet1.ClearStagingData"
    End With
    
    ' �T�u�{�^��2: �ݒ���
    Set btn = ws.Buttons.Add(30, 315, 150, 25) ' B17�Z���ʒu
    With btn
        .Caption = "�ݒ��ʂ��J��"
        .OnAction = "Sheet1.OpenConfigSheet"
    End With
    
    ' �T�u�{�^��3: ���|�[�g�o��
    Set btn = ws.Buttons.Add(30, 350, 150, 25) ' B18�Z���ʒu
    With btn
        .Caption = "�������|�[�g�o��"
        .OnAction = "Sheet1.ExportDifferenceReport"
    End With
    
    ' �T�u�{�^��4: KPI�X�V
    Set btn = ws.Buttons.Add(30, 385, 150, 25) ' B19�Z���ʒu
    With btn
        .Caption = "KPI�\���X�V"
        .OnAction = "Sheet1.RefreshKPI"
    End With
    
    ' �T�u�{�^��5: �����؂ꖳ����
    Set btn = ws.Buttons.Add(220, 280, 150, 25) ' D16�Z���ʒu
    With btn
        .Caption = "�����؂�ڋq������"
        .OnAction = "Sheet1.InactivateStaleCustomers"
    End With
    
    Exit Sub
    
ErrHandler:
    Call modCmn.LogError("CreateDashboardButtons", "�{�^���쐬�G���[: " & Err.Description)
End Sub

'=============================================================================
' KPI�\���E�X�V
'=============================================================================

' KPI�\���X�V
Public Sub RefreshKPI()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim customerCount As Long
    Dim stagingStats As Object
    Dim lastImportDate As Date
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_DASHBOARD)
    If ws Is Nothing Then Exit Sub
    
    ' �ڋq�����擾
    customerCount = GetTotalCustomerCount()
    ws.Range(KPI_TOTAL_CUSTOMERS_CELL).Value = Format(customerCount, NUMBER_FORMAT_COUNT)
    
    ' Staging���v�擾
    Set stagingStats = GetStagingStatistics()
    
    ' KPI�l�ݒ�
    ws.Range(KPI_ADDED_COUNT_CELL).Value = Format(GetRecentAddedCount(), NUMBER_FORMAT_COUNT)
    ws.Range(KPI_UPDATED_COUNT_CELL).Value = Format(GetRecentUpdatedCount(), NUMBER_FORMAT_COUNT)
    ws.Range(KPI_DUPLICATE_COUNT_CELL).Value = Format(GetDuplicateDetectionCount(), NUMBER_FORMAT_COUNT)
    ws.Range(KPI_ERROR_COUNT_CELL).Value = Format(stagingStats("ErrorCount"), NUMBER_FORMAT_COUNT)
    ws.Range(KPI_INACTIVE_COUNT_CELL).Value = Format(GetInactiveCustomerCount(), NUMBER_FORMAT_COUNT)
    
    ' �ŏI�捞����
    lastImportDate = GetLastImportDateTime()
    If lastImportDate > 0 Then
        ws.Range(KPI_LAST_IMPORT_CELL).Value = Format(lastImportDate, DATE_FORMAT_DISPLAY)
    Else
        ws.Range(KPI_LAST_IMPORT_CELL).Value = "�����s"
    End If
    
    ' ��������
    ws.Range(KPI_PROCESS_TIME_CELL).Value = GetLastProcessTime() & " �b"
    
    ' �X�V�����L�^
    ws.Range("D13").Value = "�X�V: " & Format(Now, DATE_FORMAT_DISPLAY)
    ws.Range("D13").Font.Size = 8
    ws.Range("D13").Font.Color = RGB(150, 150, 150)
    
    Call modCmn.LogInfo("RefreshKPI", "KPI�\���X�V����")
    Exit Sub
    
ErrHandler:
    Call modCmn.LogError("RefreshKPI", "KPI�X�V�G���[: " & Err.Description)
End Sub

'=============================================================================
' ���v�f�[�^�擾
'=============================================================================

' ���ڋq���擾
Private Function GetTotalCustomerCount() As Long
    On Error Resume Next
    
    Dim tbl As ListObject
    Set tbl = modData.GetCustomersTable()
    
    If Not tbl Is Nothing Then
        GetTotalCustomerCount = tbl.ListRows.Count
    Else
        GetTotalCustomerCount = 0
    End If
End Function

' Staging���v�擾
Private Function GetStagingStatistics() As Object
    On Error Resume Next
    
    Dim stats As Object
    Dim tbl As ListObject
    Dim row As ListRow
    Dim totalCount As Long
    Dim validCount As Long
    Dim errorCount As Long
    
    Set GetStagingStatistics = CreateObject("Scripting.Dictionary")
    Set stats = GetStagingStatistics
    
    stats("TotalCount") = 0
    stats("ValidCount") = 0
    stats("ErrorCount") = 0
    
    Set tbl = modData.GetStagingTable()
    If tbl Is Nothing Then Exit Function
    
    For Each row In tbl.ListRows
        totalCount = totalCount + 1
        If modCmn.GetRowValue(row, COL_IS_VALID) = True Then
            validCount = validCount + 1
        Else
            errorCount = errorCount + 1
        End If
    Next row
    
    stats("TotalCount") = totalCount
    stats("ValidCount") = validCount
    stats("ErrorCount") = errorCount
End Function

' �ŋ߂̒ǉ������擾
Private Function GetRecentAddedCount() As Long
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    Dim addedCount As Long
    Dim today As Date
    
    GetRecentAddedCount = 0
    today = Date
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_LOGS)
    If ws Is Nothing Then Exit Function
    
    Set tbl = modCmn.GetTable(ws, TABLE_LOGS)
    If tbl Is Nothing Then Exit Function
    
    ' �{���̃A�b�v�T�[�g���O����ǉ��������W�v
    For Each row In tbl.ListRows
        Dim logDate As Date
        Dim message As String
        
        logDate = modCmn.SafeDate(modCmn.GetRowText(row, "Timestamp"))
        message = modCmn.GetRowText(row, "Message")
        
        If DateValue(logDate) = today And InStr(message, "�ǉ�:") > 0 Then
            ' ���b�Z�[�W����ǉ������𒊏o�i�ȈՃp�^�[���}�b�`���O�j
            Dim parts As Variant
            parts = Split(message, "�ǉ�:")
            If UBound(parts) > 0 Then
                Dim countPart As String
                countPart = Trim(Split(parts(1), ",")(0))
                addedCount = addedCount + modCmn.SafeLong(countPart)
            End If
        End If
    Next row
    
    GetRecentAddedCount = addedCount
End Function

' �ŋ߂̍X�V�����擾
Private Function GetRecentUpdatedCount() As Long
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    Dim updatedCount As Long
    Dim today As Date
    
    GetRecentUpdatedCount = 0
    today = Date
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_LOGS)
    If ws Is Nothing Then Exit Function
    
    Set tbl = modCmn.GetTable(ws, TABLE_LOGS)
    If tbl Is Nothing Then Exit Function
    
    ' �{���̃A�b�v�T�[�g���O����X�V�������W�v
    For Each row In tbl.ListRows
        Dim logDate As Date
        Dim message As String
        
        logDate = modCmn.SafeDate(modCmn.GetRowText(row, "Timestamp"))
        message = modCmn.GetRowText(row, "Message")
        
        If DateValue(logDate) = today And InStr(message, "�X�V:") > 0 Then
            ' ���b�Z�[�W����X�V�����𒊏o
            Dim parts As Variant
            parts = Split(message, "�X�V:")
            If UBound(parts) > 0 Then
                Dim countPart As String
                countPart = Trim(Split(parts(1), ",")(0))
                updatedCount = updatedCount + modCmn.SafeLong(countPart)
            End If
        End If
    Next row
    
    GetRecentUpdatedCount = updatedCount
End Function

' �d�����o�����擾
Private Function GetDuplicateDetectionCount() As Long
    On Error Resume Next
    
    Dim tbl As ListObject
    Dim row As ListRow
    Dim duplicateCount As Long
    
    GetDuplicateDetectionCount = 0
    
    Set tbl = modData.GetStagingTable()
    If tbl Is Nothing Then Exit Function
    
    For Each row In tbl.ListRows
        Dim errorMessage As String
        errorMessage = modCmn.GetRowText(row, COL_ERROR_MESSAGE)
        If InStr(errorMessage, "�d��") > 0 Then
            duplicateCount = duplicateCount + 1
        End If
    Next row
    
    GetDuplicateDetectionCount = duplicateCount
End Function

' �����ڋq���擾
Private Function GetInactiveCustomerCount() As Long
    On Error Resume Next
    
    Dim tbl As ListObject
    Dim row As ListRow
    Dim inactiveCount As Long
    
    GetInactiveCustomerCount = 0
    
    Set tbl = modData.GetCustomersTable()
    If tbl Is Nothing Then Exit Function
    
    For Each row In tbl.ListRows
        If modCmn.GetRowText(row, COL_STATUS) = STATUS_INACTIVE Then
            inactiveCount = inactiveCount + 1
        End If
    Next row
    
    GetInactiveCustomerCount = inactiveCount
End Function

' �ŏI�捞�����擾
Private Function GetLastImportDateTime() As Date
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    Dim lastDateTime As Date
    Dim rowDateTime As Date
    
    GetLastImportDateTime = 0
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_LOGS)
    If ws Is Nothing Then Exit Function
    
    Set tbl = modCmn.GetTable(ws, TABLE_LOGS)
    If tbl Is Nothing Then Exit Function
    
    ' ���O����ŐV��CSV�捞�L�^������
    For Each row In tbl.ListRows
        Dim message As String
        message = modCmn.GetRowText(row, "Message")
        If InStr(message, "CSV��荞��") > 0 Or InStr(message, "�A�b�v�T�[�g") > 0 Then
            rowDateTime = modCmn.SafeDate(modCmn.GetRowText(row, "Timestamp"))
            If rowDateTime > lastDateTime Then
                lastDateTime = rowDateTime
            End If
        End If
    Next row
    
    GetLastImportDateTime = lastDateTime
End Function

' �ŏI�������Ԏ擾
Private Function GetLastProcessTime() As String
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    Dim lastDateTime As Date
    Dim lastProcessTime As String
    Dim rowDateTime As Date
    
    GetLastProcessTime = "N/A"
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_LOGS)
    If ws Is Nothing Then Exit Function
    
    Set tbl = modCmn.GetTable(ws, TABLE_LOGS)
    If tbl Is Nothing Then Exit Function
    
    ' �ŐV�̃A�b�v�T�[�g�������Ԃ��擾
    For Each row In tbl.ListRows
        Dim message As String
        message = modCmn.GetRowText(row, "Message")
        If InStr(message, "�A�b�v�T�[�g") > 0 Then
            rowDateTime = modCmn.SafeDate(modCmn.GetRowText(row, "Timestamp"))
            If rowDateTime > lastDateTime Then
                lastDateTime = rowDateTime
                lastProcessTime = modCmn.GetRowText(row, "ProcessTime")
            End If
        End If
    Next row
    
    GetLastProcessTime = lastProcessTime
End Function

'=============================================================================
' ���|�[�g����
'=============================================================================

' �������|�[�g�����E�o��
Public Sub GenerateDifferenceReport()
    On Error GoTo ErrHandler
    
    Dim reportText As String
    Dim filePath As String
    Dim fileNum As Integer
    
    ' ���|�[�g���e����
    reportText = CreateDifferenceReportContent()
    
    ' �t�@�C���o��
    filePath = ThisWorkbook.Path & "\customer_report_" & Format(Now, DATE_FORMAT_FILE) & ".txt"
    fileNum = FreeFile
    
    Open filePath For Output As fileNum
    Print #fileNum, reportText
    Close fileNum
    
    MsgBox "�������|�[�g���o�͂��܂����B" & vbCrLf & filePath, vbInformation
    Call modCmn.LogInfo("GenerateDifferenceReport", "���|�[�g�o�͊���: " & filePath)
    Exit Sub
    
ErrHandler:
    If fileNum > 0 Then Close fileNum
    Call modCmn.LogError("GenerateDifferenceReport", "���|�[�g�o�̓G���[: " & Err.Description)
End Sub

' �������|�[�g���e�쐬
Private Function CreateDifferenceReportContent() As String
    On Error Resume Next
    
    Dim report As String
    Dim stagingStats As Object
    
    Set stagingStats = GetStagingStatistics()
    
    report = "=== �ڋq�f�[�^�Ǘ��V�X�e�� �������|�[�g ===" & vbCrLf & vbCrLf
    report = report & "���|�[�g�쐬����: " & modCmn.GetCurrentDateTimeString() & vbCrLf
    report = report & "�V�X�e���o�[�W����: " & SYSTEM_VERSION & vbCrLf & vbCrLf
    
    report = report & "�y���݂̏󋵁z" & vbCrLf
    report = report & "���ڋq��: " & Format(GetTotalCustomerCount(), NUMBER_FORMAT_COUNT) & " ��" & vbCrLf
    report = report & "�L���ڋq��: " & Format(GetTotalCustomerCount() - GetInactiveCustomerCount(), NUMBER_FORMAT_COUNT) & " ��" & vbCrLf
    report = report & "�����ڋq��: " & Format(GetInactiveCustomerCount(), NUMBER_FORMAT_COUNT) & " ��" & vbCrLf & vbCrLf
    
    report = report & "�y�ŐV�������ʁz" & vbCrLf
    report = report & "�{���ǉ�����: " & Format(GetRecentAddedCount(), NUMBER_FORMAT_COUNT) & " ��" & vbCrLf
    report = report & "�{���X�V����: " & Format(GetRecentUpdatedCount(), NUMBER_FORMAT_COUNT) & " ��" & vbCrLf
    report = report & "�d�����o����: " & Format(GetDuplicateDetectionCount(), NUMBER_FORMAT_COUNT) & " ��" & vbCrLf
    report = report & "�G���[����: " & Format(stagingStats("ErrorCount"), NUMBER_FORMAT_COUNT) & " ��" & vbCrLf & vbCrLf
    
    report = report & "�y���������z" & vbCrLf
    Dim lastImport As Date
    lastImport = GetLastImportDateTime()
    If lastImport > 0 Then
        report = report & "�ŏI�捞����: " & Format(lastImport, DATE_FORMAT_DISPLAY) & vbCrLf
        report = report & "��������: " & GetLastProcessTime() & vbCrLf
    Else
        report = report & "�捞����: �Ȃ�" & vbCrLf
    End If
    
    report = report & vbCrLf & "�y���،��ʏڍׁz" & vbCrLf
    report = report & modValidation.GenerateValidationReport()
    
    CreateDifferenceReportContent = report
End Function

'=============================================================================
' ���[�e�B���e�B�֐�
'=============================================================================

' �V�X�e����ԃ`�F�b�N
Public Function CheckSystemHealth() As Boolean
    On Error Resume Next
    
    CheckSystemHealth = True
    
    ' �K�v�ȃV�[�g�̑��݃`�F�b�N
    If modCmn.GetWorksheetByIndex(SHEET_INDEX_CUSTOMERS) Is Nothing Then CheckSystemHealth = False
    If modCmn.GetWorksheetByIndex(SHEET_INDEX_STAGING) Is Nothing Then CheckSystemHealth = False
    If modCmn.GetWorksheetByIndex(SHEET_INDEX_CONFIG) Is Nothing Then CheckSystemHealth = False
    If modCmn.GetWorksheetByIndex(SHEET_INDEX_LOGS) Is Nothing Then CheckSystemHealth = False
    
    ' �K�v�ȃe�[�u���̑��݃`�F�b�N
    If modData.GetCustomersTable() Is Nothing Then CheckSystemHealth = False
    If modData.GetStagingTable() Is Nothing Then CheckSystemHealth = False
End Function

' �_�b�V���{�[�h�\�����b�Z�[�W
Public Sub ShowStatusMessage(ByVal message As String, Optional ByVal isError As Boolean = False)
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_DASHBOARD)
    If ws Is Nothing Then Exit Sub
    
    With ws.Range("A20")
        .Value = message
        .Font.Size = 10
        .Font.Bold = True
        If isError Then
            .Font.Color = RGB(200, 0, 0)
        Else
            .Font.Color = RGB(0, 150, 0)
        End If
    End With
    
    ' ��莞�Ԍ�Ƀ��b�Z�[�W���N���A
    Application.OnTime Now + TimeValue("00:00:05"), "modDashboard.ClearStatusMessage"
End Sub

' �X�e�[�^�X���b�Z�[�W�N���A
Public Sub ClearStatusMessage()
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_DASHBOARD)
    If Not ws Is Nothing Then
        ws.Range("A20").ClearContents
    End If
End Sub

' �����m�F�_�C�A���O
Public Function ConfirmOperation(ByVal operation As String) As Boolean
    On Error Resume Next
    
    Dim result As VbMsgBoxResult
    result = MsgBox("�ȉ��̑�������s���܂����H" & vbCrLf & vbCrLf & operation & vbCrLf & vbCrLf & _
                    "���̑���͎��������Ƃ��ł��܂���B", vbQuestion + vbYesNo, "����m�F")
    
    ConfirmOperation = (result = vbYes)
End Function