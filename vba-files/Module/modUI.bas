Attribute VB_Name = "modUI"
Option Explicit

' =============================================================================
' ���i�ݏo�Ǘ��V�X�e�� - UI����֐��i�t�H�[�}�b�g�A�{�^���ݒ蓙�j
' =============================================================================

' �W���e�[�u���t�H�[�}�b�g�K�p
Public Sub ApplyStandardTableFormat(tbl As ListObject)
    On Error GoTo ErrHandler
    
    If tbl Is Nothing Then Exit Sub
    
    ' �w�b�_�[�s�̃t�H�[�}�b�g
    With tbl.HeaderRowRange
        .Font.Bold = True
        .Interior.Color = COLOR_HEADER
        .Font.Color = vbWhite
        .Font.Size = 11
        .RowHeight = 25
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' �f�[�^�s�̌��ݐF�ݒ�
    If Not tbl.DataBodyRange Is Nothing Then
        Call ApplyAlternatingRowColors(tbl.DataBodyRange)
    End If
    
    ' �g���̐ݒ�
    With tbl.Range.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(128, 128, 128)
    End With
    
    ' �񕝂̎�������
    tbl.Range.Columns.AutoFit
    
    Exit Sub
    
ErrHandler:
    Call LogError("ApplyStandardTableFormat", Err.Number, Err.Description)
End Sub

' ���ݍs�F�ݒ�֐�
Private Sub ApplyAlternatingRowColors(dataRange As Range)
    On Error GoTo ErrHandler
    
    Dim i As Long
    For i = 1 To dataRange.Rows.Count
        If i Mod 2 = 0 Then
            dataRange.Rows(i).Interior.Color = COLOR_ALTERNATE
        Else
            dataRange.Rows(i).Interior.Color = COLOR_NORMAL
        End If
    Next i
    
    Exit Sub
    
ErrHandler:
    Call LogError("ApplyAlternatingRowColors", Err.Number, Err.Description)
End Sub

' �����t���F�����֐�
Public Sub ApplyConditionalFormatting(targetRange As Range, status As String)
    On Error GoTo ErrHandler
    
    Select Case status
        Case "��������", "�G���[", "���s"
            With targetRange
                .Interior.Color = COLOR_OVERDUE
                .Font.Color = vbWhite
                .Font.Bold = True
            End With
        Case "�����ԋ�", "�x��", "����"
            With targetRange
                .Interior.Color = COLOR_WARNING
                .Font.Color = vbBlack
                .Font.Bold = True
            End With
        Case "����", "����", "�ԋp��"
            With targetRange
                .Interior.Color = COLOR_SUCCESS
                .Font.Color = vbWhite
            End With
        Case Else
            With targetRange
                .Interior.Color = COLOR_NORMAL
                .Font.Color = vbBlack
            End With
    End Select
    
    Exit Sub
    
ErrHandler:
    Call LogError("ApplyConditionalFormatting", Err.Number, Err.Description)
End Sub

' �_�b�V���{�[�h���C�A�E�g�쐬
Public Sub CreateDashboardLayout()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = GetWorksheet(SHEET_DASHBOARD)
    If ws Is Nothing Then
        Call LogError("CreateDashboardLayout", 9, "Dashboard sheet not found")
        Exit Sub
    End If
    
    ' �����̓��e���N���A
    ws.Cells.Clear
    
    ' �^�C�g������
    With ws.Range("A1:L1")
        .Merge
        .Value = "���i�ݏo�Ǘ��V�X�e�� - �_�b�V���{�[�h"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 35
    End With
    
    ' KPI�T�}���[�Z�N�V����
    Call CreateKPISummarySection(ws)
    
    ' �{�^���z�u
    Call CreateActionButtons(ws)
    
    ' �f�[�^�\���G���A�̃��x��
    ws.Range("A7").Value = "�� �ݏo���ꗗ"
    ws.Range("A7").Font.Bold = True
    ws.Range("A7").Font.Size = 12
    
    ws.Range("H7").Value = "�� �݌ɏ�"
    ws.Range("H7").Font.Bold = True
    ws.Range("H7").Font.Size = 12
    
    ws.Range("A21").Value = "�� �������߈ꗗ"
    ws.Range("A21").Font.Bold = True
    ws.Range("A21").Font.Size = 12
    ws.Range("A21").Font.Color = COLOR_OVERDUE
    
    ' �����f�[�^�X�V
    Call UpdateDashboard
    
    Exit Sub
    
ErrHandler:
    Call LogError("CreateDashboardLayout", Err.Number, Err.Description)
End Sub

' KPI�T�}���[�Z�N�V�����쐬�i�����֐��j
Private Sub CreateKPISummarySection(ws As Worksheet)
    On Error Resume Next
    
    ' KPI���x��
    ws.Range("A3").Value = "�����i��:"
    ws.Range("A3").Font.Bold = True
    
    ws.Range("C3").Value = "�ݏo��:"
    ws.Range("C3").Font.Bold = True
    
    ws.Range("E3").Value = "��������:"
    ws.Range("E3").Font.Bold = True
    
    ws.Range("G3").Value = "���p�\:"
    ws.Range("G3").Font.Bold = True
    
    ' KPI�l�Z���i�v�Z���ʂ�����j
    With ws.Range("B3")
        .Interior.Color = COLOR_SUCCESS
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    With ws.Range("D3")
        .Interior.Color = COLOR_WARNING
        .Font.Color = vbBlack
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    With ws.Range("F3")
        .Interior.Color = COLOR_OVERDUE
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    With ws.Range("H3")
        .Interior.Color = COLOR_SUCCESS
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    On Error GoTo 0
End Sub

' �A�N�V�����{�^���쐬�i�����֐��j
Private Sub CreateActionButtons(ws As Worksheet)
    On Error GoTo ErrHandler
    
    ' �_�b�V���{�[�h�X�V�{�^��
    Dim btnUpdate As Button
    Set btnUpdate = ws.Buttons.Add(ws.Range("J3").Left, ws.Range("J3").Top, 80, 25)
    btnUpdate.Caption = "�X�V"
    btnUpdate.OnAction = "modDashboard.UpdateDashboard"
    
    ' �ݏo�o�^�{�^��
    Dim btnLend As Button
    Set btnLend = ws.Buttons.Add(ws.Range("A5").Left, ws.Range("A5").Top, 100, 25)
    btnLend.Caption = "�ݏo�o�^"
    btnLend.OnAction = "modLending.RegisterLending"
    
    ' �ԋp�o�^�{�^��
    Dim btnReturn As Button
    Set btnReturn = ws.Buttons.Add(ws.Range("C5").Left, ws.Range("C5").Top, 100, 25)
    btnReturn.Caption = "�ԋp�o�^"
    btnReturn.OnAction = "modLending.RegisterReturn"
    
    ' ���͉�ʕ\���{�^��
    Dim btnInput As Button
    Set btnInput = ws.Buttons.Add(ws.Range("E5").Left, ws.Range("E5").Top, 100, 25)
    btnInput.Caption = "���͉��"
    btnInput.OnAction = "modUI.ShowInputSheet"
    
    ' �e�X�g�f�[�^�쐬�{�^��
    Dim btnTest As Button
    Set btnTest = ws.Buttons.Add(ws.Range("G5").Left, ws.Range("G5").Top, 120, 25)
    btnTest.Caption = "�e�X�g�f�[�^�쐬"
    btnTest.OnAction = "modTestData.CreateAllTestData"
    
    Exit Sub
    
ErrHandler:
    Call LogError("CreateActionButtons", Err.Number, Err.Description)
End Sub

' ���̓V�[�g���C�A�E�g�쐬
Public Sub CreateInputLayout()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = GetWorksheet(SHEET_INPUT)
    If ws Is Nothing Then
        Call LogError("CreateInputLayout", 9, "Input sheet not found")
        Exit Sub
    End If
    
    ' �����̓��e���N���A
    ws.Cells.Clear
    
    ' �^�C�g��
    With ws.Range("A1:E1")
        .Merge
        .Value = "���i�ݏo�E�ԋp���̓t�H�[��"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 30
    End With
    
    ' ���͍��ڃ��x��
    ws.Range("A3").Value = "���iID:"
    ws.Range("A4").Value = "�ؗp��:"
    ws.Range("A5").Value = "�ݏo��:"
    ws.Range("A6").Value = "�ݏo���ԁi���j:"
    ws.Range("A7").Value = "�ԋp��:"
    
    ' ���x���̃t�H�[�}�b�g
    With ws.Range("A3:A7")
        .Font.Bold = True
        .VerticalAlignment = xlCenter
    End With
    
    ' ���̓Z���̃t�H�[�}�b�g
    With ws.Range("B3:B7")
        .Interior.Color = RGB(255, 255, 204) ' �������F
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
    End With
    
    ' �����e�L�X�g
    ws.Range("D3").Value = "��: 1001"
    ws.Range("D4").Value = "��: �c�����Y"
    ws.Range("D5").Value = "��: 2024/1/15 (��=����)"
    ws.Range("D6").Value = "��: 7 (��=7��)"
    ws.Range("D7").Value = "��: 2024/1/22 (�ԋp���̂�)"
    
    With ws.Range("D3:D7")
        .Font.Color = RGB(128, 128, 128)
        .Font.Italic = True
    End With
    
    ' �ݏo�E�ԋp����
    ws.Range("A9").Value = "�� �ݏo�o�^�菇:"
    ws.Range("A10").Value = "1. ���iID�A�ؗp�ҁA�ݏo���A�ݏo���Ԃ����"
    ws.Range("A11").Value = "2. �_�b�V���{�[�h�́u�ݏo�o�^�v�{�^�����N���b�N"
    
    ws.Range("A13").Value = "�� �ԋp�o�^�菇:"
    ws.Range("A14").Value = "1. ���iID�A�ؗp�ҁA�ԋp�������"
    ws.Range("A15").Value = "2. �_�b�V���{�[�h�́u�ԋp�o�^�v�{�^�����N���b�N"
    
    With ws.Range("A9,A13")
        .Font.Bold = True
        .Font.Color = RGB(68, 114, 196)
    End With
    
    ' �{�^���쐬
    Call CreateInputButtons(ws)
    
    Exit Sub
    
ErrHandler:
    Call LogError("CreateInputLayout", Err.Number, Err.Description)
End Sub

' ���̓V�[�g�{�^���쐬�i�����֐��j
Private Sub CreateInputButtons(ws As Worksheet)
    On Error GoTo ErrHandler
    
    ' �_�b�V���{�[�h�ɖ߂�{�^��
    Dim btnDashboard As Button
    Set btnDashboard = ws.Buttons.Add(ws.Range("A17").Left, ws.Range("A17").Top, 120, 25)
    btnDashboard.Caption = "�_�b�V���{�[�h��"
    btnDashboard.OnAction = "modUI.ShowDashboard"
    
    ' ���̓N���A�{�^��
    Dim btnClear As Button
    Set btnClear = ws.Buttons.Add(ws.Range("C17").Left, ws.Range("C17").Top, 100, 25)
    btnClear.Caption = "���̓N���A"
    btnClear.OnAction = "modUI.ClearInputForm"
    
    Exit Sub
    
ErrHandler:
    Call LogError("CreateInputButtons", Err.Number, Err.Description)
End Sub

' ���i�}�X�^�V�[�g���C�A�E�g�쐬
Public Sub CreateItemsLayout()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = GetWorksheet(SHEET_ITEMS)
    If ws Is Nothing Then
        Call LogError("CreateItemsLayout", 9, "Items sheet not found")
        Exit Sub
    End If
    
    ' �����̓��e���N���A
    ws.Cells.Clear
    
    ' �^�C�g��
    With ws.Range("A1:E1")
        .Merge
        .Value = "���i�}�X�^"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 30
    End With
    
    ' �e�[�u���쐬
    Call CreateItemsTable(ws)
    
    Exit Sub
    
ErrHandler:
    Call LogError("CreateItemsLayout", Err.Number, Err.Description)
End Sub

' ���i�e�[�u���쐬�i�����֐��j
Private Sub CreateItemsTable(ws As Worksheet)
    On Error GoTo ErrHandler
    
    ' �w�b�_�[�쐬
    ws.Range("A3").Value = COL_ITEM_ID
    ws.Range("B3").Value = COL_ITEM_NAME
    ws.Range("C3").Value = COL_CATEGORY
    ws.Range("D3").Value = COL_LOCATION
    ws.Range("E3").Value = COL_QUANTITY
    
    ' �e�[�u����
    Dim rng As Range
    Set rng = ws.Range("A3:E3")
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.Name = TABLE_ITEMS
    
    ' �t�H�[�}�b�g�K�p
    Call ApplyStandardTableFormat(tbl)
    
    Exit Sub
    
ErrHandler:
    Call LogError("CreateItemsTable", Err.Number, Err.Description)
End Sub

' �ݏo�����V�[�g���C�A�E�g�쐬
Public Sub CreateLendingLayout()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = GetWorksheet(SHEET_LENDING)
    If ws Is Nothing Then
        Call LogError("CreateLendingLayout", 9, "Lending sheet not found")
        Exit Sub
    End If
    
    ' �����̓��e���N���A
    ws.Cells.Clear
    
    ' �^�C�g��
    With ws.Range("A1:I1")
        .Merge
        .Value = "�ݏo�E�ԋp����"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 30
    End With
    
    ' �e�[�u���쐬
    Call CreateLendingTable(ws)
    
    Exit Sub
    
ErrHandler:
    Call LogError("CreateLendingLayout", Err.Number, Err.Description)
End Sub

' �ݏo�e�[�u���쐬�i�����֐��j
Private Sub CreateLendingTable(ws As Worksheet)
    On Error GoTo ErrHandler
    
    ' �w�b�_�[�쐬
    ws.Range("A3").Value = COL_RECORD_ID
    ws.Range("B3").Value = COL_LENDING_ITEM_ID
    ws.Range("C3").Value = COL_LENDING_ITEM_NAME
    ws.Range("D3").Value = COL_BORROWER
    ws.Range("E3").Value = COL_LEND_DATE
    ws.Range("F3").Value = COL_DUE_DATE
    ws.Range("G3").Value = COL_RETURN_DATE
    ws.Range("H3").Value = COL_STATUS
    ws.Range("I3").Value = COL_REMARKS
    
    ' �e�[�u����
    Dim rng As Range
    Set rng = ws.Range("A3:I3")
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.Name = TABLE_LENDING
    
    ' �t�H�[�}�b�g�K�p
    Call ApplyStandardTableFormat(tbl)
    
    Exit Sub
    
ErrHandler:
    Call LogError("CreateLendingTable", Err.Number, Err.Description)
End Sub

' �V�[�g�\���؂�ւ��֐��Q
Public Sub ShowDashboard()
    On Error Resume Next
    ThisWorkbook.Worksheets(SHEET_DASHBOARD).Activate
    On Error GoTo 0
End Sub

Public Sub ShowInputSheet()
    On Error Resume Next
    ThisWorkbook.Worksheets(SHEET_INPUT).Activate
    On Error GoTo 0
End Sub

Public Sub ShowItemsSheet()
    On Error Resume Next
    ThisWorkbook.Worksheets(SHEET_ITEMS).Activate
    On Error GoTo 0
End Sub

Public Sub ShowLendingSheet()
    On Error Resume Next
    ThisWorkbook.Worksheets(SHEET_LENDING).Activate
    On Error GoTo 0
End Sub

' ���̓t�H�[���N���A
Public Sub ClearInputForm()
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = GetWorksheet(SHEET_INPUT)
    If Not ws Is Nothing Then
        ws.Range(INPUT_ITEM_ID).Value = ""
        ws.Range(INPUT_BORROWER).Value = ""
        ws.Range(INPUT_LEND_DATE).Value = ""
        ws.Range(INPUT_LENDING_DAYS).Value = ""
        ws.Range(INPUT_RETURN_DATE).Value = ""
        MsgBox "���̓t�H�[�����N���A���܂����B", vbInformation
    End If
    
    On Error GoTo 0
End Sub

' �S�V�[�g���C�A�E�g������
Public Sub InitializeAllLayouts()
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    
    Call CreateDashboardLayout
    Call CreateInputLayout
    Call CreateItemsLayout
    Call CreateLendingLayout
    
    ' �_�b�V���{�[�h��\��
    Call ShowDashboard
    
    Application.ScreenUpdating = True
    MsgBox "�S�V�[�g�̃��C�A�E�g�����������܂����B", vbInformation
    
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Call LogError("InitializeAllLayouts", Err.Number, Err.Description)
    MsgBox "���C�A�E�g���������ɃG���[���������܂���: " & Err.Description, vbCritical
End Sub