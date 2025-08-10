Attribute VB_Name = "common"

' ===============================================
' �]�ƈ��Ǘ��V�X�e�� - ���ʃ��W���[��
' ===============================================

Option Explicit

' �]�ƈ��f�[�^�̗�ԍ��萔
Public Const COL_EMPLOYEE_ID As Integer = 1
Public Const COL_NAME As Integer = 2
Public Const COL_DEPARTMENT As Integer = 3
Public Const COL_POSITION As Integer = 4
Public Const COL_HIRE_DATE As Integer = 5
Public Const COL_SALARY As Integer = 6
Public Const COL_PHONE As Integer = 7
Public Const COL_EMAIL As Integer = 8

' �f�[�^�s�̊J�n�ʒu�i�w�b�_�[�s�̎��j
Public Const DATA_START_ROW As Integer = 2

' ===============================================
' �V�X�e���������֘A
' ===============================================

Public Sub InitializeEmployeeManagementSystem()
    ' �]�ƈ��Ǘ��V�X�e���̏�����
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    
    Call Xdebug.printx("�]�ƈ��Ǘ��V�X�e������������...")
    
    ' �w�b�_�[�̐ݒ�
    Call SetupEmployeeHeaders(ws)
    
    ' �����f�[�^���Ȃ��ꍇ�̂݃T���v���f�[�^���쐬
    If IsEmployeeDataEmpty(ws) Then
        Call CreateSampleEmployeeData(ws)
        Call Xdebug.printx("�T���v���]�ƈ��f�[�^���쐬���܂���")
    Else
        Call Xdebug.printx("�����̏]�ƈ��f�[�^��������܂���")
    End If
    
    ' UI�v�f�̐ݒ�
    Call SetupEmployeeUI(ws)
    
    Call Xdebug.printx("�]�ƈ��Ǘ��V�X�e���̏��������������܂���")
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("InitializeEmployeeManagementSystem", Err.Description)
End Sub

Public Sub SetupEmployeeHeaders(ws As Worksheet)
    ' �w�b�_�[�s�̐ݒ�
    On Error GoTo ErrorHandler
    
    With ws
        .Cells(1, COL_EMPLOYEE_ID).Value = "�]�ƈ�ID"
        .Cells(1, COL_NAME).Value = "����"
        .Cells(1, COL_DEPARTMENT).Value = "����"
        .Cells(1, COL_POSITION).Value = "��E"
        .Cells(1, COL_HIRE_DATE).Value = "���Г�"
        .Cells(1, COL_SALARY).Value = "���^"
        .Cells(1, COL_PHONE).Value = "�d�b�ԍ�"
        .Cells(1, COL_EMAIL).Value = "���[���A�h���X"
        
        ' �w�b�_�[�s�̃X�^�C���ݒ�
        With .Range(.Cells(1, 1), .Cells(1, COL_EMAIL))
            .Font.Bold = True
            .Interior.Color = RGB(200, 220, 240)
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
        End With
        
        ' �񕝂̎�������
        .Columns("A:H").AutoFit
    End With
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("SetupEmployeeHeaders", Err.Description)
End Sub

Public Function IsEmployeeDataEmpty(ws As Worksheet) As Boolean
    ' �]�ƈ��f�[�^���󂩂ǂ������`�F�b�N
    On Error GoTo ErrorHandler
    
    IsEmployeeDataEmpty = (ws.Cells(DATA_START_ROW, COL_EMPLOYEE_ID).Value = "")
    Exit Function
    
ErrorHandler:
    Call Xdebug.printError("IsEmployeeDataEmpty", Err.Description)
    IsEmployeeDataEmpty = True
End Function

' ===============================================
' �T���v���f�[�^����
' ===============================================

Public Sub CreateSampleEmployeeData(ws As Worksheet)
    ' 30���̃��A���ȏ]�ƈ��T���v���f�[�^���쐬
    On Error GoTo ErrorHandler
    
    ' �T���v���]�ƈ��f�[�^���ʂɒǉ���������ɕύX
    Call AddEmployeeRecord(ws, "EMP001", "�R�c ���Y", "�c�ƕ�", "����", "2015-04-01", 8500000, "03-1234-5678", "yamada.taro@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP002", "���� �Ԏq", "�l����", "�ے�", "2016-10-15", 7200000, "03-1234-5679", "sato.hanako@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP003", "�c�� ���Y", "�Z�p��", "��C", "2017-08-20", 6500000, "03-1234-5680", "tanaka.jiro@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP004", "���� ����", "�o����", "�W��", "2018-03-10", 5800000, "03-1234-5681", "takahashi.misaki@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP005", "�ɓ� �a��", "�c�ƕ�", "��C", "2018-07-01", 6200000, "03-1234-5682", "ito.kazuya@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP006", "�n�� ����", "������", "�ے�", "2014-09-15", 7000000, "03-1234-5683", "watanabe.satomi@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP007", "���� ����", "�Z�p��", "�G���W�j�A", "2019-04-01", 5500000, "03-1234-5684", "kobayashi.kenta@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP008", "���� �R��", "�}�[�P�e�B���O��", "�X�y�V�����X�g", "2020-01-20", 6800000, "03-1234-5685", "kato.yumi@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP009", "�g�c �Y��", "�c�ƕ�", "�W��", "2017-12-01", 6000000, "03-1234-5686", "yoshida.yuichi@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP010", "���� ����", "�l����", "�A�V�X�^���g", "2021-03-15", 4800000, "03-1234-5687", "nakamura.mai@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP011", "�� ���", "�Z�p��", "�V�j�A�G���W�j�A", "2016-05-10", 7500000, "03-1234-5688", "hayashi.daisuke@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP012", "�X �q�q", "�o����", "��C", "2019-08-25", 5700000, "03-1234-5689", "mori.tomoko@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP013", "�r�c �T��", "�c�ƕ�", "�}�l�[�W���[", "2013-11-01", 8200000, "03-1234-5690", "ikeda.shinichi@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP014", "���{ �b�q", "������", "�A�V�X�^���g", "2022-04-01", 4500000, "03-1234-5691", "hashimoto.keiko@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP015", "�ΐ� �_��", "�Z�p��", "�e�N�j�J�����[�h", "2015-07-15", 8000000, "03-1234-5692", "ishikawa.koji@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP016", "�O�c ����", "�}�[�P�e�B���O��", "�A�i���X�g", "2020-09-01", 5900000, "03-1234-5693", "maeda.yuka@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP017", "���c ����", "�c�ƕ�", "�c�ƒS��", "2021-06-10", 5200000, "03-1234-5694", "okada.tatsuya@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP018", "���� ���b", "�l����", "�̗p�S��", "2018-12-01", 6300000, "03-1234-5695", "murakami.rie@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP019", "���� ��C", "�Z�p��", "�G���W�j�A", "2022-01-15", 5400000, "03-1234-5696", "shimizu.takumi@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP020", "�R�� ��", "�o����", "�o���S��", "2019-11-20", 5600000, "03-1234-5697", "yamaguchi.ai@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP021", "���{ ����", "�c�ƕ�", "�V�j�A�Z�[���X", "2016-02-28", 7300000, "03-1234-5698", "matsumoto.masaki@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP022", "��� ��t", "������", "�����S��", "2020-05-15", 5100000, "03-1234-5699", "inoue.chiharu@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP023", "�ؑ� ��", "�Z�p��", "�v���W�F�N�g�}�l�[�W���[", "2014-08-01", 8800000, "03-1234-5700", "kimura.ken@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP024", "�ē� ����", "�}�[�P�e�B���O��", "�}�l�[�W���[", "2017-03-20", 7800000, "03-1234-5701", "saito.miho@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP025", "���� ���V", "�c�ƕ�", "�c�ƒS��", "2021-09-01", 5300000, "03-1234-5702", "nakajima.hiroyuki@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP026", "���c ���₩", "�l����", "�l���S��", "2019-06-10", 5800000, "03-1234-5703", "harada.sayaka@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP027", "��{ ����", "�Z�p��", "�V�X�e���A�[�L�e�N�g", "2015-12-01", 9200000, "03-1234-5704", "sakamoto.naoki@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP028", "�� �D��", "�o����", "�����A�i���X�g", "2020-08-15", 6700000, "03-1234-5705", "aoki.yuka@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP029", "���c ���j", "�c�ƕ�", "�c�ƒS��", "2022-02-01", 5000000, "03-1234-5706", "fujita.mitsuo@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP030", "���� ������", "������", "�鏑", "2021-11-15", 4700000, "03-1234-5707", "nishimura.akane@company.co.jp")
    
    ' �f�[�^�͈͂̏����ݒ�
    Call FormatEmployeeDataRange(ws)
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("CreateSampleEmployeeData", Err.Description)
End Sub

Public Sub AddEmployeeRecord(ws As Worksheet, empID As String, empName As String, dept As String, pos As String, hireDate As String, salary As Long, phone As String, email As String)
    ' �]�ƈ����R�[�h��ǉ�����w���p�[�֐�
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_EMPLOYEE_ID).End(xlUp).Row + 1
    
    ws.Cells(lastRow, COL_EMPLOYEE_ID).Value = empID
    ws.Cells(lastRow, COL_NAME).Value = empName
    ws.Cells(lastRow, COL_DEPARTMENT).Value = dept
    ws.Cells(lastRow, COL_POSITION).Value = pos
    ws.Cells(lastRow, COL_HIRE_DATE).Value = CDate(hireDate)
    ws.Cells(lastRow, COL_SALARY).Value = salary
    ws.Cells(lastRow, COL_PHONE).Value = phone
    ws.Cells(lastRow, COL_EMAIL).Value = email
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("AddEmployeeRecord", Err.Description)
End Sub

Public Sub FormatEmployeeDataRange(ws As Worksheet)
    ' �f�[�^�͈͂̏����ݒ�
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    Dim i As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_EMPLOYEE_ID).End(xlUp).row
    
    With ws.Range(ws.Cells(DATA_START_ROW, 1), ws.Cells(lastRow, COL_EMAIL))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        
        ' ���^��̏����ݒ�i�J���}��؂�j
        ws.Range(ws.Cells(DATA_START_ROW, COL_SALARY), ws.Cells(lastRow, COL_SALARY)).NumberFormat = "#,##0"
        
        ' ���Г���̏����ݒ�
        ws.Range(ws.Cells(DATA_START_ROW, COL_HIRE_DATE), ws.Cells(lastRow, COL_HIRE_DATE)).NumberFormat = "yyyy/mm/dd"
        
        ' ���݂̍s�̐F�t��
        For i = DATA_START_ROW To lastRow
            If i Mod 2 = 0 Then
                ws.Range(ws.Cells(i, 1), ws.Cells(i, COL_EMAIL)).Interior.Color = RGB(245, 245, 245)
            End If
        Next i
    End With
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("FormatEmployeeDataRange", Err.Description)
End Sub

' ===============================================
' UI�v�f�ݒ�
' ===============================================

Public Sub SetupEmployeeUI(ws As Worksheet)
    ' �]�ƈ��Ǘ��p��UI�v�f�i�{�^���Ȃǁj��ݒ�
    On Error GoTo ErrorHandler
    
    Call Xdebug.printx("�]�ƈ��Ǘ�UI��ݒ蒆...")
    
    ' �����̃{�^�����폜
    Call ClearExistingButtons(ws)
    
    ' �V�K�]�ƈ��ǉ��{�^��
    Call CreateButton(ws, "�V�K�ǉ�", 10, 10, 80, 25, "common.AddNewEmployee_Click")
    
    ' �폜�{�^��
    Call CreateButton(ws, "�I���s�폜", 100, 10, 80, 25, "common.DeleteSelectedEmployee_Click")
    
    ' �����{�^��
    Call CreateButton(ws, "����", 190, 10, 60, 25, "common.SearchEmployee_Click")
    
    ' ���Z�b�g�{�^��
    Call CreateButton(ws, "���Z�b�g", 260, 10, 60, 25, "common.ResetView_Click")
    
    ' �����{�b�N�X
    Call CreateSearchBox(ws)
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("SetupEmployeeUI", Err.Description)
End Sub

Public Sub ClearExistingButtons(ws As Worksheet)
    ' �����̃{�^�����폜
    On Error Resume Next
    Dim obj As Object
    For Each obj In ws.Shapes
        If obj.Type = 1 Then ' msoFormControl
            obj.Delete
        End If
    Next obj
    On Error GoTo 0
End Sub

Public Sub CreateButton(ws As Worksheet, buttonText As String, left As Double, top As Double, width As Double, height As Double, macroName As String)
    ' �{�^�����쐬
    On Error GoTo ErrorHandler
    
    Dim btn As Object
    Set btn = ws.Shapes.AddFormControl(xlButtonControl, left, top, width, height)
    btn.TextFrame.Characters.Text = buttonText
    btn.OnAction = macroName
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("CreateButton", Err.Description)
End Sub

Public Sub CreateSearchBox(ws As Worksheet)
    ' �����p�e�L�X�g�{�b�N�X���쐬
    On Error GoTo ErrorHandler
    
    ' �������x��
    ws.Cells(1, 10).Value = "����:"
    ws.Cells(1, 10).Font.Bold = True
    
    ' �����{�b�N�X�i�Z���j
    ws.Cells(1, 11).Value = ""
    ws.Cells(1, 11).Interior.Color = RGB(255, 255, 200)
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("CreateSearchBox", Err.Description)
End Sub

' ===============================================
' �f�[�^����֐��iCRUD�j
' ===============================================

Public Sub AddNewEmployee_Click()
    ' �V�K�]�ƈ��ǉ�
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_EMPLOYEE_ID).End(xlUp).row + 1
    
    ' �V�K�]�ƈ�ID�𐶐�
    Dim newEmployeeID As String
    newEmployeeID = "EMP" & Format(lastRow - 1, "000")
    
    ' �V�����s�Ƀf�t�H���g�l��ݒ�
    ws.Cells(lastRow, COL_EMPLOYEE_ID).Value = newEmployeeID
    ws.Cells(lastRow, COL_NAME).Value = "�V�K�]�ƈ�"
    ws.Cells(lastRow, COL_DEPARTMENT).Value = "���ݒ�"
    ws.Cells(lastRow, COL_POSITION).Value = "���ݒ�"
    ws.Cells(lastRow, COL_HIRE_DATE).Value = Date
    ws.Cells(lastRow, COL_SALARY).Value = 0
    ws.Cells(lastRow, COL_PHONE).Value = ""
    ws.Cells(lastRow, COL_EMAIL).Value = ""
    
    ' �V�����s��I��
    ws.Cells(lastRow, COL_NAME).Select
    
    Call Xdebug.printx("�V�K�]�ƈ���ǉ����܂���: " & newEmployeeID)
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("AddNewEmployee_Click", Err.Description)
End Sub

Public Sub DeleteSelectedEmployee_Click()
    ' �I�����ꂽ�]�ƈ����폜
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim selectedRow As Long
    selectedRow = Selection.row
    
    ' �w�b�_�[�s���s�̍폜��h��
    If selectedRow < DATA_START_ROW Then
        MsgBox "�w�b�_�[�s�͍폜�ł��܂���B", vbExclamation
        Exit Sub
    End If
    
    If ws.Cells(selectedRow, COL_EMPLOYEE_ID).Value = "" Then
        MsgBox "�폜����]�ƈ����I������Ă��܂���B", vbExclamation
        Exit Sub
    End If
    
    ' �m�F���b�Z�[�W
    Dim employeeName As String
    employeeName = ws.Cells(selectedRow, COL_NAME).Value
    
    If MsgBox("�]�ƈ��u" & employeeName & "�v���폜���܂����H", vbYesNo + vbQuestion) = vbYes Then
        ws.Rows(selectedRow).Delete
        Call Xdebug.printx("�]�ƈ����폜���܂���: " & employeeName)
        MsgBox "�]�ƈ����폜���܂����B", vbInformation
    End If
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("DeleteSelectedEmployee_Click", Err.Description)
End Sub

' ===============================================
' �����E�t�B���^�@�\
' ===============================================

Public Sub SearchEmployee_Click()
    ' �]�ƈ������@�\
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim searchTerm As String
    searchTerm = ws.Cells(1, 11).Value ' �����{�b�N�X�̒l���擾
    
    If searchTerm = "" Then
        MsgBox "�����L�[���[�h����͂��Ă��������B", vbExclamation
        Exit Sub
    End If
    
    Call FilterEmployeeData(ws, searchTerm)
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("SearchEmployee_Click", Err.Description)
End Sub

Public Sub FilterEmployeeData(ws As Worksheet, searchTerm As String)
    ' �f�[�^���t�B���^�����O
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_EMPLOYEE_ID).End(xlUp).row
    
    ' �I�[�g�t�B���^��K�p
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
    
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, COL_EMAIL)).AutoFilter
    
    ' ������ł̌����i���O�A�����A��E�j
    Dim foundMatch As Boolean
    foundMatch = False
    
    ' ���O�ł̌���
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, COL_EMAIL)).AutoFilter Field:=COL_NAME, Criteria1:="*" & searchTerm & "*"
    If CountVisibleRows(ws) > 1 Then
        foundMatch = True
    Else
        ' �����ł̌���
        ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, COL_EMAIL)).AutoFilter Field:=COL_NAME
        ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, COL_EMAIL)).AutoFilter Field:=COL_DEPARTMENT, Criteria1:="*" & searchTerm & "*"
        If CountVisibleRows(ws) > 1 Then
            foundMatch = True
        Else
            ' ��E�ł̌���
            ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, COL_EMAIL)).AutoFilter Field:=COL_DEPARTMENT
            ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, COL_EMAIL)).AutoFilter Field:=COL_POSITION, Criteria1:="*" & searchTerm & "*"
            If CountVisibleRows(ws) > 1 Then
                foundMatch = True
            End If
        End If
    End If
    
    If Not foundMatch Then
        MsgBox "�u" & searchTerm & "�v�Ɉ�v����]�ƈ���������܂���ł����B", vbInformation
    End If
    
    Call Xdebug.printx("�������s: " & searchTerm)
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("FilterEmployeeData", Err.Description)
End Sub

Public Function CountVisibleRows(ws As Worksheet) As Long
    ' ���s�����J�E���g
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    Dim visibleCount As Long
    Dim i As Long
    
    lastRow = ws.Cells(ws.Rows.Count, COL_EMPLOYEE_ID).End(xlUp).row
    visibleCount = 0
    
    For i = 1 To lastRow
        If Not ws.Rows(i).Hidden Then
            visibleCount = visibleCount + 1
        End If
    Next i
    
    CountVisibleRows = visibleCount
    Exit Function
    
ErrorHandler:
    Call Xdebug.printError("CountVisibleRows", Err.Description)
    CountVisibleRows = 0
End Function

Public Sub ResetView_Click()
    ' �r���[�����Z�b�g�i�t�B���^�������j
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' �t�B���^������
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
    
    ' �����{�b�N�X���N���A
    ws.Cells(1, 11).Value = ""
    
    Call Xdebug.printx("�r���[�����Z�b�g���܂���")
    MsgBox "�t�B���^���������܂����B", vbInformation
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("ResetView_Click", Err.Description)
End Sub

' ===============================================
' �o���f�[�V�����֐�
' ===============================================

Public Function ValidateEmployeeData(ws As Worksheet, row As Long) As Boolean
    ' �]�ƈ��f�[�^�̃o���f�[�V����
    On Error GoTo ErrorHandler
    
    Dim isValid As Boolean
    isValid = True
    Dim errorMsg As String
    errorMsg = ""
    
    ' �]�ƈ�ID�̌���
    If ws.Cells(row, COL_EMPLOYEE_ID).Value = "" Then
        isValid = False
        errorMsg = errorMsg & "�]�ƈ�ID�����͂���Ă��܂���B" & vbCrLf
    End If
    
    ' �����̌���
    If ws.Cells(row, COL_NAME).Value = "" Then
        isValid = False
        errorMsg = errorMsg & "���������͂���Ă��܂���B" & vbCrLf
    End If
    
    ' �����̌���
    If ws.Cells(row, COL_DEPARTMENT).Value = "" Then
        isValid = False
        errorMsg = errorMsg & "���������͂���Ă��܂���B" & vbCrLf
    End If
    
    ' ���^�̌���
    If Not IsNumeric(ws.Cells(row, COL_SALARY).Value) Or ws.Cells(row, COL_SALARY).Value < 0 Then
        isValid = False
        errorMsg = errorMsg & "���^�͐��̐��l�œ��͂��Ă��������B" & vbCrLf
    End If
    
    ' ���[���A�h���X�̊ȒP�Ȍ���
    Dim email As String
    email = ws.Cells(row, COL_EMAIL).Value
    If email <> "" And InStr(email, "@") = 0 Then
        isValid = False
        errorMsg = errorMsg & "���[���A�h���X�̌`��������������܂���B" & vbCrLf
    End If
    
    If Not isValid Then
        MsgBox "�f�[�^�̌��؂Ɏ��s���܂���:" & vbCrLf & errorMsg, vbExclamation
    End If
    
    ValidateEmployeeData = isValid
    Exit Function
    
ErrorHandler:
    Call Xdebug.printError("ValidateEmployeeData", Err.Description)
    ValidateEmployeeData = False
End Function
