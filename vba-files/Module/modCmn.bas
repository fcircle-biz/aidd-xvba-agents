Attribute VB_Name = "modCmn"
'=============================================================================
' modCmn.bas - ���ʃ��[�e�B���e�B�E�ėp�@�\
'=============================================================================
' �T�v:
'   �Ɩ��V�X�e�����ʂŎg�p����ėp�@�\���W��
'   �f�[�^�A�N�Z�X�A�����񏈗��A���O�A�t�H�[�}�b�g�A���؂ȂǍė��p�\�ȋ@�\�Q
'=============================================================================
Option Explicit

'=============================================================================
' �Ɨ�����̂��߂̒萔��`
'=============================================================================
' �G���[���b�Z�[�W�萔
Private Const ERR_SHEET_NOT_FOUND As String = "�V�[�g��������܂���: "
Private Const ERR_TABLE_NOT_FOUND As String = "�e�[�u����������܂���: "
Private Const ERR_COLUMN_NOT_FOUND As String = "�񂪌�����܂���: "

' �t�H���g�ݒ�萔
Private Const FONT_NAME As String = "Yu Gothic UI"
Private Const FONT_SIZE_NORMAL As Integer = 10
Private Const FONT_SIZE_HEADER As Integer = 12
Private Const FONT_SIZE_BUTTON As Integer = 11
Private Const FONT_COLOR_NORMAL As Long = 0
Private Const FONT_COLOR_HEADER As Long = 16777215

' �F�ݒ�萔
Private Const BG_COLOR_HEADER As Long = 5287936
Private Const BG_COLOR_ALTERNATE As Long = 15921906
Private Const BORDER_COLOR_DEFAULT As Long = 8421504

' ���K�\���p�^�[���萔
Private Const REGEX_EMAIL As String = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
Private Const REGEX_PHONE As String = "^[0-9]{2,4}-[0-9]{3,4}-[0-9]{4}$"
Private Const REGEX_ZIP As String = "^\d{3}-\d{4}$"
Private Const REGEX_CUSTOMERID As String = "^[A-Za-z0-9]{3,20}$"

' �f�t�H���g�ݒ�l�萔
Private Const DEFAULT_LOG_DIR As String = "C:\git\xvba-mock-creator\logs\"

'=============================================================================
' �f�[�^�A�N�Z�X�E����n�֐�
'=============================================================================

' ���S�ȃ��[�N�V�[�g�擾�i�V�[�g���j
Public Function GetWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheet = ThisWorkbook.Worksheets(sheetName)
    If Err.Number <> 0 Then
        Call LogError("GetWorksheet", ERR_SHEET_NOT_FOUND & sheetName)
        Set GetWorksheet = Nothing
    End If
End Function

' ���S�ȃ��[�N�V�[�g�擾�i�C���f�b�N�X�ԍ��j
Public Function GetWorksheetByIndex(ByVal sheetIndex As Integer) As Worksheet
    On Error GoTo ErrHandler
    
    If sheetIndex <= 0 Or sheetIndex > ThisWorkbook.Worksheets.Count Then
        Call LogError("GetWorksheetByIndex", "�����ȃV�[�g�C���f�b�N�X: " & sheetIndex)
        Set GetWorksheetByIndex = Nothing
        Exit Function
    End If
    
    Set GetWorksheetByIndex = ThisWorkbook.Worksheets(sheetIndex)
    Exit Function
    
ErrHandler:
    Call LogError("GetWorksheetByIndex", "�V�[�g�C���f�b�N�X " & sheetIndex & " �ŃG���[: " & Err.Description)
    Set GetWorksheetByIndex = Nothing
End Function

' ���S�ȃe�[�u���擾
Public Function GetTable(ByVal ws As Worksheet, ByVal tableName As String) As ListObject
    On Error GoTo ErrHandler
    
    ' �e�[�u�������݂��邩�`�F�b�N
    If Not TableExists(ws, tableName) Then
        Call LogError("GetTable", ERR_TABLE_NOT_FOUND & tableName & " (�V�[�g: " & ws.Name & ")")
        Set GetTable = Nothing
        Exit Function
    End If
    
    Set GetTable = ws.ListObjects(tableName)
    Exit Function
    
ErrHandler:
    Call LogError("GetTable", "�e�[�u���擾�G���[: " & tableName & " - " & Err.Description)
    Set GetTable = Nothing
End Function

' �e�[�u�����݃`�F�b�N
Public Function TableExists(ByVal ws As Worksheet, ByVal tableName As String) As Boolean
    On Error Resume Next
    Dim i As Integer
    
    ' �����`�F�b�N
    If ws Is Nothing Or Len(tableName) = 0 Then
        TableExists = False
        Exit Function
    End If
    
    ' ListObjects�R���N�V���������[�v���Ċm�F
    For i = 1 To ws.ListObjects.Count
        If ws.ListObjects(i).Name = tableName Then
            TableExists = True
            Exit Function
        End If
    Next i
    
    TableExists = False
End Function

' ��C���f�b�N�X�擾�i0�`�F�b�N�K�{�j
Public Function GetColumnIndex(ByVal tbl As ListObject, ByVal columnName As String) As Integer
    On Error GoTo ErrHandler
    
    Dim col As ListColumn
    For Each col In tbl.ListColumns
        If col.Name = columnName Then
            GetColumnIndex = col.Index
            Exit Function
        End If
    Next col
    
    GetColumnIndex = 0
    Call LogError("GetColumnIndex", ERR_COLUMN_NOT_FOUND & columnName)
    Exit Function
    
ErrHandler:
    GetColumnIndex = 0
    Call LogError("GetColumnIndex", Err.Description & " (Column: " & columnName & ")")
End Function

' ���S�ȃV�[�g�N���A
Public Sub SafeClearSheet(ByVal ws As Worksheet, Optional ByVal keepFormats As Boolean = False)
    On Error GoTo ErrHandler
    
    If ws Is Nothing Then
        Call LogError("SafeClearSheet", "���[�N�V�[�g��Null�ł�")
        Exit Sub
    End If
    
    ' �g�p�͈͂����
    Dim found As Range
    Set found = ws.Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If Not found Is Nothing Then
        With ws.Range("A1", ws.Cells(found.Row, found.Column))
            If keepFormats Then
                .ClearContents
            Else
                .Clear
            End If
        End With
    End If
    Exit Sub
    
ErrHandler:
    Call LogError("SafeClearSheet", Err.Description)
End Sub

'=============================================================================
' �f�[�^�擾�E�ݒ�֐��i�e�[�u���s����j
'=============================================================================

' �e�[�u���s����l���擾
Public Function GetRowValue(ByVal row As ListRow, ByVal columnName As String) As Variant
    On Error GoTo ErrHandler
    
    Dim colIndex As Integer
    colIndex = GetColumnIndex(row.Parent, columnName)
    If colIndex = 0 Then
        GetRowValue = Empty
        Exit Function
    End If
    
    GetRowValue = row.Range.Cells(1, colIndex).Value
    Exit Function
    
ErrHandler:
    GetRowValue = Empty
    Call LogError("GetRowValue", Err.Description & " (Column: " & columnName & ")")
End Function

' �e�[�u���s�ɒl��ݒ�
Public Sub SetRowValue(ByVal row As ListRow, ByVal columnName As String, ByVal value As Variant)
    On Error GoTo ErrHandler
    
    Dim colIndex As Integer
    colIndex = GetColumnIndex(row.Parent, columnName)
    If colIndex = 0 Then Exit Sub
    
    row.Range.Cells(1, colIndex).Value = value
    Exit Sub
    
ErrHandler:
    Call LogError("SetRowValue", Err.Description & " (Column: " & columnName & ")")
End Sub

' �e�[�u���s���當����擾
Public Function GetRowText(ByVal row As ListRow, ByVal columnName As String) As String
    On Error Resume Next
    GetRowText = CStr(GetRowValue(row, columnName))
    If Err.Number <> 0 Then GetRowText = ""
End Function

' �e�[�u���s�ɕ�����ݒ�
Public Sub SetRowText(ByVal row As ListRow, ByVal columnName As String, ByVal text As String)
    Call SetRowValue(row, columnName, text)
End Sub

' �e�[�u���s������t�擾
Public Function GetRowDate(ByVal row As ListRow, ByVal columnName As String) As Date
    On Error Resume Next
    GetRowDate = CDate(GetRowValue(row, columnName))
    If Err.Number <> 0 Then GetRowDate = 0
End Function

' �e�[�u���s�ɓ��t�ݒ�
Public Sub SetRowDate(ByVal row As ListRow, ByVal columnName As String, ByVal dateValue As Date)
    Call SetRowValue(row, columnName, dateValue)
End Sub

'=============================================================================
' �����񏈗����[�e�B���e�B
'=============================================================================

' ������g�����i�S�p�X�y�[�X���Ή��j
Public Function TrimAll(ByVal text As String) As String
    On Error Resume Next
    ' �O��̔��p�E�S�p�X�y�[�X�A�^�u�A���s������
    TrimAll = text
    TrimAll = Replace(TrimAll, vbTab, " ")
    TrimAll = Replace(TrimAll, vbCrLf, " ")
    TrimAll = Replace(TrimAll, vbCr, " ")
    TrimAll = Replace(TrimAll, vbLf, " ")
    TrimAll = Replace(TrimAll, "�@", " ")  ' �S�p�X�y�[�X�����p�X�y�[�X
    
    ' �A���X�y�[�X��P�ꉻ
    Do While InStr(TrimAll, "  ") > 0
        TrimAll = Replace(TrimAll, "  ", " ")
    Loop
    
    TrimAll = Trim(TrimAll)
End Function

' �d�b�ԍ����K��
Public Function NormalizePhone(ByVal phone As String) As String
    On Error Resume Next
    
    NormalizePhone = TrimAll(phone)
    ' �S�p�����E�n�C�t���𔼊p�ɕϊ�
    NormalizePhone = StrConv(NormalizePhone, vbNarrow)
    
    ' �s�v�����폜
    NormalizePhone = Replace(NormalizePhone, "(", "")
    NormalizePhone = Replace(NormalizePhone, ")", "")
    NormalizePhone = Replace(NormalizePhone, " ", "")
    NormalizePhone = Replace(NormalizePhone, "�@", "")
    
    ' �n�C�t���̐��K���i03-1234-5678�`���j
    If Len(NormalizePhone) = 10 Or Len(NormalizePhone) = 11 Then
        ' ��x�n�C�t����S�č폜
        NormalizePhone = Replace(NormalizePhone, "-", "")
        
        ' �K�؂Ȉʒu�Ƀn�C�t����}��
        If Len(NormalizePhone) = 10 Then
            ' 03-XXXX-XXXX �܂��� 06-XXXX-XXXX
            If Left(NormalizePhone, 2) = "03" Or Left(NormalizePhone, 2) = "06" Then
                NormalizePhone = Left(NormalizePhone, 2) & "-" & Mid(NormalizePhone, 3, 4) & "-" & Right(NormalizePhone, 4)
            Else
                ' 0XX-XXX-XXXX
                NormalizePhone = Left(NormalizePhone, 3) & "-" & Mid(NormalizePhone, 4, 3) & "-" & Right(NormalizePhone, 4)
            End If
        ElseIf Len(NormalizePhone) = 11 Then
            ' 090-XXXX-XXXX
            NormalizePhone = Left(NormalizePhone, 3) & "-" & Mid(NormalizePhone, 4, 4) & "-" & Right(NormalizePhone, 4)
        End If
    End If
End Function

' �X�֔ԍ����K��
Public Function NormalizeZip(ByVal zip As String) As String
    On Error Resume Next
    
    NormalizeZip = TrimAll(zip)
    ' �S�p�����E�n�C�t���𔼊p�ɕϊ�
    NormalizeZip = StrConv(NormalizeZip, vbNarrow)
    
    ' �s�v�����폜
    NormalizeZip = Replace(NormalizeZip, " ", "")
    NormalizeZip = Replace(NormalizeZip, "�@", "")
    NormalizeZip = Replace(NormalizeZip, "��", "")
    
    ' �n�C�t�����Ȃ��ꍇ�͒ǉ��i1234567 �� 123-4567�j
    If Len(NormalizeZip) = 7 And InStr(NormalizeZip, "-") = 0 Then
        NormalizeZip = Left(NormalizeZip, 3) & "-" & Right(NormalizeZip, 4)
    End If
End Function

' ���[���A�h���X���K��
Public Function NormalizeEmail(ByVal email As String) As String
    On Error Resume Next
    
    NormalizeEmail = TrimAll(email)
    ' �������ɓ���
    NormalizeEmail = LCase(NormalizeEmail)
    
    ' �S�p�p�����𔼊p�ɕϊ�
    NormalizeEmail = StrConv(NormalizeEmail, vbNarrow)
End Function

' ��������R���N�V�����ɕ���
Public Function SplitToCollection(ByVal str As String, Optional ByVal delimiter As String = ",") As Collection
    On Error Resume Next
    
    Set SplitToCollection = New Collection
    
    If Len(str) = 0 Then Exit Function
    
    Dim parts As Variant
    parts = Split(str, delimiter)
    
    Dim i As Integer
    For i = LBound(parts) To UBound(parts)
        Dim part As String
        part = TrimAll(CStr(parts(i)))
        If part <> "" Then
            SplitToCollection.Add part
        End If
    Next i
End Function

'=============================================================================
' ���O�Ǘ��֐�
'=============================================================================

' �O�����O�t�@�C���o��
Public Sub LogError(ByVal functionName As String, ByVal errorMessage As String)
    On Error Resume Next
    
    Dim logMessage As String
    logMessage = Format(Now, "yyyy-mm-dd hh:nn:ss") & " [ERROR] " & functionName & ": " & errorMessage
    
    ' �f�o�b�O�o�́i�J�����̂݁j
    Debug.Print logMessage
    
    ' �O�����O�t�@�C���ɋL�^�i�G���[�����Ȃ��j
    Call WriteExternalLogSafe(logMessage)
    
    ' �G���[�_�C�A���O��\�����ċ����I��
    MsgBox "�V�X�e���G���[���������܂����B" & vbCrLf & vbCrLf & _
           "�֐�: " & functionName & vbCrLf & _
           "�G���[: " & errorMessage & vbCrLf & vbCrLf & _
           "�ڍׂ̓��O�t�@�C�������m�F���������B" & vbCrLf & _
           "�A�v���P�[�V�������I�����܂��B", vbCritical, "�V�X�e���G���["
    
    ' �����I���i�ۑ��m�F�Ȃ��j
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    ThisWorkbook.Saved = True  ' �ۑ��ς݂Ƃ��ă}�[�N
    Application.Quit
    
    ' ��L�ŏI�����Ȃ��ꍇ�̍ŏI��i
    End
End Sub

' ��񃍃O�L�^
Public Sub LogInfo(ByVal functionName As String, ByVal message As String)
    On Error Resume Next
    
    Dim logMessage As String
    logMessage = Format(Now, "yyyy-mm-dd hh:nn:ss") & " [INFO] " & functionName & ": " & message
    
    
    ' �f�o�b�O�o�́i�J�����̂݁j
    Debug.Print logMessage
End Sub

' �x�����O�L�^
Public Sub LogWarn(ByVal functionName As String, ByVal message As String)
    On Error Resume Next
    
    Dim logMessage As String
    logMessage = Format(Now, "yyyy-mm-dd hh:nn:ss") & " [WARN] " & functionName & ": " & message
    
    
    ' �f�o�b�O�o�́i�J�����̂݁j
    Debug.Print logMessage
End Sub


' �O�����O�t�@�C���o�́i���S�Łj
Private Sub WriteExternalLogSafe(ByVal logMessage As String)
    On Error Resume Next
    
    Dim logDir As String
    Dim logFilePath As String
    Dim fileNum As Integer
    
    logDir = DEFAULT_LOG_DIR
    
    ' �f�B���N�g�������݂��Ȃ��ꍇ�͍쐬�i�G���[�����j
    If Dir(logDir, vbDirectory) = "" Then
        MkDir logDir
    End If
    
    ' ���t�ʃ��O�t�@�C��
    logFilePath = logDir & "system_" & Format(Now, "yyyymmdd") & ".log"
    
    fileNum = FreeFile
    Open logFilePath For Append As fileNum
    Print #fileNum, logMessage
    Close fileNum
End Sub

' �O�����O�t�@�C���o�́i�]���Łj
Private Sub WriteExternalLog(ByVal logMessage As String)
    On Error Resume Next
    Call WriteExternalLogSafe(logMessage)
End Sub

'=============================================================================
' �t�H���g�E�F�ݒ�֐�
'=============================================================================

' �V�X�e���W���t�H���g�K�p
Public Sub ApplySystemFont(ByVal targetRange As Range)
    On Error GoTo ErrHandler
    
    If targetRange Is Nothing Then Exit Sub
    
    With targetRange.Font
        .Name = FONT_NAME
        .Size = FONT_SIZE_NORMAL
        .Color = FONT_COLOR_NORMAL
        .Bold = False
    End With
    
    Exit Sub
    
ErrHandler:
    Call LogError("ApplySystemFont", Err.Description)
End Sub

' �w�b�_�[�t�H���g�K�p
Public Sub ApplyHeaderFont(ByVal targetRange As Range)
    On Error GoTo ErrHandler
    
    If targetRange Is Nothing Then Exit Sub
    
    With targetRange
        With .Font
            .Name = FONT_NAME
            .Size = FONT_SIZE_HEADER
            .Color = FONT_COLOR_HEADER
            .Bold = True
        End With
        .Interior.Color = BG_COLOR_HEADER
    End With
    
    Exit Sub
    
ErrHandler:
    Call LogError("ApplyHeaderFont", Err.Description)
End Sub

' �{�^���t�H���g�K�p
Public Sub ApplyButtonFont(ByVal targetRange As Range)
    On Error GoTo ErrHandler
    
    If targetRange Is Nothing Then Exit Sub
    
    With targetRange.Font
        .Name = FONT_BUTTON
        .Size = FONT_SIZE_BUTTON
        .Color = FONT_COLOR_NORMAL
        .Bold = True
    End With
    
    Exit Sub
    
ErrHandler:
    Call LogError("ApplyButtonFont", Err.Description)
End Sub

' �V�[�g�S�̃t�H���g����
Public Sub ApplySheetFont(ByVal ws As Worksheet)
    On Error GoTo ErrHandler
    
    If ws Is Nothing Then Exit Sub
    
    ' �V�[�g�S�̂ɕW���t�H���g��K�p
    With ws.Cells.Font
        .Name = FONT_NAME
        .Size = FONT_SIZE_NORMAL
        .Color = FONT_COLOR_NORMAL
    End With
    
    Exit Sub
    
ErrHandler:
    Call LogError("ApplySheetFont", Err.Description)
End Sub

' �e�[�u���W���t�H�[�}�b�g�K�p
Public Sub ApplyStandardTableFormat(ByVal tbl As ListObject)
    On Error GoTo ErrHandler
    
    If tbl Is Nothing Then Exit Sub
    
    ' �w�b�_�[�s�t�H�[�}�b�g
    Call ApplyHeaderFont(tbl.HeaderRowRange)
    
    ' �f�[�^�s�t�H�[�}�b�g
    If Not tbl.DataBodyRange Is Nothing Then
        Call ApplySystemFont(tbl.DataBodyRange)
        
        ' ���ݍs�̔w�i�F�ݒ�i�[�u���ȁj
        Dim i As Long
        For i = 1 To tbl.ListRows.Count Step 2
            tbl.ListRows(i).Range.Interior.Color = BG_COLOR_ALTERNATE
        Next i
    End If
    
    ' �g���ݒ�
    With tbl.Range.Borders
        .LineStyle = xlContinuous
        .Color = BORDER_COLOR_DEFAULT
        .Weight = xlThin
    End With
    
    Exit Sub
    
ErrHandler:
    Call LogError("ApplyStandardTableFormat", Err.Description)
End Sub

'=============================================================================
' �i���\���Ǘ�
'=============================================================================

' �i���\���J�n
Public Sub ShowProgressStart(ByVal message As String)
    On Error Resume Next
    Application.StatusBar = message
    Application.ScreenUpdating = False
End Sub

' �i���X�V
Public Sub UpdateProgress(ByVal message As String)
    On Error Resume Next
    Application.StatusBar = message
End Sub

' �i���\���I��
Public Sub HideProgress()
    On Error Resume Next
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub

'=============================================================================
' ���؃w���p�[�֐�
'=============================================================================

' ��l�`�F�b�N
Public Function IsEmpty(ByVal value As Variant) As Boolean
    IsEmpty = (VarType(value) = vbEmpty Or VarType(value) = vbNull Or Len(Trim(CStr(value))) = 0)
End Function

' ���t���S�ϊ�
Public Function SafeDate(ByVal value As Variant) As Date
    On Error Resume Next
    SafeDate = CDate(value)
    If Err.Number <> 0 Then SafeDate = 0
End Function

' ���l���S�ϊ�
Public Function SafeLong(ByVal value As Variant) As Long
    On Error Resume Next
    SafeLong = CLng(value)
    If Err.Number <> 0 Then SafeLong = 0
End Function

'=============================================================================
' ���K�\�����؊֐�
'=============================================================================

' ���[���A�h���X�`������
Public Function IsValidEmail(ByVal email As String) As Boolean
    On Error Resume Next
    
    If Len(email) < 5 Or InStr(email, "@") = 0 Then
        IsValidEmail = False
        Exit Function
    End If
    
    ' ���K�\���p�^�[���`�F�b�N
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = REGEX_EMAIL
    regex.IgnoreCase = True
    
    IsValidEmail = regex.Test(email)
End Function

' �d�b�ԍ��`������
Public Function IsValidPhone(ByVal phone As String) As Boolean
    On Error Resume Next
    
    If Len(phone) < 10 Or Len(phone) > 15 Then
        IsValidPhone = False
        Exit Function
    End If
    
    ' ���K�\���p�^�[���`�F�b�N
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = REGEX_PHONE
    
    IsValidPhone = regex.Test(phone)
End Function

' �X�֔ԍ��`������
Public Function IsValidZip(ByVal zip As String) As Boolean
    On Error Resume Next
    
    If Len(zip) <> 8 And Len(zip) <> 7 Then  ' 123-4567 �܂��� 1234567
        IsValidZip = False
        Exit Function
    End If
    
    ' ���K�\���p�^�[���`�F�b�N
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = REGEX_ZIP
    
    IsValidZip = regex.Test(zip)
End Function

' �ڋqID�`������
Public Function IsValidCustomerId(ByVal customerId As String) As Boolean
    On Error Resume Next
    
    If Len(customerId) < 3 Or Len(customerId) > 20 Then
        IsValidCustomerId = False
        Exit Function
    End If
    
    ' ���K�\���p�^�[���`�F�b�N�i�p�����̂݁j
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = REGEX_CUSTOMERID
    regex.IgnoreCase = True
    
    IsValidCustomerId = regex.Test(customerId)
End Function

'=============================================================================
' �ǉ��̃w���p�[�֐��i�݌v�d�l�Ή��j
'=============================================================================

' ���S�Ȑ��l�ϊ�
Public Function SafeInteger(ByVal value As Variant) As Integer
    On Error Resume Next
    SafeInteger = CInt(value)
    If Err.Number <> 0 Then SafeInteger = 0
End Function

' ���S��Double�ϊ�
Public Function SafeDouble(ByVal value As Variant) As Double
    On Error Resume Next
    SafeDouble = CDbl(value)
    If Err.Number <> 0 Then SafeDouble = 0
End Function

' �󔒂܂��͋󕶎��`�F�b�N
Public Function IsNullOrEmpty(ByVal value As Variant) As Boolean
    IsNullOrEmpty = (IsNull(value) Or VarType(value) = vbEmpty Or Len(Trim(CStr(value & ""))) = 0)
End Function

' ������̍ő咷����
Public Function LimitLength(ByVal text As String, ByVal maxLength As Integer) As String
    If Len(text) > maxLength Then
        LimitLength = Left(text, maxLength)
    Else
        LimitLength = text
    End If
End Function

' ���{�̗X�֔ԍ��p�^�[���ڍ׌���
Public Function IsValidJapaneseZip(ByVal zip As String) As Boolean
    On Error Resume Next
    
    IsValidJapaneseZip = False
    
    If Len(zip) <> 8 And Len(zip) <> 7 Then Exit Function
    
    ' �n�C�t������̏ꍇ (123-4567)
    If Len(zip) = 8 Then
        If Mid(zip, 4, 1) <> "-" Then Exit Function
        If Not IsNumeric(Left(zip, 3)) Then Exit Function
        If Not IsNumeric(Right(zip, 4)) Then Exit Function
        IsValidJapaneseZip = True
    ' �n�C�t���Ȃ��̏ꍇ (1234567)
    ElseIf Len(zip) = 7 Then
        If IsNumeric(zip) Then IsValidJapaneseZip = True
    End If
End Function

' ���{�̓d�b�ԍ��p�^�[������
Public Function IsValidJapanesePhone(ByVal phone As String) As Boolean
    On Error Resume Next
    
    IsValidJapanesePhone = False
    
    ' ��{�����`�F�b�N
    If Len(phone) < 10 Or Len(phone) > 15 Then Exit Function
    
    ' �p�^�[���`�F�b�N�i�Œ�d�b�A�g�ѓd�b�j
    If Left(phone, 1) = "0" Then
        ' �Œ�d�b: 0X-XXXX-XXXX �܂��� 0XX-XXX-XXXX
        ' �g�ѓd�b: 090-XXXX-XXXX, 080-XXXX-XXXX, 070-XXXX-XXXX
        Dim parts As Variant
        parts = Split(phone, "-")
        
        If UBound(parts) = 2 Then ' 3�̕����ɕ�����Ă���
            If Left(phone, 3) = "090" Or Left(phone, 3) = "080" Or Left(phone, 3) = "070" Then
                ' �g�ѓd�b�p�^�[��
                If Len(parts(0)) = 3 And Len(parts(1)) = 4 And Len(parts(2)) = 4 Then
                    IsValidJapanesePhone = True
                End If
            ElseIf Left(phone, 2) = "03" Or Left(phone, 2) = "06" Then
                ' ��v�s�s�Œ�d�b�p�^�[��
                If Len(parts(0)) = 2 And Len(parts(1)) = 4 And Len(parts(2)) = 4 Then
                    IsValidJapanesePhone = True
                End If
            Else
                ' ���̑��Œ�d�b�p�^�[��
                If Len(parts(0)) = 3 And Len(parts(1)) = 3 And Len(parts(2)) = 4 Then
                    IsValidJapanesePhone = True
                End If
            End If
        End If
    End If
End Function

' CSV�s�̃t�B�[���h������
Public Function ValidateCsvFieldCount(ByVal fields As Variant, ByVal expectedCount As Integer) As Boolean
    On Error Resume Next
    
    If IsArray(fields) Then
        ValidateCsvFieldCount = (UBound(fields) + 1 = expectedCount)
    Else
        ValidateCsvFieldCount = False
    End If
End Function

' �ݒ�l�̃f�t�H���g�擾
Public Function GetDefaultConfigValue(ByVal key As String) As String
    On Error Resume Next
    
    Select Case UCase(key)
        Case "CSV_DIR"
            GetDefaultConfigValue = "C:\Data\Import\"
        Case "CSV_FILE"
            GetDefaultConfigValue = "customers_*.csv"
        Case "PRIMARY_KEY"
            GetDefaultConfigValue = "CustomerID"
        Case "ALT_KEY"
            GetDefaultConfigValue = "Email+CustomerName"
        Case "REQUIRED"
            GetDefaultConfigValue = "CustomerID,CustomerName,Status"
        Case "INACTIVATE_DAYS"
            GetDefaultConfigValue = "180"
        Case "EMAIL_REGEX"
            GetDefaultConfigValue = REGEX_EMAIL
        Case "ZIP_REGEX"
            GetDefaultConfigValue = REGEX_ZIP
        Case "PHONE_REGEX"
            GetDefaultConfigValue = REGEX_PHONE
        Case "CUSTOMERID_REGEX"
            GetDefaultConfigValue = REGEX_CUSTOMERID
        Case Else
            GetDefaultConfigValue = ""
    End Select
End Function

' �R���N�V��������z��ւ̕ϊ�
Public Function CollectionToArray(ByVal col As Collection) As Variant
    On Error Resume Next
    
    If col.Count = 0 Then
        CollectionToArray = Array()
        Exit Function
    End If
    
    Dim arr() As Variant
    ReDim arr(col.Count - 1)
    
    Dim i As Integer
    For i = 1 To col.Count
        arr(i - 1) = col(i)
    Next i
    
    CollectionToArray = arr
End Function

' �z�񂩂�R���N�V�����ւ̕ϊ�
Public Function ArrayToCollection(ByVal arr As Variant) As Collection
    On Error Resume Next
    
    Set ArrayToCollection = New Collection
    
    If Not IsArray(arr) Then Exit Function
    
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        ArrayToCollection.Add arr(i)
    Next i
End Function

' �t�@�C�����݃`�F�b�N
Public Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(filePath) <> "")
End Function

' �f�B���N�g�����݃`�F�b�N
Public Function DirectoryExists(ByVal dirPath As String) As Boolean
    On Error Resume Next
    DirectoryExists = (Dir(dirPath, vbDirectory) <> "")
End Function

' ���S�ȃf�B���N�g���쐬
Public Function CreateDirectoryIfNotExists(ByVal dirPath As String) As Boolean
    On Error Resume Next
    
    CreateDirectoryIfNotExists = False
    
    If DirectoryExists(dirPath) Then
        CreateDirectoryIfNotExists = True
        Exit Function
    End If
    
    MkDir dirPath
    If Err.Number = 0 Then
        CreateDirectoryIfNotExists = True
    End If
End Function

' ���݂̓����𕶎���Ƃ��Ď擾
Public Function GetCurrentDateTimeString() As String
    GetCurrentDateTimeString = Format(Now, "yyyy-mm-dd hh:nn:ss")
End Function

' ���݂̓��t�𕶎���Ƃ��Ď擾
Public Function GetCurrentDateString() As String
    GetCurrentDateString = Format(Date, "yyyy-mm-dd")
End Function

' �~���b���܂ތ��ݎ����擾
Public Function GetCurrentTimeStamp() As String
    GetCurrentTimeStamp = Format(Now, "yyyy-mm-dd hh:nn:ss") & "." & Format(Timer Mod 1 * 1000, "000")
End Function