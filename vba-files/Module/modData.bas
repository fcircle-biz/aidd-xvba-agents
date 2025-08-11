Attribute VB_Name = "modData"
'=============================================================================
' modData.bas - �f�[�^�A�N�Z�X�ECSV�������W���[��
'=============================================================================
' �T�v:
'   CSV��荞�݁A�ݒ�l�Ǘ��A�e�[�u������A�t�@�C��I/O���̃f�[�^�A�N�Z�X�w
'   �O���f�[�^�\�[�X�Ƃ̘A�g�A�ݒ�l�̎擾�E�ۑ��@�\���
'=============================================================================
Option Explicit

'=============================================================================
' �ݒ�l�Ǘ�
'=============================================================================

' �ݒ�l�擾
Public Function GetConfigValue(ByVal configKey As String) As String
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_CONFIG)
    If ws Is Nothing Then GoTo ErrHandler
    
    Set tbl = modCmn.GetTable(ws, TABLE_CONFIG)
    If tbl Is Nothing Then GoTo ErrHandler
    
    ' �ݒ�e�[�u������w��L�[�̒l������
    For Each row In tbl.ListRows
        If modCmn.GetRowText(row, "ConfigKey") = configKey Then
            GetConfigValue = modCmn.GetRowText(row, "ConfigValue")
            Exit Function
        End If
    Next row
    
    ' �ݒ肪������Ȃ��ꍇ�̓f�t�H���g�l��Ԃ�
    GetConfigValue = GetDefaultConfigValue(configKey)
    Call modCmn.LogWarn("GetConfigValue", "�ݒ�L�[����`�A�f�t�H���g�l�g�p: " & configKey)
    Exit Function
    
ErrHandler:
    GetConfigValue = GetDefaultConfigValue(configKey)
    Call modCmn.LogError("GetConfigValue", "�ݒ�l�擾�G���[: " & configKey & " - " & Err.Description)
End Function

' �f�t�H���g�ݒ�l�擾
Public Function GetDefaultConfigValue(ByVal key As String) As String
    On Error Resume Next
    
    Select Case UCase(key)
        Case CONFIG_CSV_DIR
            GetDefaultConfigValue = DEFAULT_CSV_DIR
        Case CONFIG_CSV_FILE
            GetDefaultConfigValue = DEFAULT_CSV_FILE
        Case CONFIG_PRIMARY_KEY
            GetDefaultConfigValue = DEFAULT_PRIMARY_KEY
        Case CONFIG_ALT_KEY
            GetDefaultConfigValue = DEFAULT_ALT_KEY
        Case CONFIG_REQUIRED
            GetDefaultConfigValue = DEFAULT_REQUIRED
        Case CONFIG_INACTIVATE_DAYS
            GetDefaultConfigValue = CStr(DEFAULT_INACTIVATE_DAYS)
        Case CONFIG_EMAIL_REGEX
            GetDefaultConfigValue = REGEX_EMAIL_STRICT
        Case CONFIG_ZIP_REGEX
            GetDefaultConfigValue = REGEX_ZIP_JAPAN
        Case CONFIG_PHONE_REGEX
            GetDefaultConfigValue = REGEX_PHONE_JAPAN
        Case CONFIG_BACKUP_DIR
            GetDefaultConfigValue = DEFAULT_BACKUP_DIR
        Case CONFIG_BACKUP_ENABLED
            GetDefaultConfigValue = "True"
        Case Else
            GetDefaultConfigValue = ""
    End Select
End Function

' �ݒ�l�ۑ�
Public Sub SetConfigValue(ByVal configKey As String, ByVal configValue As String)
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    Dim found As Boolean
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_CONFIG)
    If ws Is Nothing Then GoTo ErrHandler
    
    Set tbl = modCmn.GetTable(ws, TABLE_CONFIG)
    If tbl Is Nothing Then GoTo ErrHandler
    
    ' �����̐ݒ������
    For Each row In tbl.ListRows
        If modCmn.GetRowText(row, "ConfigKey") = configKey Then
            Call modCmn.SetRowText(row, "ConfigValue", configValue)
            found = True
            Exit For
        End If
    Next row
    
    ' �V�����ݒ�̏ꍇ�͒ǉ�
    If Not found Then
        Set row = tbl.ListRows.Add
        Call modCmn.SetRowText(row, "ConfigKey", configKey)
        Call modCmn.SetRowText(row, "ConfigValue", configValue)
    End If
    
    Call modCmn.LogInfo("SetConfigValue", "�ݒ�ۑ�����: " & configKey & " = " & configValue)
    Exit Sub
    
ErrHandler:
    Call modCmn.LogError("SetConfigValue", "�ݒ�l�ۑ��G���[: " & configKey & " - " & Err.Description)
End Sub

'=============================================================================
' CSV�����֐�
'=============================================================================

' CSV�t�@�C���ꗗ�擾
Public Function GetCsvFileList() As Collection
    On Error GoTo ErrHandler
    
    Dim csvDir As String
    Dim filePattern As String
    Dim fileName As String
    
    Set GetCsvFileList = New Collection
    
    csvDir = GetConfigValue(CONFIG_CSV_DIR)
    filePattern = GetConfigValue(CONFIG_CSV_FILE)
    
    ' �f�B���N�g�����݃`�F�b�N
    If Not modCmn.DirectoryExists(csvDir) Then
        Call modCmn.LogWarn("GetCsvFileList", "CSV�f�B���N�g�������݂��܂���: " & csvDir)
        Exit Function
    End If
    
    ' ���C���h�J�[�h�p�^�[����Dir�֐��p�ɕϊ�
    fileName = Dir(csvDir & filePattern)
    Do While fileName <> ""
        GetCsvFileList.Add csvDir & fileName
        fileName = Dir()
    Loop
    
    Call modCmn.LogInfo("GetCsvFileList", "CSV�t�@�C�� " & GetCsvFileList.Count & " ������")
    Exit Function
    
ErrHandler:
    Set GetCsvFileList = New Collection
    Call modCmn.LogError("GetCsvFileList", "CSV�t�@�C���ꗗ�擾�G���[: " & Err.Description)
End Function

' CSV��Staging��荞��
Public Sub ImportCsvToStaging()
    On Error GoTo ErrHandler
    
    Dim csvFiles As Collection
    Dim filePath As Variant
    Dim totalRecords As Long
    Dim startTime As Double
    
    startTime = Timer
    Call modCmn.ShowProgressStart(MSG_IMPORT_STARTED)
    
    ' Staging�N���A
    Call ClearStagingData
    
    ' CSV�t�@�C���ꗗ�擾
    Set csvFiles = GetCsvFileList()
    If csvFiles.Count = 0 Then
        Call modCmn.LogWarn("ImportCsvToStaging", ERR_CSV_NOT_FOUND)
        MsgBox ERR_CSV_NOT_FOUND, vbExclamation
        GoTo ExitHandler
    End If
    
    ' �eCSV�t�@�C��������
    For Each filePath In csvFiles
        Call modCmn.UpdateProgress("CSV������: " & Dir(CStr(filePath)))
        totalRecords = totalRecords + ImportSingleCsvFile(CStr(filePath))
    Next filePath
    
    ' Staging�f�[�^���K��
    Call NormalizeStagingData
    
    ' ���O�L�^
    Call LogImportOperation("CSV��荞�݊���", totalRecords, Timer - startTime, "")
    
    MsgBox MSG_IMPORT_COMPLETED & vbCrLf & _
           "��������: " & Format(totalRecords, NUMBER_FORMAT_COUNT) & " ��" & vbCrLf & _
           "��������: " & Format(Timer - startTime, "0.0") & " �b", vbInformation
    
ExitHandler:
    Call modCmn.HideProgress
    Exit Sub
    
ErrHandler:
    Call modCmn.HideProgress
    Call modCmn.LogError("ImportCsvToStaging", "CSV��荞�݃G���[: " & Err.Description)
End Sub

' �P��CSV�t�@�C����荞��
Private Function ImportSingleCsvFile(ByVal filePath As String) As Long
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fileNum As Integer
    Dim lineData As String
    Dim fields As Variant
    Dim recordCount As Long
    Dim lineNumber As Long
    Dim row As ListRow
    Dim fileName As String
    
    ImportSingleCsvFile = 0
    fileName = Dir(filePath)
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_STAGING)
    If ws Is Nothing Then GoTo ErrHandler
    
    Set tbl = modCmn.GetTable(ws, TABLE_STAGING)
    If tbl Is Nothing Then GoTo ErrHandler
    
    ' �t�@�C���ǂݍ��݊J�n
    fileNum = FreeFile
    Open filePath For Input As fileNum
    
    ' �w�b�_�[�s�X�L�b�v
    Line Input #fileNum, lineData
    lineNumber = 1
    
    ' �f�[�^�s����
    Do Until EOF(fileNum)
        Line Input #fileNum, lineData
        lineNumber = lineNumber + 1
        
        If Len(Trim(lineData)) > 0 Then
            fields = ParseCsvLine(lineData)
            If IsArray(fields) And UBound(fields) >= 8 Then ' �Œ�K�v�ȗ񐔃`�F�b�N
                Set row = tbl.ListRows.Add
                Call SetStagingRowFromCsv(row, fields, fileName)
                recordCount = recordCount + 1
                
                ' �o�b�`�����Ńv���O���X�X�V
                If recordCount Mod BATCH_SIZE_CSV_IMPORT = 0 Then
                    Call modCmn.UpdateProgress("CSV������: " & fileName & " (" & recordCount & " ��)")
                End If
            Else
                Call modCmn.LogWarn("ImportSingleCsvFile", "�s����CSV�s���X�L�b�v: " & fileName & " �s" & lineNumber)
            End If
        End If
    Loop
    
    Close fileNum
    ImportSingleCsvFile = recordCount
    
    Call modCmn.LogInfo("ImportSingleCsvFile", fileName & " ��荞�݊���: " & recordCount & " ��")
    Exit Function
    
ErrHandler:
    If fileNum > 0 Then Close fileNum
    ImportSingleCsvFile = 0
    Call modCmn.LogError("ImportSingleCsvFile", "�t�@�C����荞�݃G���[: " & filePath & " - " & Err.Description)
End Function

' CSV�s�p�[�X�i�ȈՔŁj
Private Function ParseCsvLine(ByVal lineData As String) As Variant
    On Error Resume Next
    
    ' �J���}��؂�ŕ����i���p�����̃J���}�͍l�����Ȃ��ȈՔŁj
    ParseCsvLine = Split(lineData, CSV_DELIMITER)
End Function

' Staging�e�[�u����CSV�f�[�^�ݒ�
Private Sub SetStagingRowFromCsv(ByVal row As ListRow, ByVal fields As Variant, ByVal fileName As String)
    On Error Resume Next
    
    If Not IsArray(fields) Then Exit Sub
    
    ' CSV���Staging�e�[�u����Ƀ}�b�s���O
    Call modCmn.SetRowText(row, COL_CUSTOMER_ID, GetFieldValue(fields, 0))
    Call modCmn.SetRowText(row, COL_CUSTOMER_NAME, GetFieldValue(fields, 1))
    Call modCmn.SetRowText(row, COL_EMAIL, GetFieldValue(fields, 2))
    Call modCmn.SetRowText(row, COL_PHONE, GetFieldValue(fields, 3))
    Call modCmn.SetRowText(row, COL_ZIP, GetFieldValue(fields, 4))
    Call modCmn.SetRowText(row, COL_ADDRESS1, GetFieldValue(fields, 5))
    Call modCmn.SetRowText(row, COL_ADDRESS2, GetFieldValue(fields, 6))
    Call modCmn.SetRowText(row, COL_CATEGORY, GetFieldValue(fields, 7))
    Call modCmn.SetRowText(row, COL_STATUS, GetFieldValue(fields, 8))
    Call modCmn.SetRowText(row, COL_SOURCE_FILE, fileName)
End Sub

' �z�񂩂���S�Ƀt�B�[���h�l�擾
Private Function GetFieldValue(ByVal fields As Variant, ByVal index As Integer) As String
    On Error Resume Next
    
    If IsArray(fields) And index <= UBound(fields) Then
        GetFieldValue = modCmn.TrimAll(CStr(fields(index)))
    Else
        GetFieldValue = ""
    End If
End Function

'=============================================================================
' �f�[�^���K��
'=============================================================================

' Staging�f�[�^���K��
Public Sub NormalizeStagingData()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    Dim recordCount As Long
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_STAGING)
    If ws Is Nothing Then GoTo ErrHandler
    
    Set tbl = modCmn.GetTable(ws, TABLE_STAGING)
    If tbl Is Nothing Then GoTo ErrHandler
    
    Call modCmn.ShowProgressStart(MSG_VALIDATION_STARTED)
    
    For Each row In tbl.ListRows
        Call NormalizeStagingRow(row)
        recordCount = recordCount + 1
        
        If recordCount Mod BATCH_SIZE_VALIDATION = 0 Then
            Call modCmn.UpdateProgress("���K��������: " & recordCount & " ��")
        End If
    Next row
    
    Call modCmn.LogInfo("NormalizeStagingData", "�f�[�^���K������: " & recordCount & " ��")
    Call modCmn.HideProgress
    Exit Sub
    
ErrHandler:
    Call modCmn.HideProgress
    Call modCmn.LogError("NormalizeStagingData", "�f�[�^���K���G���[: " & Err.Description)
End Sub

' Staging�s�̐��K��
Private Sub NormalizeStagingRow(ByVal row As ListRow)
    On Error Resume Next
    
    Dim email As String
    Dim phone As String
    Dim zip As String
    Dim customerName As String
    Dim keyCandidate As String
    
    ' �ڋq�����K��
    customerName = modCmn.TrimAll(modCmn.GetRowText(row, COL_CUSTOMER_NAME))
    Call modCmn.SetRowText(row, COL_CUSTOMER_NAME, customerName)
    
    ' ���[�����K��
    email = modCmn.NormalizeEmail(modCmn.GetRowText(row, COL_EMAIL))
    Call modCmn.SetRowText(row, COL_EMAIL_NORM, email)
    
    ' �d�b�ԍ����K��
    phone = modCmn.NormalizePhone(modCmn.GetRowText(row, COL_PHONE))
    Call modCmn.SetRowText(row, COL_PHONE_NORM, phone)
    
    ' �X�֔ԍ����K��
    zip = modCmn.NormalizeZip(modCmn.GetRowText(row, COL_ZIP))
    Call modCmn.SetRowText(row, COL_ZIP_NORM, zip)
    
    ' ��փL�[����
    keyCandidate = BuildAlternateKey(row)
    Call modCmn.SetRowText(row, COL_KEY_CANDIDATE, keyCandidate)
End Sub

' ��փL�[�\�z�iEmail + CustomerName�j
Private Function BuildAlternateKey(ByVal row As ListRow) As String
    On Error Resume Next
    
    Dim email As String
    Dim customerName As String
    
    email = modCmn.GetRowText(row, COL_EMAIL_NORM)
    customerName = modCmn.GetRowText(row, COL_CUSTOMER_NAME)
    
    If Len(email) > 0 And Len(customerName) > 0 Then
        BuildAlternateKey = email & "+" & customerName
    ElseIf Len(customerName) > 0 Then
        BuildAlternateKey = customerName
    Else
        BuildAlternateKey = ""
    End If
End Function

'=============================================================================
' �e�[�u���Ǘ�
'=============================================================================

' Staging�f�[�^�N���A
Public Sub ClearStagingData()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_STAGING)
    If ws Is Nothing Then GoTo ErrHandler
    
    Call modCmn.SafeClearSheet(ws, keepFormats:=True)
    Call EnsureStagingTableStructure
    
    Call modCmn.LogInfo("ClearStagingData", "Staging�f�[�^�N���A����")
    Exit Sub
    
ErrHandler:
    Call modCmn.LogError("ClearStagingData", "Staging�N���A�G���[: " & Err.Description)
End Sub

' Staging�e�[�u���\���m��
Private Sub EnsureStagingTableStructure()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_STAGING)
    If ws Is Nothing Then GoTo ErrHandler
    
    ' �����e�[�u���폜
    If modCmn.TableExists(ws, TABLE_STAGING) Then
        ws.ListObjects(TABLE_STAGING).Delete
    End If
    
    ' �w�b�_�[�ݒ�
    Call SetTableHeaders(ws, "A1", STAGING_HEADERS)
    
    ' �e�[�u���쐬
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes)
    tbl.Name = TABLE_STAGING
    
    ' �t�H�[�}�b�g�K�p
    Call modCmn.ApplyStandardTableFormat(tbl)
    
    Exit Sub
    
ErrHandler:
    Call modCmn.LogError("EnsureStagingTableStructure", "Staging�e�[�u���\���G���[: " & Err.Description)
End Sub

' �e�[�u���w�b�_�[�ݒ�
Private Sub SetTableHeaders(ByVal ws As Worksheet, ByVal startCell As String, ByVal headers As String)
    On Error Resume Next
    
    Dim headerArray As Variant
    Dim i As Integer
    Dim startRange As Range
    
    headerArray = Split(headers, ",")
    Set startRange = ws.Range(startCell)
    
    For i = 0 To UBound(headerArray)
        startRange.Offset(0, i).Value = Trim(headerArray(i))
    Next i
End Sub

' �ڋq�e�[�u���擾
Public Function GetCustomersTable() As ListObject
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_CUSTOMERS)
    If ws Is Nothing Then GoTo ErrHandler
    
    Set GetCustomersTable = modCmn.GetTable(ws, TABLE_CUSTOMERS)
    Exit Function
    
ErrHandler:
    Set GetCustomersTable = Nothing
    Call modCmn.LogError("GetCustomersTable", "�ڋq�e�[�u���擾�G���[: " & Err.Description)
End Function

' Staging�e�[�u���擾
Public Function GetStagingTable() As ListObject
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_STAGING)
    If ws Is Nothing Then GoTo ErrHandler
    
    Set GetStagingTable = modCmn.GetTable(ws, TABLE_STAGING)
    Exit Function
    
ErrHandler:
    Set GetStagingTable = Nothing
    Call modCmn.LogError("GetStagingTable", "Staging�e�[�u���擾�G���[: " & Err.Description)
End Function

'=============================================================================
' ���O�L�^
'=============================================================================

' �C���|�[�g���샍�O�L�^
Public Sub LogImportOperation(ByVal message As String, ByVal recordCount As Long, _
                             ByVal processTime As Double, ByVal sourceFile As String)
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_LOGS)
    If ws Is Nothing Then Exit Sub
    
    Set tbl = modCmn.GetTable(ws, TABLE_LOGS)
    If tbl Is Nothing Then Exit Sub
    
    Set row = tbl.ListRows.Add
    Call modCmn.SetRowText(row, "Timestamp", modCmn.GetCurrentDateTimeString())
    Call modCmn.SetRowText(row, "Level", LOG_LEVEL_INFO)
    Call modCmn.SetRowText(row, "Module", "modData")
    Call modCmn.SetRowText(row, "Message", message)
    Call modCmn.SetRowText(row, "RecordCount", CStr(recordCount))
    Call modCmn.SetRowText(row, "ProcessTime", Format(processTime, "0.00") & "�b")
    Call modCmn.SetRowText(row, "SourceFile", sourceFile)
    
    Exit Sub
    
ErrHandler:
    Call modCmn.LogError("LogImportOperation", "���O�L�^�G���[: " & Err.Description)
End Sub

' �G���[���O�L�^
Public Sub LogErrorOperation(ByVal message As String, ByVal details As String, ByVal sourceFile As String)
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_LOGS)
    If ws Is Nothing Then Exit Sub
    
    Set tbl = modCmn.GetTable(ws, TABLE_LOGS)
    If tbl Is Nothing Then Exit Sub
    
    Set row = tbl.ListRows.Add
    Call modCmn.SetRowText(row, "Timestamp", modCmn.GetCurrentDateTimeString())
    Call modCmn.SetRowText(row, "Level", LOG_LEVEL_ERROR)
    Call modCmn.SetRowText(row, "Module", "modData")
    Call modCmn.SetRowText(row, "Message", message)
    Call modCmn.SetRowText(row, "Details", details)
    Call modCmn.SetRowText(row, "SourceFile", sourceFile)
    
    Exit Sub
    
ErrHandler:
    ' ���O�G���[�̏ꍇ�͊O�����O�̂ݏo��
    Call modCmn.LogError("LogErrorOperation", "���O�L�^�G���[: " & Err.Description)
End Sub