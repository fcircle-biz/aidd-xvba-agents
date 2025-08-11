Attribute VB_Name = "modUtils"
'=============================================================================
' modUtils.bas - ���[�e�B���e�B�E�w���p�[�֐����W���[��
'=============================================================================
' �T�v:
'   �V�X�e���S�̂Ŏg�p����郆�[�e�B���e�B�֐��Q
'   �t�@�C������A�e�[�u���������A�f�[�^�ϊ��A�V�X�e���Ǘ����̕⏕�@�\
'=============================================================================
Option Explicit

'=============================================================================
' �V�X�e�����������[�e�B���e�B
'=============================================================================

' �V�X�e���S�̏�����
Public Sub InitializeCustomerSystem()
    On Error GoTo ErrHandler
    
    Call modCmn.LogInfo("InitializeCustomerSystem", "�ڋq�Ǘ��V�X�e���������J�n")
    
    ' �e�V�[�g�̏�����
    Call InitializeAllSheets()
    
    ' �e�[�u���\���m�F�E�쐬
    Call EnsureAllTableStructures()
    
    ' �f�t�H���g�ݒ�l�ݒ�
    Call SetupDefaultConfiguration()
    
    ' �T���v���f�[�^�쐬�i����̂݁j
    Call CreateInitialSampleData()
    
    ' �_�b�V���{�[�h������
    Call modDashboard.InitializeDashboard()
    
    Call modCmn.LogInfo("InitializeCustomerSystem", "�ڋq�Ǘ��V�X�e������������")
    
    MsgBox SYSTEM_NAME & " �̏��������������܂����B" & vbCrLf & _
           "�o�[�W����: " & SYSTEM_VERSION, vbInformation, "����������"
    
    Exit Sub
    
ErrHandler:
    Call modCmn.LogError("InitializeCustomerSystem", "�V�X�e���������G���[: " & Err.Description)
End Sub

' �S�V�[�g������
Private Sub InitializeAllSheets()
    On Error Resume Next
    
    ' �V�[�g���ύX
    Call RenameSystemSheets()
    
    ' �e�V�[�g�̃t�H���g����
    Dim i As Integer
    For i = 1 To 6
        Dim ws As Worksheet
        Set ws = modCmn.GetWorksheetByIndex(i)
        If Not ws Is Nothing Then
            Call modCmn.ApplySheetFont(ws)
        End If
    Next i
End Sub

' �V�X�e���V�[�g���ύX
Private Sub RenameSystemSheets()
    On Error Resume Next
    
    Dim ws As Worksheet
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_DASHBOARD)
    If Not ws Is Nothing Then ws.Name = SHEET_DASHBOARD
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_CUSTOMERS)
    If Not ws Is Nothing Then ws.Name = SHEET_CUSTOMERS
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_STAGING)
    If Not ws Is Nothing Then ws.Name = SHEET_STAGING
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_CONFIG)
    If Not ws Is Nothing Then ws.Name = SHEET_CONFIG
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_LOGS)
    If Not ws Is Nothing Then ws.Name = SHEET_LOGS
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_CODEBOOK)
    If Not ws Is Nothing Then ws.Name = SHEET_CODEBOOK
End Sub

'=============================================================================
' �e�[�u���\���Ǘ�
'=============================================================================

' �S�e�[�u���\���m�F�E�쐬
Public Sub EnsureAllTableStructures()
    On Error Resume Next
    
    Call EnsureCustomersTableStructure()
    Call EnsureStagingTableStructure()
    Call EnsureConfigTableStructure()
    Call EnsureLogsTableStructure()
    Call EnsureCodebookTableStructure()
End Sub

' Customers�e�[�u���\���m�F
Private Sub EnsureCustomersTableStructure()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_CUSTOMERS)
    If ws Is Nothing Then Exit Sub
    
    ' �����e�[�u���`�F�b�N
    If Not modCmn.TableExists(ws, TABLE_CUSTOMERS) Then
        ' �e�[�u���V�K�쐬
        Call SetTableHeaders(ws, "A1", CUSTOMERS_HEADERS)
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes)
        tbl.Name = TABLE_CUSTOMERS
        
        ' �t�H�[�}�b�g�K�p
        Call modCmn.ApplyStandardTableFormat(tbl)
        
        Call modCmn.LogInfo("EnsureCustomersTableStructure", "Customers�e�[�u���쐬����")
    End If
    
    Exit Sub
    
ErrHandler:
    Call modCmn.LogError("EnsureCustomersTableStructure", "Customers�e�[�u���\���G���[: " & Err.Description)
End Sub

' Staging�e�[�u���\���m�F
Private Sub EnsureStagingTableStructure()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_STAGING)
    If ws Is Nothing Then Exit Sub
    
    ' �����e�[�u���폜�iStaging�͓s�x�č쐬�j
    If modCmn.TableExists(ws, TABLE_STAGING) Then
        ws.ListObjects(TABLE_STAGING).Delete
    End If
    
    ' �V�[�g�N���A
    Call modCmn.SafeClearSheet(ws, keepFormats:=False)
    
    ' �e�[�u���V�K�쐬
    Call SetTableHeaders(ws, "A1", STAGING_HEADERS)
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes)
    tbl.Name = TABLE_STAGING
    
    ' �t�H�[�}�b�g�K�p
    Call modCmn.ApplyStandardTableFormat(tbl)
    
    Call modCmn.LogInfo("EnsureStagingTableStructure", "Staging�e�[�u���쐬����")
    Exit Sub
    
ErrHandler:
    Call modCmn.LogError("EnsureStagingTableStructure", "Staging�e�[�u���\���G���[: " & Err.Description)
End Sub

' Config�e�[�u���\���m�F
Private Sub EnsureConfigTableStructure()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_CONFIG)
    If ws Is Nothing Then Exit Sub
    
    ' �����e�[�u���`�F�b�N
    If Not modCmn.TableExists(ws, TABLE_CONFIG) Then
        Call SetTableHeaders(ws, "A1", CONFIG_HEADERS)
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes)
        tbl.Name = TABLE_CONFIG
        
        ' �t�H�[�}�b�g�K�p
        Call modCmn.ApplyStandardTableFormat(tbl)
        
        Call modCmn.LogInfo("EnsureConfigTableStructure", "Config�e�[�u���쐬����")
    End If
    
    Exit Sub
    
ErrHandler:
    Call modCmn.LogError("EnsureConfigTableStructure", "Config�e�[�u���\���G���[: " & Err.Description)
End Sub

' Logs�e�[�u���\���m�F
Private Sub EnsureLogsTableStructure()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_LOGS)
    If ws Is Nothing Then Exit Sub
    
    ' �����e�[�u���`�F�b�N
    If Not modCmn.TableExists(ws, TABLE_LOGS) Then
        Call SetTableHeaders(ws, "A1", LOGS_HEADERS)
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes)
        tbl.Name = TABLE_LOGS
        
        ' �t�H�[�}�b�g�K�p
        Call modCmn.ApplyStandardTableFormat(tbl)
        
        Call modCmn.LogInfo("EnsureLogsTableStructure", "Logs�e�[�u���쐬����")
    End If
    
    Exit Sub
    
ErrHandler:
    Call modCmn.LogError("EnsureLogsTableStructure", "Logs�e�[�u���\���G���[: " & Err.Description)
End Sub

' Codebook�e�[�u���\���m�F
Private Sub EnsureCodebookTableStructure()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_CODEBOOK)
    If ws Is Nothing Then Exit Sub
    
    ' �����e�[�u���`�F�b�N
    If Not modCmn.TableExists(ws, TABLE_CODEBOOK) Then
        Call SetTableHeaders(ws, "A1", CODEBOOK_HEADERS)
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes)
        tbl.Name = TABLE_CODEBOOK
        
        ' �t�H�[�}�b�g�K�p
        Call modCmn.ApplyStandardTableFormat(tbl)
        
        ' �f�t�H���g�}�b�s���O�ݒ�
        Call SetupDefaultColumnMappings(tbl)
        
        Call modCmn.LogInfo("EnsureCodebookTableStructure", "Codebook�e�[�u���쐬����")
    End If
    
    Exit Sub
    
ErrHandler:
    Call modCmn.LogError("EnsureCodebookTableStructure", "Codebook�e�[�u���\���G���[: " & Err.Description)
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
        With startRange.Offset(0, i)
            .Value = Trim(headerArray(i))
            Call modCmn.ApplyHeaderFont(.Cells)
        End With
    Next i
End Sub

'=============================================================================
' �f�t�H���g�f�[�^�ݒ�
'=============================================================================

' �f�t�H���g�ݒ�l�ݒ�
Public Sub SetupDefaultConfiguration()
    On Error Resume Next
    
    ' ��{�ݒ�l
    Call modData.SetConfigValue(CONFIG_CSV_DIR, DEFAULT_CSV_DIR)
    Call modData.SetConfigValue(CONFIG_CSV_FILE, DEFAULT_CSV_FILE)
    Call modData.SetConfigValue(CONFIG_PRIMARY_KEY, DEFAULT_PRIMARY_KEY)
    Call modData.SetConfigValue(CONFIG_ALT_KEY, DEFAULT_ALT_KEY)
    Call modData.SetConfigValue(CONFIG_REQUIRED, DEFAULT_REQUIRED)
    Call modData.SetConfigValue(CONFIG_INACTIVATE_DAYS, CStr(DEFAULT_INACTIVATE_DAYS))
    Call modData.SetConfigValue(CONFIG_EMAIL_REGEX, REGEX_EMAIL_STRICT)
    Call modData.SetConfigValue(CONFIG_ZIP_REGEX, REGEX_ZIP_JAPAN)
    Call modData.SetConfigValue(CONFIG_PHONE_REGEX, REGEX_PHONE_JAPAN)
    Call modData.SetConfigValue(CONFIG_BACKUP_ENABLED, "True")
    Call modData.SetConfigValue(CONFIG_BACKUP_DIR, DEFAULT_BACKUP_DIR)
    
    Call modCmn.LogInfo("SetupDefaultConfiguration", "�f�t�H���g�ݒ�l�ݒ芮��")
End Sub

' �f�t�H���g��}�b�s���O�ݒ�
Private Sub SetupDefaultColumnMappings(ByVal codebookTbl As ListObject)
    On Error Resume Next
    
    Dim mappings As Variant
    Dim i As Integer
    
    ' CSV�w�b�_�[�Ɠ�����̃}�b�s���O��`
    mappings = Array( _
        Array("customer_id", COL_CUSTOMER_ID, "������", "", "", "True", "�ڋq����ID"), _
        Array("customer_name", COL_CUSTOMER_NAME, "������", "", "TrimAll", "True", "�ڋq��"), _
        Array("email", COL_EMAIL, "������", REGEX_EMAIL_STRICT, "NormalizeEmail", "False", "���[���A�h���X"), _
        Array("phone", COL_PHONE, "������", REGEX_PHONE_JAPAN, "NormalizePhone", "False", "�d�b�ԍ�"), _
        Array("zip", COL_ZIP, "������", REGEX_ZIP_JAPAN, "NormalizeZip", "False", "�X�֔ԍ�"), _
        Array("address1", COL_ADDRESS1, "������", "", "TrimAll", "False", "�Z��1"), _
        Array("address2", COL_ADDRESS2, "������", "", "TrimAll", "False", "�Z��2"), _
        Array("category", COL_CATEGORY, "������", "", "", "False", "�ڋq�J�e�S��"), _
        Array("status", COL_STATUS, "������", "", "", "True", "�ڋq�X�e�[�^�X") _
    )
    
    ' �}�b�s���O�f�[�^�ǉ�
    For i = 0 To UBound(mappings)
        Dim row As ListRow
        Set row = codebookTbl.ListRows.Add
        
        Call modCmn.SetRowText(row, "ExternalColumnName", CStr(mappings(i)(0)))
        Call modCmn.SetRowText(row, "InternalColumnName", CStr(mappings(i)(1)))
        Call modCmn.SetRowText(row, "DataType", CStr(mappings(i)(2)))
        Call modCmn.SetRowText(row, "ValidationRule", CStr(mappings(i)(3)))
        Call modCmn.SetRowText(row, "NormalizationRule", CStr(mappings(i)(4)))
        Call modCmn.SetRowText(row, "Required", CStr(mappings(i)(5)))
        Call modCmn.SetRowText(row, "Description", CStr(mappings(i)(6)))
    Next i
End Sub

' �����T���v���f�[�^�쐬
Public Sub CreateInitialSampleData()
    On Error Resume Next
    
    Dim customerTbl As ListObject
    Set customerTbl = modData.GetCustomersTable()
    
    ' ���Ƀf�[�^������ꍇ�̓X�L�b�v
    If Not customerTbl Is Nothing Then
        If customerTbl.ListRows.Count > 0 Then Exit Sub
    End If
    
    ' �T���v���ڋq�f�[�^
    Dim sampleData As Variant
    sampleData = Array( _
        Array("SAMPLE001", "�T���v���������", "sample@example.com", "03-1234-5678", "100-0001", "�����s���c��", "���c1-1-1", CATEGORY_B2B, STATUS_ACTIVE, "�V�X�e�������f�[�^"), _
        Array("SAMPLE002", "�e�X�g����", "test@business.co.jp", "06-9876-5432", "530-0001", "���{���s�k��", "�~�c2-2-2", CATEGORY_PARTNER, STATUS_ACTIVE, "�V�X�e�������f�[�^") _
    )
    
    ' �T���v���f�[�^�ǉ�
    Dim i As Integer
    For i = 0 To UBound(sampleData)
        Dim row As ListRow
        Set row = customerTbl.ListRows.Add
        
        Call modCmn.SetRowText(row, COL_CUSTOMER_ID, CStr(sampleData(i)(0)))
        Call modCmn.SetRowText(row, COL_CUSTOMER_NAME, CStr(sampleData(i)(1)))
        Call modCmn.SetRowText(row, COL_EMAIL, CStr(sampleData(i)(2)))
        Call modCmn.SetRowText(row, COL_PHONE, CStr(sampleData(i)(3)))
        Call modCmn.SetRowText(row, COL_ZIP, CStr(sampleData(i)(4)))
        Call modCmn.SetRowText(row, COL_ADDRESS1, CStr(sampleData(i)(5)))
        Call modCmn.SetRowText(row, COL_ADDRESS2, CStr(sampleData(i)(6)))
        Call modCmn.SetRowText(row, COL_CATEGORY, CStr(sampleData(i)(7)))
        Call modCmn.SetRowText(row, COL_STATUS, CStr(sampleData(i)(8)))
        Call modCmn.SetRowDate(row, COL_CREATED_AT, Now)
        Call modCmn.SetRowDate(row, COL_UPDATED_AT, Now)
        Call modCmn.SetRowText(row, COL_SOURCE_FILE, "�����f�[�^")
        Call modCmn.SetRowText(row, COL_NOTES, CStr(sampleData(i)(9)))
    Next i
    
    Call modCmn.LogInfo("CreateInitialSampleData", "�T���v���f�[�^�쐬����")
End Sub

'=============================================================================
' �f�[�^�ϊ����[�e�B���e�B
'=============================================================================

' ����������S�ɔz��ɕϊ�
Public Function SafeSplitString(ByVal inputString As String, ByVal delimiter As String) As Variant
    On Error Resume Next
    
    If Len(inputString) = 0 Then
        SafeSplitString = Array()
    Else
        SafeSplitString = Split(inputString, delimiter)
    End If
End Function

' �z������S�ɕ�����ɕϊ�
Public Function SafeJoinArray(ByVal inputArray As Variant, ByVal delimiter As String) As String
    On Error Resume Next
    
    If IsArray(inputArray) Then
        SafeJoinArray = Join(inputArray, delimiter)
    Else
        SafeJoinArray = ""
    End If
End Function

' CSV�s���S�p�[�X
Public Function SafeParseCsvLine(ByVal csvLine As String) As Variant
    On Error Resume Next
    
    Dim fields As Variant
    Dim i As Integer
    
    ' ��{�I��CSV�����i���p�������Ȃ��j
    fields = Split(csvLine, CSV_DELIMITER)
    
    ' �e�t�B�[���h���g����
    For i = 0 To UBound(fields)
        fields(i) = modCmn.TrimAll(CStr(fields(i)))
        ' ���p������
        If Left(fields(i), 1) = CSV_QUOTE_CHAR And Right(fields(i), 1) = CSV_QUOTE_CHAR Then
            If Len(fields(i)) > 1 Then
                fields(i) = Mid(fields(i), 2, Len(fields(i)) - 2)
            Else
                fields(i) = ""
            End If
        End If
    Next i
    
    SafeParseCsvLine = fields
End Function

'=============================================================================
' �p�t�H�[�}���X�œK�����[�e�B���e�B
'=============================================================================

' �p�t�H�[�}���X�œK���J�n
Public Sub StartPerformanceOptimization()
    On Error Resume Next
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
End Sub

' �p�t�H�[�}���X�œK���I��
Public Sub EndPerformanceOptimization()
    On Error Resume Next
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub

' �������N���[���A�b�v
Public Sub CleanupMemory()
    On Error Resume Next
    
    ' �K�x�[�W�R���N�V�������s�i�\�ȏꍇ�j
    DoEvents
    
    ' �ꎞ�ϐ��N���A
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "Temp*" Then ws.Delete
    Next ws
    
    Call modCmn.LogInfo("CleanupMemory", "�������N���[���A�b�v���s")
End Sub

'=============================================================================
' �V�X�e���f�f���[�e�B���e�B
'=============================================================================

' �V�X�e�����S���`�F�b�N
Public Function PerformSystemHealthCheck() As Object
    On Error Resume Next
    
    Dim result As Object
    Set PerformSystemHealthCheck = CreateObject("Scripting.Dictionary")
    Set result = PerformSystemHealthCheck
    
    result("OverallHealth") = "Healthy"
    result("Issues") = New Collection
    result("Warnings") = New Collection
    
    ' �V�[�g���݃`�F�b�N
    Dim i As Integer
    For i = 1 To 6
        If modCmn.GetWorksheetByIndex(i) Is Nothing Then
            result("Issues").Add "�K�v�ȃV�[�g��������܂���: �C���f�b�N�X " & i
            result("OverallHealth") = "Critical"
        End If
    Next i
    
    ' �e�[�u�����݃`�F�b�N
    If modData.GetCustomersTable() Is Nothing Then
        result("Issues").Add "Customers�e�[�u����������܂���"
        result("OverallHealth") = "Critical"
    End If
    
    ' �ݒ�l�`�F�b�N
    If Len(modData.GetConfigValue(CONFIG_CSV_DIR)) = 0 Then
        result("Warnings").Add "CSV�f�B���N�g�����ݒ肳��Ă��܂���"
        If result("OverallHealth") = "Healthy" Then result("OverallHealth") = "Warning"
    End If
    
    ' �f�B�X�N�e�ʃ`�F�b�N�i�ȈՁj
    If modCmn.DirectoryExists(modData.GetConfigValue(CONFIG_CSV_DIR)) = False Then
        result("Warnings").Add "CSV�f�B���N�g���ɃA�N�Z�X�ł��܂���"
        If result("OverallHealth") = "Healthy" Then result("OverallHealth") = "Warning"
    End If
End Function

' �V�X�e�����擾
Public Function GetSystemInformation() As Object
    On Error Resume Next
    
    Dim info As Object
    Set GetSystemInformation = CreateObject("Scripting.Dictionary")
    Set info = GetSystemInformation
    
    info("SystemName") = SYSTEM_NAME
    info("Version") = SYSTEM_VERSION
    info("Author") = SYSTEM_AUTHOR
    info("CurrentTime") = modCmn.GetCurrentDateTimeString()
    info("WorkbookPath") = ThisWorkbook.FullName
    info("WorkbookName") = ThisWorkbook.Name
    info("ExcelVersion") = Application.Version
    info("SheetCount") = ThisWorkbook.Worksheets.Count
    
    ' �e�[�u�����v
    Dim customerTbl As ListObject
    Set customerTbl = modData.GetCustomersTable()
    If Not customerTbl Is Nothing Then
        info("CustomerCount") = customerTbl.ListRows.Count
    Else
        info("CustomerCount") = 0
    End If
    
    Dim stagingTbl As ListObject
    Set stagingTbl = modData.GetStagingTable()
    If Not stagingTbl Is Nothing Then
        info("StagingCount") = stagingTbl.ListRows.Count
    Else
        info("StagingCount") = 0
    End If
End Function

'=============================================================================
' �f�o�b�O�E�J���x�����[�e�B���e�B
'=============================================================================

' �V�X�e����ԃ_���v�i�f�o�b�O�p�j
Public Sub DumpSystemState()
    On Error Resume Next
    
    Debug.Print "=== �V�X�e����ԃ_���v ==="
    Debug.Print "����: " & modCmn.GetCurrentDateTimeString()
    
    ' �V�X�e�����
    Dim info As Object
    Set info = GetSystemInformation()
    
    Dim key As Variant
    For Each key In info.Keys
        Debug.Print key & ": " & info(key)
    Next key
    
    ' ���S���`�F�b�N
    Dim health As Object
    Set health = PerformSystemHealthCheck()
    Debug.Print "�V�X�e�����S��: " & health("OverallHealth")
    
    Debug.Print "=== �_���v�I�� ==="
End Sub

' �e�X�g�f�[�^�����i�J���p�j
Public Sub GenerateTestData()
    On Error Resume Next
    
    Dim customerTbl As ListObject
    Set customerTbl = modData.GetCustomersTable()
    If customerTbl Is Nothing Then Exit Sub
    
    ' �e�X�g�f�[�^100������
    Dim i As Integer
    For i = 1 To 100
        Dim row As ListRow
        Set row = customerTbl.ListRows.Add
        
        Call modCmn.SetRowText(row, COL_CUSTOMER_ID, "TEST" & Format(i, "000"))
        Call modCmn.SetRowText(row, COL_CUSTOMER_NAME, "�e�X�g�ڋq" & i)
        Call modCmn.SetRowText(row, COL_EMAIL, "test" & i & "@example.com")
        Call modCmn.SetRowText(row, COL_PHONE, "03-" & Format(i, "0000") & "-0000")
        Call modCmn.SetRowText(row, COL_ZIP, "100-000" & (i Mod 10))
        Call modCmn.SetRowText(row, COL_ADDRESS1, "�����s���c��")
        Call modCmn.SetRowText(row, COL_ADDRESS2, "�e�X�g" & i & "-1-1")
        Call modCmn.SetRowText(row, COL_CATEGORY, IIf(i Mod 2 = 0, CATEGORY_B2B, CATEGORY_B2C))
        Call modCmn.SetRowText(row, COL_STATUS, IIf(i Mod 10 = 0, STATUS_INACTIVE, STATUS_ACTIVE))
        Call modCmn.SetRowDate(row, COL_CREATED_AT, Now - i)
        Call modCmn.SetRowDate(row, COL_UPDATED_AT, Now)
        Call modCmn.SetRowText(row, COL_SOURCE_FILE, "�e�X�g�f�[�^")
        Call modCmn.SetRowText(row, COL_NOTES, "���������e�X�g�f�[�^")
    Next i
    
    Call modCmn.LogInfo("GenerateTestData", "�e�X�g�f�[�^100����������")
End Sub