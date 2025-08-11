Attribute VB_Name = "modCustomerSystem"
'=============================================================================
' modCustomerSystem.bas - �ڋq�Ǘ��V�X�e�����C�����䃂�W���[��
'=============================================================================
' �T�v:
'   �V�X�e���S�̂̃I�[�P�X�g���[�V�����A���C�������t���[�̐���
'   �e���W���[���Ԃ̘A�g�A�G���[�n���h�����O�A�g�����U�N�V�����Ǘ�
'=============================================================================
Option Explicit

'=============================================================================
' ���C�������t���[
'=============================================================================

' CSV�ꊇ�捞�E�X�V�����i���C�������j
Public Sub ExecuteFullImportProcess()
    On Error GoTo ErrHandler
    
    Dim startTime As Double
    Dim result As Object
    Dim totalProcessTime As Double
    
    startTime = Timer
    
    ' �����m�F
    If Not modDashboard.ConfirmOperation("CSV�ꊇ�捞�E�X�V���������s���܂��B" & vbCrLf & _
                                       "1. CSV�t�@�C���捞" & vbCrLf & _
                                       "2. �f�[�^���K���E����" & vbCrLf & _
                                       "3. �ڋq�f�[�^�X�V�i�ǉ��E�X�V�j" & vbCrLf & _
                                       "4. KPI�\���X�V") Then
        Call modDashboard.ShowStatusMessage("�������L�����Z������܂���")
        Exit Sub
    End If
    
    ' �p�t�H�[�}���X�œK���J�n
    Call modUtils.StartPerformanceOptimization()
    
    ' �o�b�N�A�b�v�쐬
    Call modDashboard.ShowStatusMessage("�o�b�N�A�b�v�쐬��...")
    If Not modUpsert.BackupCustomerData() Then
        Call modDashboard.ShowStatusMessage("�o�b�N�A�b�v�쐬�Ɏ��s���܂������������p�����܂�", isError:=True)
    End If
    
    ' �X�e�b�v1: CSV�捞
    Call modDashboard.ShowStatusMessage("CSV�捞������...")
    Call modData.ImportCsvToStaging()
    
    ' �X�e�b�v2: �f�[�^����
    Call modDashboard.ShowStatusMessage("�f�[�^���؏�����...")
    Dim validationErrors As Long
    validationErrors = modValidation.ValidateStagingData()
    
    If validationErrors < 0 Then
        Err.Raise vbObjectError + 1001, "ExecuteFullImportProcess", "�f�[�^���؏����ŃG���[���������܂���"
    End If
    
    ' �X�e�b�v3: �A�b�v�T�[�g����
    Call modDashboard.ShowStatusMessage("�f�[�^�X�V������...")
    Set result = modUpsert.ExecuteUpsertOperation()
    
    If Not result("Success") Then
        Err.Raise vbObjectError + 1002, "ExecuteFullImportProcess", result("Message")
    End If
    
    ' �X�e�b�v4: �����؂�ڋq�������i�C�Ӂj
    Dim inactivatedCount As Long
    If LCase(modData.GetConfigValue("AUTO_INACTIVATE")) = "true" Then
        Call modDashboard.ShowStatusMessage("�����؂�ڋq��������...")
        inactivatedCount = modUpsert.InactivateStaleCustomers()
    End If
    
    ' �X�e�b�v5: KPI�X�V
    Call modDashboard.ShowStatusMessage("KPI�\���X�V��...")
    Call modDashboard.RefreshKPI()
    
    ' ��������
    totalProcessTime = Timer - startTime
    
    ' ���ʕ\��
    Dim resultMsg As String
    resultMsg = "CSV�ꊇ�捞�E�X�V�������������܂����B" & vbCrLf & vbCrLf
    resultMsg = resultMsg & "�y�������ʁz" & vbCrLf
    resultMsg = resultMsg & "�ǉ�����: " & Format(result("AddedCount"), NUMBER_FORMAT_COUNT) & " ��" & vbCrLf
    resultMsg = resultMsg & "�X�V����: " & Format(result("UpdatedCount"), NUMBER_FORMAT_COUNT) & " ��" & vbCrLf
    resultMsg = resultMsg & "�X�L�b�v����: " & Format(result("SkippedCount"), NUMBER_FORMAT_COUNT) & " ��" & vbCrLf
    resultMsg = resultMsg & "���؃G���[����: " & Format(validationErrors, NUMBER_FORMAT_COUNT) & " ��" & vbCrLf
    If inactivatedCount > 0 Then
        resultMsg = resultMsg & "����������: " & Format(inactivatedCount, NUMBER_FORMAT_COUNT) & " ��" & vbCrLf
    End If
    resultMsg = resultMsg & "����������: " & Format(totalProcessTime, "0.0") & " �b"
    
    ' ���O�L�^
    Call modData.LogImportOperation("�t����������", result("AddedCount") + result("UpdatedCount"), _
                                   totalProcessTime, "���؃G���[:" & validationErrors & ", ������:" & inactivatedCount)
    
    MsgBox resultMsg, vbInformation, "��������"
    Call modDashboard.ShowStatusMessage("��������: " & Format(Now, "hh:mm"))
    
ExitHandler:
    Call modUtils.EndPerformanceOptimization()
    Exit Sub
    
ErrHandler:
    Call modUtils.EndPerformanceOptimization()
    Call modDashboard.ShowStatusMessage("�����G���[���������܂���", isError:=True)
    Call modCmn.LogError("ExecuteFullImportProcess", "�t�������G���[: " & Err.Description)
    MsgBox "�������ɃG���[���������܂����B" & vbCrLf & vbCrLf & _
           "�G���[���e: " & Err.Description & vbCrLf & vbCrLf & _
           "�ڍׂ̓��O�t�@�C�������m�F���������B", vbCritical, "�����G���["
End Sub

' CSV�捞�̂ݎ��s
Public Sub ExecuteImportOnly()
    On Error GoTo ErrHandler
    
    If Not modDashboard.ConfirmOperation("CSV�t�@�C���̎捞�݂̂����s���܂��B") Then Exit Sub
    
    Call modUtils.StartPerformanceOptimization()
    Call modData.ImportCsvToStaging()
    Call modUtils.EndPerformanceOptimization()
    
    Call modDashboard.RefreshKPI()
    MsgBox "CSV�捞�������������܂����B", vbInformation
    Exit Sub
    
ErrHandler:
    Call modUtils.EndPerformanceOptimization()
    Call modCmn.LogError("ExecuteImportOnly", Err.Description)
End Sub

' �f�[�^���؂̂ݎ��s
Public Sub ExecuteValidationOnly()
    On Error GoTo ErrHandler
    
    Dim errorCount As Long
    
    Call modUtils.StartPerformanceOptimization()
    errorCount = modValidation.ValidateStagingData()
    Call modUtils.EndPerformanceOptimization()
    
    If errorCount >= 0 Then
        MsgBox "�f�[�^���؂��������܂����B" & vbCrLf & _
               "�G���[����: " & Format(errorCount, NUMBER_FORMAT_COUNT) & " ��", vbInformation
    Else
        MsgBox "�f�[�^���؏����ŃG���[���������܂����B", vbCritical
    End If
    
    Call modDashboard.RefreshKPI()
    Exit Sub
    
ErrHandler:
    Call modUtils.EndPerformanceOptimization()
    Call modCmn.LogError("ExecuteValidationOnly", Err.Description)
End Sub

' �A�b�v�T�[�g�̂ݎ��s
Public Sub ExecuteUpsertOnly()
    On Error GoTo ErrHandler
    
    Dim result As Object
    
    If Not modDashboard.ConfirmOperation("Staging�f�[�^����ڋq�e�[�u���̍X�V�݂̂����s���܂��B") Then Exit Sub
    
    Call modUtils.StartPerformanceOptimization()
    Set result = modUpsert.ExecuteUpsertOperation()
    Call modUtils.EndPerformanceOptimization()
    
    If result("Success") Then
        MsgBox "�A�b�v�T�[�g�������������܂����B" & vbCrLf & _
               "�ǉ�: " & result("AddedCount") & " ��, �X�V: " & result("UpdatedCount") & " ��", vbInformation
    Else
        MsgBox "�A�b�v�T�[�g�����ŃG���[���������܂����B" & vbCrLf & result("Message"), vbCritical
    End If
    
    Call modDashboard.RefreshKPI()
    Exit Sub
    
ErrHandler:
    Call modUtils.EndPerformanceOptimization()
    Call modCmn.LogError("ExecuteUpsertOnly", Err.Description)
End Sub

'=============================================================================
' �V�X�e���Ǘ�����
'=============================================================================

' �V�X�e������������
Public Sub InitializeSystem()
    On Error GoTo ErrHandler
    
    Call modCmn.LogInfo("InitializeSystem", "�V�X�e���������J�n")
    
    ' �V�X�e���S�̏�����
    Call modUtils.InitializeCustomerSystem()
    
    ' ������������̃X�v���b�V���\��
    Call ShowSystemSplash()
    
    Call modCmn.LogInfo("InitializeSystem", "�V�X�e������������")
    Exit Sub
    
ErrHandler:
    Call modCmn.LogError("InitializeSystem", "�V�X�e���������G���[: " & Err.Description)
End Sub

' �V�X�e���X�v���b�V���\��
Private Sub ShowSystemSplash()
    On Error Resume Next
    
    Dim splashMsg As String
    splashMsg = SYSTEM_NAME & vbCrLf
    splashMsg = splashMsg & "�o�[�W���� " & SYSTEM_VERSION & vbCrLf & vbCrLf
    splashMsg = splashMsg & "�V�X�e���̏��������������܂����B" & vbCrLf & vbCrLf
    splashMsg = splashMsg & "�g�p���@:" & vbCrLf
    splashMsg = splashMsg & "1. Dashboard�V�[�g�Ŋe�푀������s" & vbCrLf
    splashMsg = splashMsg & "2. _Config�V�[�g�Őݒ�𒲐�" & vbCrLf
    splashMsg = splashMsg & "3. Logs�V�[�g�ŏ����������m�F" & vbCrLf & vbCrLf
    splashMsg = splashMsg & SYSTEM_COPYRIGHT
    
    MsgBox splashMsg, vbInformation, "�V�X�e���N������"
End Sub

' �V�X�e�����S���`�F�b�N���s
Public Sub ExecuteSystemHealthCheck()
    On Error GoTo ErrHandler
    
    Dim health As Object
    Dim healthMsg As String
    Dim issue As Variant
    Dim warning As Variant
    
    Set health = modUtils.PerformSystemHealthCheck()
    
    healthMsg = "=== �V�X�e�����S���`�F�b�N���� ===" & vbCrLf & vbCrLf
    healthMsg = healthMsg & "��������: " & health("OverallHealth") & vbCrLf & vbCrLf
    
    ' �d��Ȗ��
    If health("Issues").Count > 0 Then
        healthMsg = healthMsg & "�y�d��Ȗ��z" & vbCrLf
        For Each issue In health("Issues")
            healthMsg = healthMsg & "- " & issue & vbCrLf
        Next issue
        healthMsg = healthMsg & vbCrLf
    End If
    
    ' �x��
    If health("Warnings").Count > 0 Then
        healthMsg = healthMsg & "�y�x���z" & vbCrLf
        For Each warning In health("Warnings")
            healthMsg = healthMsg & "- " & warning & vbCrLf
        Next warning
        healthMsg = healthMsg & vbCrLf
    End If
    
    If health("Issues").Count = 0 And health("Warnings").Count = 0 Then
        healthMsg = healthMsg & "���͌��o����܂���ł����B"
    End If
    
    MsgBox healthMsg, vbInformation, "�V�X�e�����S���`�F�b�N"
    Exit Sub
    
ErrHandler:
    Call modCmn.LogError("ExecuteSystemHealthCheck", Err.Description)
End Sub

' �V�X�e�����\��
Public Sub ShowSystemInformation()
    On Error Resume Next
    
    Dim info As Object
    Dim infoMsg As String
    Dim key As Variant
    
    Set info = modUtils.GetSystemInformation()
    
    infoMsg = "=== �V�X�e����� ===" & vbCrLf & vbCrLf
    
    ' ��{���
    infoMsg = infoMsg & "�V�X�e����: " & info("SystemName") & vbCrLf
    infoMsg = infoMsg & "�o�[�W����: " & info("Version") & vbCrLf
    infoMsg = infoMsg & "�쐬��: " & info("Author") & vbCrLf
    infoMsg = infoMsg & "���ݓ���: " & info("CurrentTime") & vbCrLf
    infoMsg = infoMsg & "Excel�o�[�W����: " & info("ExcelVersion") & vbCrLf & vbCrLf
    
    ' �t�@�C�����
    infoMsg = infoMsg & "���[�N�u�b�N��: " & info("WorkbookName") & vbCrLf
    infoMsg = infoMsg & "�t�@�C���p�X: " & info("WorkbookPath") & vbCrLf
    infoMsg = infoMsg & "�V�[�g��: " & info("SheetCount") & vbCrLf & vbCrLf
    
    ' �f�[�^���v
    infoMsg = infoMsg & "�ڋq�f�[�^����: " & Format(info("CustomerCount"), NUMBER_FORMAT_COUNT) & " ��" & vbCrLf
    infoMsg = infoMsg & "Staging�f�[�^����: " & Format(info("StagingCount"), NUMBER_FORMAT_COUNT) & " ��" & vbCrLf
    
    MsgBox infoMsg, vbInformation, "�V�X�e�����"
End Sub

'=============================================================================
' �f�[�^�����e�i���X����
'=============================================================================

' �S�f�[�^�N���A�i�J���E�e�X�g�p�j
Public Sub ClearAllData()
    On Error GoTo ErrHandler
    
    Dim confirmMsg As String
    confirmMsg = "�S�Ẵf�[�^���N���A���܂��B" & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "���̑���ɂ��ȉ��̃f�[�^���폜����܂�:" & vbCrLf
    confirmMsg = confirmMsg & "- �ڋq�}�X�^�f�[�^" & vbCrLf
    confirmMsg = confirmMsg & "- Staging�f�[�^" & vbCrLf
    confirmMsg = confirmMsg & "- ���O�f�[�^" & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "���ݒ�f�[�^�͕ێ�����܂�" & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "�{���Ɏ��s���܂����H"
    
    If MsgBox(confirmMsg, vbQuestion + vbYesNo + vbDefaultButton2, "�f�[�^�N���A�m�F") <> vbYes Then
        Exit Sub
    End If
    
    ' 2��ڂ̊m�F
    If MsgBox("�ŏI�m�F�F�S�f�[�^���N���A���܂����H" & vbCrLf & "���̑���͎��������Ƃ��ł��܂���B", _
              vbCritical + vbYesNo + vbDefaultButton2, "�ŏI�m�F") <> vbYes Then
        Exit Sub
    End If
    
    Call modUtils.StartPerformanceOptimization()
    
    ' �ڋq�f�[�^�N���A
    Dim ws As Worksheet
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_CUSTOMERS)
    If Not ws Is Nothing Then
        Call modCmn.SafeClearSheet(ws, keepFormats:=True)
        Call modUtils.EnsureCustomersTableStructure()
    End If
    
    ' Staging�f�[�^�N���A
    Call modData.ClearStagingData()
    
    ' ���O�f�[�^�N���A
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_LOGS)
    If Not ws Is Nothing Then
        Call modCmn.SafeClearSheet(ws, keepFormats:=True)
        Call modUtils.EnsureLogsTableStructure()
    End If
    
    Call modUtils.EndPerformanceOptimization()
    
    ' KPI�X�V
    Call modDashboard.RefreshKPI()
    
    Call modCmn.LogInfo("ClearAllData", "�S�f�[�^�N���A���s")
    MsgBox "�S�f�[�^�̃N���A���������܂����B", vbInformation
    Exit Sub
    
ErrHandler:
    Call modUtils.EndPerformanceOptimization()
    Call modCmn.LogError("ClearAllData", Err.Description)
End Sub

' �f�[�^�x�[�X�œK��
Public Sub OptimizeDatabase()
    On Error GoTo ErrHandler
    
    If Not modDashboard.ConfirmOperation("�f�[�^�x�[�X�̍œK�������s���܂��B" & vbCrLf & _
                                       "�E�Â����O�̍폜" & vbCrLf & _
                                       "�E�d�����R�[�h�̃N���[���A�b�v" & vbCrLf & _
                                       "�E�e�[�u���\���̍œK��") Then Exit Sub
    
    Call modUtils.StartPerformanceOptimization()
    
    Dim optimizedCount As Long
    optimizedCount = 0
    
    ' �Â����O�폜�i30���ȑO�j
    Call CleanupOldLogs(30, optimizedCount)
    
    ' �d�����R�[�h�`�F�b�N�i�x���̂݁j
    Call CheckForDuplicateCustomers()
    
    ' �e�[�u���\���œK��
    Call modUtils.EnsureAllTableStructures()
    
    ' �������N���[���A�b�v
    Call modUtils.CleanupMemory()
    
    Call modUtils.EndPerformanceOptimization()
    
    MsgBox "�f�[�^�x�[�X�œK�����������܂����B" & vbCrLf & _
           "��������: " & Format(optimizedCount, NUMBER_FORMAT_COUNT) & " ��", vbInformation
    Exit Sub
    
ErrHandler:
    Call modUtils.EndPerformanceOptimization()
    Call modCmn.LogError("OptimizeDatabase", Err.Description)
End Sub

' �Â����O�N���[���A�b�v
Private Sub CleanupOldLogs(ByVal daysToKeep As Integer, ByRef cleanedCount As Long)
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    Dim cutoffDate As Date
    Dim i As Long
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_LOGS)
    If ws Is Nothing Then Exit Sub
    
    Set tbl = modCmn.GetTable(ws, TABLE_LOGS)
    If tbl Is Nothing Then Exit Sub
    
    cutoffDate = Now - daysToKeep
    
    ' ��납��폜�i�C���f�b�N�X�ύX������邽�߁j
    For i = tbl.ListRows.Count To 1 Step -1
        Set row = tbl.ListRows(i)
        Dim logDate As Date
        logDate = modCmn.SafeDate(modCmn.GetRowText(row, "Timestamp"))
        
        If logDate > 0 And logDate < cutoffDate Then
            row.Delete
            cleanedCount = cleanedCount + 1
        End If
    Next i
End Sub

' �d���ڋq�`�F�b�N
Private Sub CheckForDuplicateCustomers()
    On Error Resume Next
    
    Dim tbl As ListObject
    Set tbl = modData.GetCustomersTable()
    If tbl Is Nothing Then Exit Sub
    
    Dim duplicates As Collection
    Set duplicates = New Collection
    
    ' �d�����o���W�b�N�i�ȈՔŁj
    ' �����̏ڍׂ͏ȗ�
    
    If duplicates.Count > 0 Then
        MsgBox "�d���̉\��������ڋq�f�[�^�� " & duplicates.Count & " ��������܂����B" & vbCrLf & _
               "�ڍׂ�Logs�V�[�g�����m�F���������B", vbExclamation
    End If
End Sub

'=============================================================================
' �G�N�X�|�[�g�E���|�[�g����
'=============================================================================

' �ڋq�f�[�^�G�N�X�|�[�g
Public Sub ExportCustomerData()
    On Error GoTo ErrHandler
    
    Dim exportPath As String
    Dim fileName As String
    
    fileName = "customer_export_" & Format(Now, DATE_FORMAT_FILE) & ".csv"
    exportPath = ThisWorkbook.Path & "\" & fileName
    
    Call modUtils.StartPerformanceOptimization()
    
    ' CSV�o�͏���
    If ExportCustomerTableToCsv(exportPath) Then
        Call modUtils.EndPerformanceOptimization()
        MsgBox "�ڋq�f�[�^�̃G�N�X�|�[�g���������܂����B" & vbCrLf & exportPath, vbInformation
        Call modCmn.LogInfo("ExportCustomerData", "�G�N�X�|�[�g����: " & exportPath)
    Else
        Call modUtils.EndPerformanceOptimization()
        MsgBox "�G�N�X�|�[�g�����ŃG���[���������܂����B", vbCritical
    End If
    Exit Sub
    
ErrHandler:
    Call modUtils.EndPerformanceOptimization()
    Call modCmn.LogError("ExportCustomerData", Err.Description)
End Sub

' �ڋq�e�[�u��CSV�o��
Private Function ExportCustomerTableToCsv(ByVal filePath As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim tbl As ListObject
    Dim fileNum As Integer
    Dim row As ListRow
    Dim csvLine As String
    Dim exportCount As Long
    
    ExportCustomerTableToCsv = False
    
    Set tbl = modData.GetCustomersTable()
    If tbl Is Nothing Then Exit Function
    
    fileNum = FreeFile
    Open filePath For Output As fileNum
    
    ' �w�b�_�[�s�o��
    Print #fileNum, CUSTOMERS_HEADERS
    
    ' �f�[�^�s�o��
    For Each row In tbl.ListRows
        csvLine = ""
        csvLine = csvLine & QuoteCsvField(modCmn.GetRowText(row, COL_CUSTOMER_ID)) & ","
        csvLine = csvLine & QuoteCsvField(modCmn.GetRowText(row, COL_CUSTOMER_NAME)) & ","
        csvLine = csvLine & QuoteCsvField(modCmn.GetRowText(row, COL_EMAIL)) & ","
        csvLine = csvLine & QuoteCsvField(modCmn.GetRowText(row, COL_PHONE)) & ","
        csvLine = csvLine & QuoteCsvField(modCmn.GetRowText(row, COL_ZIP)) & ","
        csvLine = csvLine & QuoteCsvField(modCmn.GetRowText(row, COL_ADDRESS1)) & ","
        csvLine = csvLine & QuoteCsvField(modCmn.GetRowText(row, COL_ADDRESS2)) & ","
        csvLine = csvLine & QuoteCsvField(modCmn.GetRowText(row, COL_CATEGORY)) & ","
        csvLine = csvLine & QuoteCsvField(modCmn.GetRowText(row, COL_STATUS)) & ","
        csvLine = csvLine & QuoteCsvField(CStr(modCmn.GetRowDate(row, COL_CREATED_AT))) & ","
        csvLine = csvLine & QuoteCsvField(CStr(modCmn.GetRowDate(row, COL_UPDATED_AT))) & ","
        csvLine = csvLine & QuoteCsvField(modCmn.GetRowText(row, COL_SOURCE_FILE)) & ","
        csvLine = csvLine & QuoteCsvField(modCmn.GetRowText(row, COL_NOTES))
        
        Print #fileNum, csvLine
        exportCount = exportCount + 1
    Next row
    
    Close fileNum
    ExportCustomerTableToCsv = True
    
    Call modCmn.LogInfo("ExportCustomerTableToCsv", "CSV�o�͊���: " & exportCount & " ��")
    Exit Function
    
ErrHandler:
    If fileNum > 0 Then Close fileNum
    ExportCustomerTableToCsv = False
    Call modCmn.LogError("ExportCustomerTableToCsv", Err.Description)
End Function

' CSV�t�B�[���h���p������
Private Function QuoteCsvField(ByVal field As String) As String
    On Error Resume Next
    
    ' �J���}����s���܂܂��ꍇ�͈��p���ň͂�
    If InStr(field, ",") > 0 Or InStr(field, vbCrLf) > 0 Or InStr(field, """") > 0 Then
        ' ���p�����G�X�P�[�v
        field = Replace(field, """", """""")
        QuoteCsvField = """" & field & """"
    Else
        QuoteCsvField = field
    End If
End Function

'=============================================================================
' �G���[���J�o������
'=============================================================================

' �V�X�e���ً}��~
Public Sub EmergencyStop()
    On Error Resume Next
    
    ' �S�Ă̏������~
    Application.EnableEvents = False
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Call modCmn.LogError("EmergencyStop", "�V�X�e���ً}��~�����s����܂���")
    
    MsgBox "�V�X�e���ً̋}��~�����s���܂����B" & vbCrLf & _
           "�������ĊJ����ꍇ�́A���[�N�u�b�N���ēx�J���Ă��������B", vbCritical, "�ً}��~"
End Sub