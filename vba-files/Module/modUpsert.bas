Attribute VB_Name = "modUpsert"
'=============================================================================
' modUpsert.bas - �A�b�v�T�[�g�iInsert/Update�j�������W���[��
'=============================================================================
' �T�v:
'   Staging����ڋq�e�[�u���ւ̈��S�Ȓǉ��E�X�V����
'   �_���폜�A�����X�V�A�d���������̃f�[�^�����@�\���
'=============================================================================
Option Explicit

'=============================================================================
' ���C���A�b�v�T�[�g����
'=============================================================================

' Staging����ڋq�e�[�u���ւ̃A�b�v�T�[�g���s
Public Function ExecuteUpsertOperation() As Object
    On Error GoTo ErrHandler
    
    Dim result As Object
    Dim stagingTbl As ListObject
    Dim customerTbl As ListObject
    Dim stagingRow As ListRow
    Dim startTime As Double
    Dim processedCount As Long
    Dim addedCount As Long
    Dim updatedCount As Long
    Dim skippedCount As Long
    
    startTime = Timer
    Set result = CreateObject("Scripting.Dictionary")
    
    ' ������
    result("Success") = False
    result("AddedCount") = 0
    result("UpdatedCount") = 0
    result("SkippedCount") = 0
    result("ProcessTime") = 0
    result("Message") = ""
    
    ' �e�[�u���擾
    Set stagingTbl = modData.GetStagingTable()
    Set customerTbl = modData.GetCustomersTable()
    
    If stagingTbl Is Nothing Or customerTbl Is Nothing Then
        result("Message") = "�K�v�ȃe�[�u�����擾�ł��܂���ł���"
        Set ExecuteUpsertOperation = result
        Exit Function
    End If
    
    Call modCmn.ShowProgressStart(MSG_UPSERT_STARTED)
    
    ' �ڋq�����쐬�i���������p�j
    Dim customerDict As Object
    Set customerDict = CreateCustomerSearchDictionary(customerTbl)
    
    ' �eStaging�s������
    For Each stagingRow In stagingTbl.ListRows
        processedCount = processedCount + 1
        
        ' �L���ȃ��R�[�h�̂ݏ���
        If modCmn.GetRowValue(stagingRow, COL_IS_VALID) = True Then
            Dim upsertResult As Integer
            upsertResult = ProcessSingleUpsert(stagingRow, customerTbl, customerDict)
            
            Select Case upsertResult
                Case 1 ' �ǉ�
                    addedCount = addedCount + 1
                Case 2 ' �X�V
                    updatedCount = updatedCount + 1
                Case 0 ' �X�L�b�v
                    skippedCount = skippedCount + 1
            End Select
        Else
            skippedCount = skippedCount + 1
        End If
        
        ' �v���O���X�X�V
        If processedCount Mod BATCH_SIZE_UPSERT = 0 Then
            Call modCmn.UpdateProgress("�f�[�^�X�V��: " & processedCount & " ������")
        End If
    Next stagingRow
    
    ' ���ʐݒ�
    result("Success") = True
    result("AddedCount") = addedCount
    result("UpdatedCount") = updatedCount
    result("SkippedCount") = skippedCount
    result("ProcessTime") = Timer - startTime
    result("Message") = "�A�b�v�T�[�g����: �ǉ�" & addedCount & "��, �X�V" & updatedCount & "��"
    
    ' ���O�L�^
    Call modData.LogImportOperation("�A�b�v�T�[�g��������", processedCount, Timer - startTime, _
                                   "�ǉ�:" & addedCount & ", �X�V:" & updatedCount & ", �X�L�b�v:" & skippedCount)
    
    Set ExecuteUpsertOperation = result
    Call modCmn.HideProgress
    Exit Function
    
ErrHandler:
    Call modCmn.HideProgress
    result("Success") = False
    result("Message") = "�A�b�v�T�[�g�����G���[: " & Err.Description
    Set ExecuteUpsertOperation = result
    Call modCmn.LogError("ExecuteUpsertOperation", Err.Description)
End Function

' �P�ꃌ�R�[�h�̃A�b�v�T�[�g����
Private Function ProcessSingleUpsert(ByVal stagingRow As ListRow, ByVal customerTbl As ListObject, _
                                    ByVal customerDict As Object) As Integer
    On Error Resume Next
    
    Dim primaryKey As String
    Dim altKey As String
    Dim existingRow As ListRow
    
    ProcessSingleUpsert = 0 ' �f�t�H���g�F�X�L�b�v
    
    ' �L�[���擾
    primaryKey = modCmn.GetRowText(stagingRow, COL_CUSTOMER_ID)
    altKey = modCmn.GetRowText(stagingRow, COL_KEY_CANDIDATE)
    
    ' �������R�[�h����
    Set existingRow = FindExistingCustomer(customerTbl, customerDict, primaryKey, altKey)
    
    If existingRow Is Nothing Then
        ' �V�K�ǉ�
        If AddNewCustomer(stagingRow, customerTbl) Then
            ProcessSingleUpsert = 1 ' �ǉ�����
        End If
    Else
        ' �����X�V
        If UpdateExistingCustomer(stagingRow, existingRow) Then
            ProcessSingleUpsert = 2 ' �X�V����
        End If
    End If
End Function

'=============================================================================
' �ڋq�����E�����Ǘ�
'=============================================================================

' �ڋq���������쐬
Private Function CreateCustomerSearchDictionary(ByVal customerTbl As ListObject) As Object
    On Error Resume Next
    
    Dim customerDict As Object
    Dim row As ListRow
    Dim primaryKey As String
    Dim email As String
    Dim customerName As String
    Dim altKey As String
    
    Set CreateCustomerSearchDictionary = CreateObject("Scripting.Dictionary")
    Set customerDict = CreateCustomerSearchDictionary
    
    Dim rowIndex As Long
    rowIndex = 1
    
    For Each row In customerTbl.ListRows
        ' ��L�[�}�b�s���O�i�ڋqID�j
        primaryKey = modCmn.GetRowText(row, COL_CUSTOMER_ID)
        If Len(primaryKey) > 0 Then
            customerDict("PK:" & primaryKey) = rowIndex
        End If
        
        ' ��փL�[�}�b�s���O�iEmail + CustomerName�j
        email = modCmn.NormalizeEmail(modCmn.GetRowText(row, COL_EMAIL))
        customerName = modCmn.GetRowText(row, COL_CUSTOMER_NAME)
        If Len(email) > 0 And Len(customerName) > 0 Then
            altKey = email & "+" & customerName
            customerDict("AK:" & altKey) = rowIndex
        End If
        
        rowIndex = rowIndex + 1
    Next row
End Function

' �����ڋq����
Private Function FindExistingCustomer(ByVal customerTbl As ListObject, ByVal customerDict As Object, _
                                     ByVal primaryKey As String, ByVal altKey As String) As ListRow
    On Error Resume Next
    
    Dim rowIndex As Variant
    
    Set FindExistingCustomer = Nothing
    
    ' ��L�[�Ō���
    If Len(primaryKey) > 0 Then
        rowIndex = customerDict("PK:" & primaryKey)
        If Not IsEmpty(rowIndex) Then
            Set FindExistingCustomer = customerTbl.ListRows(CLng(rowIndex))
            Exit Function
        End If
    End If
    
    ' ��փL�[�Ō����i��L�[���Ȃ��ꍇ�j
    If Len(primaryKey) = 0 And Len(altKey) > 0 Then
        rowIndex = customerDict("AK:" & altKey)
        If Not IsEmpty(rowIndex) Then
            Set FindExistingCustomer = customerTbl.ListRows(CLng(rowIndex))
        End If
    End If
End Function

'=============================================================================
' �V�K�ǉ�����
'=============================================================================

' �V�K�ڋq�ǉ�
Private Function AddNewCustomer(ByVal stagingRow As ListRow, ByVal customerTbl As ListObject) As Boolean
    On Error GoTo ErrHandler
    
    Dim newRow As ListRow
    
    AddNewCustomer = False
    
    ' �V�K�s�ǉ�
    Set newRow = customerTbl.ListRows.Add
    
    ' �f�[�^�R�s�[
    Call CopyDataFromStaging(stagingRow, newRow, isNewRecord:=True)
    
    AddNewCustomer = True
    Exit Function
    
ErrHandler:
    AddNewCustomer = False
    Call modCmn.LogError("AddNewCustomer", "�V�K�ڋq�ǉ��G���[: " & Err.Description)
End Function

'=============================================================================
' �����X�V����
'=============================================================================

' �����ڋq�X�V
Private Function UpdateExistingCustomer(ByVal stagingRow As ListRow, ByVal existingRow As ListRow) As Boolean
    On Error GoTo ErrHandler
    
    UpdateExistingCustomer = False
    
    ' �����`�F�b�N���X�V
    If HasDataDifferences(stagingRow, existingRow) Then
        Call CopyDataFromStaging(stagingRow, existingRow, isNewRecord:=False)
        UpdateExistingCustomer = True
    End If
    
    Exit Function
    
ErrHandler:
    UpdateExistingCustomer = False
    Call modCmn.LogError("UpdateExistingCustomer", "�ڋq�X�V�G���[: " & Err.Description)
End Function

' �f�[�^�����`�F�b�N
Private Function HasDataDifferences(ByVal stagingRow As ListRow, ByVal existingRow As ListRow) As Boolean
    On Error Resume Next
    
    HasDataDifferences = False
    
    ' ��v�t�B�[���h�̍����`�F�b�N
    If CompareFieldValues(stagingRow, existingRow, COL_CUSTOMER_NAME) Then HasDataDifferences = True
    If CompareFieldValues(stagingRow, existingRow, COL_EMAIL) Then HasDataDifferences = True
    If CompareFieldValues(stagingRow, existingRow, COL_PHONE) Then HasDataDifferences = True
    If CompareFieldValues(stagingRow, existingRow, COL_ZIP) Then HasDataDifferences = True
    If CompareFieldValues(stagingRow, existingRow, COL_ADDRESS1) Then HasDataDifferences = True
    If CompareFieldValues(stagingRow, existingRow, COL_ADDRESS2) Then HasDataDifferences = True
    If CompareFieldValues(stagingRow, existingRow, COL_CATEGORY) Then HasDataDifferences = True
    If CompareFieldValues(stagingRow, existingRow, COL_STATUS) Then HasDataDifferences = True
    If CompareFieldValues(stagingRow, existingRow, COL_NOTES) Then HasDataDifferences = True
End Function

' �t�B�[���h�l��r
Private Function CompareFieldValues(ByVal stagingRow As ListRow, ByVal existingRow As ListRow, _
                                   ByVal fieldName As String) As Boolean
    On Error Resume Next
    
    Dim stagingValue As String
    Dim existingValue As String
    
    ' ���K���ςݒl���g�p�i���p�\�ȏꍇ�j
    Select Case fieldName
        Case COL_EMAIL
            stagingValue = modCmn.TrimAll(modCmn.GetRowText(stagingRow, COL_EMAIL_NORM))
            existingValue = modCmn.NormalizeEmail(modCmn.GetRowText(existingRow, COL_EMAIL))
        Case COL_PHONE
            stagingValue = modCmn.TrimAll(modCmn.GetRowText(stagingRow, COL_PHONE_NORM))
            existingValue = modCmn.NormalizePhone(modCmn.GetRowText(existingRow, COL_PHONE))
        Case COL_ZIP
            stagingValue = modCmn.TrimAll(modCmn.GetRowText(stagingRow, COL_ZIP_NORM))
            existingValue = modCmn.NormalizeZip(modCmn.GetRowText(existingRow, COL_ZIP))
        Case Else
            stagingValue = modCmn.TrimAll(modCmn.GetRowText(stagingRow, fieldName))
            existingValue = modCmn.TrimAll(modCmn.GetRowText(existingRow, fieldName))
    End Select
    
    CompareFieldValues = (stagingValue <> existingValue)
End Function

'=============================================================================
' �f�[�^�R�s�[����
'=============================================================================

' Staging����Customers�ւ̃f�[�^�R�s�[
Private Sub CopyDataFromStaging(ByVal stagingRow As ListRow, ByVal customerRow As ListRow, _
                               ByVal isNewRecord As Boolean)
    On Error Resume Next
    
    Dim sourceFile As String
    sourceFile = modCmn.GetRowText(stagingRow, COL_SOURCE_FILE)
    
    ' ��{�f�[�^�R�s�[�i���K���ςݒl���g�p�j
    Call modCmn.SetRowText(customerRow, COL_CUSTOMER_ID, modCmn.GetRowText(stagingRow, COL_CUSTOMER_ID))
    Call modCmn.SetRowText(customerRow, COL_CUSTOMER_NAME, modCmn.GetRowText(stagingRow, COL_CUSTOMER_NAME))
    Call modCmn.SetRowText(customerRow, COL_EMAIL, modCmn.GetRowText(stagingRow, COL_EMAIL_NORM))
    Call modCmn.SetRowText(customerRow, COL_PHONE, modCmn.GetRowText(stagingRow, COL_PHONE_NORM))
    Call modCmn.SetRowText(customerRow, COL_ZIP, modCmn.GetRowText(stagingRow, COL_ZIP_NORM))
    Call modCmn.SetRowText(customerRow, COL_ADDRESS1, modCmn.GetRowText(stagingRow, COL_ADDRESS1))
    Call modCmn.SetRowText(customerRow, COL_ADDRESS2, modCmn.GetRowText(stagingRow, COL_ADDRESS2))
    Call modCmn.SetRowText(customerRow, COL_CATEGORY, modCmn.GetRowText(stagingRow, COL_CATEGORY))
    Call modCmn.SetRowText(customerRow, COL_STATUS, modCmn.GetRowText(stagingRow, COL_STATUS))
    Call modCmn.SetRowText(customerRow, COL_NOTES, modCmn.GetRowText(stagingRow, COL_NOTES))
    Call modCmn.SetRowText(customerRow, COL_SOURCE_FILE, sourceFile)
    
    ' �������ݒ�
    If isNewRecord Then
        Call modCmn.SetRowDate(customerRow, COL_CREATED_AT, Now)
    End If
    Call modCmn.SetRowDate(customerRow, COL_UPDATED_AT, Now)
End Sub

'=============================================================================
' �_���폜����
'=============================================================================

' �����؂�ڋq�̘_���폜
Public Function InactivateStaleCustomers() As Long
    On Error GoTo ErrHandler
    
    Dim customerTbl As ListObject
    Dim row As ListRow
    Dim inactivateDays As Long
    Dim cutoffDate As Date
    Dim lastUpdated As Date
    Dim inactivatedCount As Long
    Dim startTime As Double
    
    startTime = Timer
    InactivateStaleCustomers = 0
    
    ' �ݒ�l�擾
    inactivateDays = CLng(modData.GetConfigValue(CONFIG_INACTIVATE_DAYS))
    If inactivateDays <= 0 Then
        Call modCmn.LogWarn("InactivateStaleCustomers", "���������������ݒ�܂��͖����ł�")
        Exit Function
    End If
    
    cutoffDate = Now - inactivateDays
    
    Set customerTbl = modData.GetCustomersTable()
    If customerTbl Is Nothing Then Exit Function
    
    Call modCmn.ShowProgressStart(MSG_CLEANUP_STARTED)
    
    ' �e�ڋq���R�[�h���`�F�b�N
    For Each row In customerTbl.ListRows
        lastUpdated = modCmn.GetRowDate(row, COL_UPDATED_AT)
        
        ' �����؂ꂩ�L���ȃ��R�[�h�𖳌���
        If lastUpdated > 0 And lastUpdated < cutoffDate Then
            If modCmn.GetRowText(row, COL_STATUS) = STATUS_ACTIVE Then
                Call modCmn.SetRowText(row, COL_STATUS, STATUS_INACTIVE)
                Call modCmn.SetRowDate(row, COL_UPDATED_AT, Now)
                inactivatedCount = inactivatedCount + 1
            End If
        End If
    Next row
    
    ' ���O�L�^
    Call modData.LogImportOperation("�����؂�ڋq������", inactivatedCount, Timer - startTime, _
                                   "������臒l: " & inactivateDays & "��")
    
    InactivateStaleCustomers = inactivatedCount
    Call modCmn.HideProgress
    Exit Function
    
ErrHandler:
    Call modCmn.HideProgress
    InactivateStaleCustomers = -1
    Call modCmn.LogError("InactivateStaleCustomers", "�_���폜�G���[: " & Err.Description)
End Function

'=============================================================================
' �o�b�N�A�b�v����
'=============================================================================

' �ڋq�f�[�^�o�b�N�A�b�v
Public Function BackupCustomerData() As Boolean
    On Error GoTo ErrHandler
    
    Dim customerTbl As ListObject
    Dim backupDir As String
    Dim backupFileName As String
    Dim backupFilePath As String
    
    BackupCustomerData = False
    
    ' �o�b�N�A�b�v���L�����`�F�b�N
    If LCase(modData.GetConfigValue(CONFIG_BACKUP_ENABLED)) <> "true" Then
        Call modCmn.LogInfo("BackupCustomerData", "�o�b�N�A�b�v�@�\������������Ă��܂�")
        BackupCustomerData = True ' �����̏ꍇ�͐�������
        Exit Function
    End If
    
    backupDir = modData.GetConfigValue(CONFIG_BACKUP_DIR)
    If Not modCmn.DirectoryExists(backupDir) Then
        If Not modCmn.CreateDirectoryIfNotExists(backupDir) Then
            Call modCmn.LogWarn("BackupCustomerData", "�o�b�N�A�b�v�f�B���N�g���쐬���s: " & backupDir)
            Exit Function
        End If
    End If
    
    ' �o�b�N�A�b�v�t�@�C��������
    backupFileName = BACKUP_FILE_PREFIX & Format(Now, DATE_FORMAT_FILE) & ".xlsx"
    backupFilePath = backupDir & backupFileName
    
    ' ���݂̃��[�N�u�b�N���o�b�N�A�b�v
    Application.DisplayAlerts = False
    ThisWorkbook.SaveCopyAs backupFilePath
    Application.DisplayAlerts = True
    
    ' �Â��o�b�N�A�b�v�t�@�C���폜
    Call CleanupOldBackups(backupDir)
    
    BackupCustomerData = True
    Call modCmn.LogInfo("BackupCustomerData", "�o�b�N�A�b�v�쐬����: " & backupFileName)
    Exit Function
    
ErrHandler:
    Application.DisplayAlerts = True
    BackupCustomerData = False
    Call modCmn.LogError("BackupCustomerData", "�o�b�N�A�b�v�G���[: " & Err.Description)
End Function

' �Â��o�b�N�A�b�v�t�@�C���폜
Private Sub CleanupOldBackups(ByVal backupDir As String)
    On Error Resume Next
    
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim files As Collection
    Dim fileName As Variant
    Dim fileCount As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(backupDir)
    Set files = New Collection
    
    ' �o�b�N�A�b�v�t�@�C���ꗗ�擾
    For Each file In folder.Files
        If Left(file.Name, Len(BACKUP_FILE_PREFIX)) = BACKUP_FILE_PREFIX Then
            files.Add file.Name
            fileCount = fileCount + 1
        End If
    Next file
    
    ' �ő�ێ����𒴂��Ă���ꍇ�͌Â����̂��폜
    If fileCount > MAX_BACKUP_FILES Then
        ' �t�@�C�����Ń\�[�g�i���t���ɂȂ�j
        Dim sortedFiles As New Collection
        ' �ȈՓI�ȍ폜�i�����ȗ����j
        For Each fileName In files
            If fileCount > MAX_BACKUP_FILES Then
                Kill backupDir & fileName
                fileCount = fileCount - 1
                Call modCmn.LogInfo("CleanupOldBackups", "�Â��o�b�N�A�b�v�폜: " & fileName)
            End If
        Next fileName
    End If
End Sub

'=============================================================================
' ���[�e�B���e�B�֐�
'=============================================================================

' �A�b�v�T�[�g���v�擾
Public Function GetUpsertStatistics() As Object
    On Error Resume Next
    
    Dim stats As Object
    Dim stagingTbl As ListObject
    Dim row As ListRow
    Dim totalCount As Long
    Dim validCount As Long
    Dim errorCount As Long
    
    Set GetUpsertStatistics = CreateObject("Scripting.Dictionary")
    Set stats = GetUpsertStatistics
    
    Set stagingTbl = modData.GetStagingTable()
    If stagingTbl Is Nothing Then Exit Function
    
    For Each row In stagingTbl.ListRows
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

' �Ō�̃A�b�v�T�[�g�����擾
Public Function GetLastUpsertDateTime() As Date
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    Dim lastDateTime As Date
    Dim rowDateTime As Date
    
    GetLastUpsertDateTime = 0
    
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_LOGS)
    If ws Is Nothing Then Exit Function
    
    Set tbl = modCmn.GetTable(ws, TABLE_LOGS)
    If tbl Is Nothing Then Exit Function
    
    ' ���O����ŐV�̃A�b�v�T�[�g�L�^������
    For Each row In tbl.ListRows
        If InStr(modCmn.GetRowText(row, "Message"), "�A�b�v�T�[�g") > 0 Then
            rowDateTime = modCmn.SafeDate(modCmn.GetRowText(row, "Timestamp"))
            If rowDateTime > lastDateTime Then
                lastDateTime = rowDateTime
            End If
        End If
    Next row
    
    GetLastUpsertDateTime = lastDateTime
End Function