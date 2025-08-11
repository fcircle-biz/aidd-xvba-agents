Attribute VB_Name = "modValidation"
'=============================================================================
' modValidation.bas - �f�[�^���؁E�d�����o���W���[��
'=============================================================================
' �T�v:
'   Staging�f�[�^�̌��؁A�d�����o�A�`���`�F�b�N�A�K�{���ڃ`�F�b�N��
'   ���؃��[���̊Ǘ��A�G���[���̋L�^�@�\���
'=============================================================================
Option Explicit

'=============================================================================
' ���C�����؏���
'=============================================================================

' Staging�f�[�^���؎��s
Public Function ValidateStagingData() As Long
    On Error GoTo ErrHandler
    
    Dim tbl As ListObject
    Dim row As ListRow
    Dim errorCount As Long
    Dim totalCount As Long
    Dim startTime As Double
    
    startTime = Timer
    errorCount = 0
    totalCount = 0
    
    Set tbl = modData.GetStagingTable()
    If tbl Is Nothing Then
        ValidateStagingData = -1
        Exit Function
    End If
    
    Call modCmn.ShowProgressStart(MSG_VALIDATION_STARTED)
    
    ' �e�s�̌���
    For Each row In tbl.ListRows
        totalCount = totalCount + 1
        
        If ValidateStagingRow(row) > 0 Then
            errorCount = errorCount + 1
        End If
        
        ' �v���O���X�X�V
        If totalCount Mod BATCH_SIZE_VALIDATION = 0 Then
            Call modCmn.UpdateProgress("���؏�����: " & totalCount & " �� (�G���[: " & errorCount & " ��)")
        End If
    Next row
    
    ' �d�����o�iStaging���j
    errorCount = errorCount + DetectStagingDuplicates()
    
    ' �����ڋq�Ƃ̏d�����o
    errorCount = errorCount + DetectCustomerDuplicates()
    
    Call modData.LogImportOperation("�f�[�^���؊���", totalCount, Timer - startTime, _
                                   "�G���[����: " & errorCount)
    
    ValidateStagingData = errorCount
    Call modCmn.HideProgress
    Exit Function
    
ErrHandler:
    Call modCmn.HideProgress
    ValidateStagingData = -1
    Call modCmn.LogError("ValidateStagingData", "���؏����G���[: " & Err.Description)
End Function

' �P��s����
Private Function ValidateStagingRow(ByVal row As ListRow) As Long
    On Error Resume Next
    
    Dim errorMessages As Collection
    Dim errorMsg As Variant
    Dim fullErrorMessage As String
    
    Set errorMessages = New Collection
    
    ' �K�{�t�B�[���h�`�F�b�N
    Call CheckRequiredFields(row, errorMessages)
    
    ' �`���`�F�b�N
    Call CheckDataFormats(row, errorMessages)
    
    ' �r�W�l�X���[���`�F�b�N
    Call CheckBusinessRules(row, errorMessages)
    
    ' �G���[���b�Z�[�W����
    If errorMessages.Count > 0 Then
        For Each errorMsg In errorMessages
            If Len(fullErrorMessage) > 0 Then fullErrorMessage = fullErrorMessage & "; "
            fullErrorMessage = fullErrorMessage & CStr(errorMsg)
        Next errorMsg
        
        Call modCmn.SetRowValue(row, COL_IS_VALID, False)
        Call modCmn.SetRowText(row, COL_ERROR_MESSAGE, fullErrorMessage)
        ValidateStagingRow = 1
    Else
        Call modCmn.SetRowValue(row, COL_IS_VALID, True)
        Call modCmn.SetRowText(row, COL_ERROR_MESSAGE, "")
        ValidateStagingRow = 0
    End If
End Function

'=============================================================================
' �ʌ��؃��[��
'=============================================================================

' �K�{�t�B�[���h�`�F�b�N
Private Sub CheckRequiredFields(ByVal row As ListRow, ByVal errorMessages As Collection)
    On Error Resume Next
    
    Dim requiredFields As Collection
    Dim fieldName As Variant
    Dim fieldValue As String
    
    Set requiredFields = modCmn.SplitToCollection(modData.GetConfigValue(CONFIG_REQUIRED))
    
    For Each fieldName In requiredFields
        fieldValue = modCmn.TrimAll(modCmn.GetRowText(row, CStr(fieldName)))
        If Len(fieldValue) = 0 Then
            errorMessages.Add "�K�{�t�B�[���h����ł�: " & fieldName
        End If
    Next fieldName
End Sub

' �f�[�^�`���`�F�b�N
Private Sub CheckDataFormats(ByVal row As ListRow, ByVal errorMessages As Collection)
    On Error Resume Next
    
    Dim emailNorm As String
    Dim phoneNorm As String
    Dim zipNorm As String
    Dim customerId As String
    
    ' ���[���A�h���X�`���`�F�b�N
    emailNorm = modCmn.GetRowText(row, COL_EMAIL_NORM)
    If Len(emailNorm) > 0 And Not IsValidEmailFormat(emailNorm) Then
        errorMessages.Add ERR_INVALID_EMAIL_FORMAT & ": " & emailNorm
    End If
    
    ' �d�b�ԍ��`���`�F�b�N
    phoneNorm = modCmn.GetRowText(row, COL_PHONE_NORM)
    If Len(phoneNorm) > 0 And Not IsValidPhoneFormat(phoneNorm) Then
        errorMessages.Add ERR_INVALID_PHONE_FORMAT & ": " & phoneNorm
    End If
    
    ' �X�֔ԍ��`���`�F�b�N
    zipNorm = modCmn.GetRowText(row, COL_ZIP_NORM)
    If Len(zipNorm) > 0 And Not IsValidZipFormat(zipNorm) Then
        errorMessages.Add ERR_INVALID_ZIP_FORMAT & ": " & zipNorm
    End If
    
    ' �ڋqID�`���`�F�b�N
    customerId = modCmn.GetRowText(row, COL_CUSTOMER_ID)
    If Len(customerId) > 0 And Not IsValidCustomerIdFormat(customerId) Then
        errorMessages.Add "�����ȌڋqID�`��: " & customerId
    End If
End Sub

' �r�W�l�X���[���`�F�b�N
Private Sub CheckBusinessRules(ByVal row As ListRow, ByVal errorMessages As Collection)
    On Error Resume Next
    
    Dim status As String
    Dim category As String
    Dim customerName As String
    
    ' �X�e�[�^�X�l�`�F�b�N
    status = modCmn.GetRowText(row, COL_STATUS)
    If Not IsValidStatus(status) Then
        errorMessages.Add "�����ȃX�e�[�^�X�l: " & status
    End If
    
    ' �J�e�S���l�`�F�b�N
    category = modCmn.GetRowText(row, COL_CATEGORY)
    If Len(category) > 0 And Not IsValidCategory(category) Then
        errorMessages.Add "�����ȃJ�e�S���l: " & category
    End If
    
    ' �ڋq�������`�F�b�N
    customerName = modCmn.GetRowText(row, COL_CUSTOMER_NAME)
    If Len(customerName) > 100 Then
        errorMessages.Add "�ڋq�����������܂��i100�����ȓ��j: " & Left(customerName, 20) & "..."
    End If
End Sub

'=============================================================================
' �d�����o
'=============================================================================

' Staging���d�����o
Private Function DetectStagingDuplicates() As Long
    On Error GoTo ErrHandler
    
    Dim tbl As ListObject
    Dim keyDict As Object
    Dim row As ListRow
    Dim primaryKey As String
    Dim altKey As String
    Dim duplicateCount As Long
    
    DetectStagingDuplicates = 0
    
    Set tbl = modData.GetStagingTable()
    If tbl Is Nothing Then Exit Function
    
    Set keyDict = CreateObject("Scripting.Dictionary")
    
    ' 1��ځF�L�[���W
    For Each row In tbl.ListRows
        primaryKey = modCmn.GetRowText(row, COL_CUSTOMER_ID)
        altKey = modCmn.GetRowText(row, COL_KEY_CANDIDATE)
        
        ' ��L�[�`�F�b�N
        If Len(primaryKey) > 0 Then
            If keyDict.Exists("PK:" & primaryKey) Then
                keyDict("PK:" & primaryKey) = keyDict("PK:" & primaryKey) + 1
            Else
                keyDict("PK:" & primaryKey) = 1
            End If
        End If
        
        ' ��փL�[�`�F�b�N
        If Len(altKey) > 0 Then
            If keyDict.Exists("AK:" & altKey) Then
                keyDict("AK:" & altKey) = keyDict("AK:" & altKey) + 1
            Else
                keyDict("AK:" & altKey) = 1
            End If
        End If
    Next row
    
    ' 2��ځF�d���}�[�N
    For Each row In tbl.ListRows
        If modCmn.GetRowValue(row, COL_IS_VALID) = True Then ' ���̃G���[���Ȃ��ꍇ�̂�
            primaryKey = modCmn.GetRowText(row, COL_CUSTOMER_ID)
            altKey = modCmn.GetRowText(row, COL_KEY_CANDIDATE)
            
            Dim isDuplicate As Boolean
            Dim dupMessage As String
            
            ' ��L�[�d���`�F�b�N
            If Len(primaryKey) > 0 And keyDict.Exists("PK:" & primaryKey) Then
                If keyDict("PK:" & primaryKey) > 1 Then
                    isDuplicate = True
                    dupMessage = "Staging���ŌڋqID���d��: " & primaryKey
                End If
            End If
            
            ' ��փL�[�d���`�F�b�N�i��L�[���Ȃ��ꍇ�j
            If Not isDuplicate And Len(primaryKey) = 0 And Len(altKey) > 0 Then
                If keyDict.Exists("AK:" & altKey) And keyDict("AK:" & altKey) > 1 Then
                    isDuplicate = True
                    dupMessage = "Staging���ő�փL�[���d��: " & altKey
                End If
            End If
            
            ' �d���}�[�N
            If isDuplicate Then
                Call modCmn.SetRowValue(row, COL_IS_VALID, False)
                Dim existingMsg As String
                existingMsg = modCmn.GetRowText(row, COL_ERROR_MESSAGE)
                If Len(existingMsg) > 0 Then dupMessage = existingMsg & "; " & dupMessage
                Call modCmn.SetRowText(row, COL_ERROR_MESSAGE, dupMessage)
                duplicateCount = duplicateCount + 1
            End If
        End If
    Next row
    
    DetectStagingDuplicates = duplicateCount
    
    If duplicateCount > 0 Then
        Call modCmn.LogWarn("DetectStagingDuplicates", "Staging���d�����o: " & duplicateCount & " ��")
    End If
    
    Exit Function
    
ErrHandler:
    DetectStagingDuplicates = 0
    Call modCmn.LogError("DetectStagingDuplicates", "Staging�d�����o�G���[: " & Err.Description)
End Function

' �����ڋq�Ƃ̏d�����o
Private Function DetectCustomerDuplicates() As Long
    On Error GoTo ErrHandler
    
    Dim stagingTbl As ListObject
    Dim customerTbl As ListObject
    Dim stagingRow As ListRow
    Dim customerDict As Object
    Dim duplicateCount As Long
    
    DetectCustomerDuplicates = 0
    
    Set stagingTbl = modData.GetStagingTable()
    Set customerTbl = modData.GetCustomersTable()
    If stagingTbl Is Nothing Or customerTbl Is Nothing Then Exit Function
    
    ' �����ڋq�̃L�[�����쐬
    Set customerDict = CreateCustomerKeyDictionary(customerTbl)
    
    ' Staging�f�[�^�Ƃ̏d���`�F�b�N
    For Each stagingRow In stagingTbl.ListRows
        If modCmn.GetRowValue(stagingRow, COL_IS_VALID) = True Then ' ���̃G���[���Ȃ��ꍇ�̂�
            If IsCustomerDuplicate(stagingRow, customerDict) Then
                Call modCmn.SetRowValue(stagingRow, COL_IS_VALID, False)
                Dim existingMsg As String
                existingMsg = modCmn.GetRowText(stagingRow, COL_ERROR_MESSAGE)
                Dim dupMsg As String
                dupMsg = "�����ڋq�f�[�^�Ƃ̏d�������o����܂���"
                If Len(existingMsg) > 0 Then dupMsg = existingMsg & "; " & dupMsg
                Call modCmn.SetRowText(stagingRow, COL_ERROR_MESSAGE, dupMsg)
                duplicateCount = duplicateCount + 1
            End If
        End If
    Next stagingRow
    
    DetectCustomerDuplicates = duplicateCount
    
    If duplicateCount > 0 Then
        Call modCmn.LogWarn("DetectCustomerDuplicates", "�����ڋq�d�����o: " & duplicateCount & " ��")
    End If
    
    Exit Function
    
ErrHandler:
    DetectCustomerDuplicates = 0
    Call modCmn.LogError("DetectCustomerDuplicates", "�ڋq�d�����o�G���[: " & Err.Description)
End Function

' �ڋq�L�[�����쐬
Private Function CreateCustomerKeyDictionary(ByVal customerTbl As ListObject) As Object
    On Error Resume Next
    
    Dim customerDict As Object
    Dim row As ListRow
    Dim primaryKey As String
    Dim email As String
    Dim customerName As String
    Dim altKey As String
    
    Set CreateCustomerKeyDictionary = CreateObject("Scripting.Dictionary")
    Set customerDict = CreateCustomerKeyDictionary
    
    For Each row In customerTbl.ListRows
        ' ��L�[�i�ڋqID�j
        primaryKey = modCmn.GetRowText(row, COL_CUSTOMER_ID)
        If Len(primaryKey) > 0 Then
            customerDict("PK:" & primaryKey) = True
        End If
        
        ' ��փL�[�iEmail + CustomerName�j
        email = modCmn.NormalizeEmail(modCmn.GetRowText(row, COL_EMAIL))
        customerName = modCmn.GetRowText(row, COL_CUSTOMER_NAME)
        If Len(email) > 0 And Len(customerName) > 0 Then
            altKey = email & "+" & customerName
            customerDict("AK:" & altKey) = True
        End If
    Next row
End Function

' �ڋq�d������
Private Function IsCustomerDuplicate(ByVal stagingRow As ListRow, ByVal customerDict As Object) As Boolean
    On Error Resume Next
    
    Dim primaryKey As String
    Dim altKey As String
    
    IsCustomerDuplicate = False
    
    ' ��L�[�ɂ��d���`�F�b�N
    primaryKey = modCmn.GetRowText(stagingRow, COL_CUSTOMER_ID)
    If Len(primaryKey) > 0 Then
        If customerDict.Exists("PK:" & primaryKey) Then
            IsCustomerDuplicate = True
            Exit Function
        End If
    End If
    
    ' ��փL�[�ɂ��d���`�F�b�N�i��L�[���Ȃ��ꍇ�j
    If Len(primaryKey) = 0 Then
        altKey = modCmn.GetRowText(stagingRow, COL_KEY_CANDIDATE)
        If Len(altKey) > 0 Then
            If customerDict.Exists("AK:" & altKey) Then
                IsCustomerDuplicate = True
            End If
        End If
    End If
End Function

'=============================================================================
' �`�����؃w���p�[�֐�
'=============================================================================

' ���[���A�h���X�`������
Private Function IsValidEmailFormat(ByVal email As String) As Boolean
    On Error Resume Next
    
    If Len(email) < 5 Or InStr(email, "@") = 0 Then
        IsValidEmailFormat = False
        Exit Function
    End If
    
    ' ��{�I�Ȑ��K�\���`�F�b�N
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = modData.GetConfigValue(CONFIG_EMAIL_REGEX)
    regex.IgnoreCase = True
    
    IsValidEmailFormat = regex.Test(email)
End Function

' �d�b�ԍ��`������
Private Function IsValidPhoneFormat(ByVal phone As String) As Boolean
    On Error Resume Next
    
    IsValidPhoneFormat = modCmn.IsValidJapanesePhone(phone)
End Function

' �X�֔ԍ��`������
Private Function IsValidZipFormat(ByVal zip As String) As Boolean
    On Error Resume Next
    
    IsValidZipFormat = modCmn.IsValidJapaneseZip(zip)
End Function

' �ڋqID�`������
Private Function IsValidCustomerIdFormat(ByVal customerId As String) As Boolean
    On Error Resume Next
    
    IsValidCustomerIdFormat = modCmn.IsValidCustomerId(customerId)
End Function

' �X�e�[�^�X�l����
Private Function IsValidStatus(ByVal status As String) As Boolean
    On Error Resume Next
    
    Select Case status
        Case STATUS_ACTIVE, STATUS_INACTIVE, STATUS_SUSPENDED
            IsValidStatus = True
        Case Else
            IsValidStatus = False
    End Select
End Function

' �J�e�S���l����
Private Function IsValidCategory(ByVal category As String) As Boolean
    On Error Resume Next
    
    Select Case category
        Case CATEGORY_B2B, CATEGORY_B2C, CATEGORY_PARTNER, CATEGORY_RESELLER
            IsValidCategory = True
        Case Else
            IsValidCategory = False
    End Select
End Function

'=============================================================================
' ���؃��|�[�g����
'=============================================================================

' ���؃G���[���|�[�g����
Public Function GenerateValidationReport() As String
    On Error GoTo ErrHandler
    
    Dim tbl As ListObject
    Dim row As ListRow
    Dim report As String
    Dim totalCount As Long
    Dim errorCount As Long
    Dim validCount As Long
    
    Set tbl = modData.GetStagingTable()
    If tbl Is Nothing Then
        GenerateValidationReport = "Staging�e�[�u�����擾�ł��܂���ł����B"
        Exit Function
    End If
    
    ' ���v���W
    For Each row In tbl.ListRows
        totalCount = totalCount + 1
        If modCmn.GetRowValue(row, COL_IS_VALID) = True Then
            validCount = validCount + 1
        Else
            errorCount = errorCount + 1
        End If
    Next row
    
    ' ���|�[�g����
    report = "=== �f�[�^���؃��|�[�g ===" & vbCrLf
    report = report & "���ؓ���: " & modCmn.GetCurrentDateTimeString() & vbCrLf
    report = report & "������: " & Format(totalCount, NUMBER_FORMAT_COUNT) & " ��" & vbCrLf
    report = report & "�L������: " & Format(validCount, NUMBER_FORMAT_COUNT) & " ��" & vbCrLf
    report = report & "�G���[����: " & Format(errorCount, NUMBER_FORMAT_COUNT) & " ��" & vbCrLf
    
    If totalCount > 0 Then
        report = report & "�L����: " & Format(validCount / totalCount * 100, "0.0") & "%" & vbCrLf
    End If
    
    GenerateValidationReport = report
    Exit Function
    
ErrHandler:
    GenerateValidationReport = "���؃��|�[�g�����G���[: " & Err.Description
    Call modCmn.LogError("GenerateValidationReport", Err.Description)
End Function

' �L����Staging���R�[�h���擾
Public Function GetValidStagingCount() As Long
    On Error Resume Next
    
    Dim tbl As ListObject
    Dim row As ListRow
    
    GetValidStagingCount = 0
    
    Set tbl = modData.GetStagingTable()
    If tbl Is Nothing Then Exit Function
    
    For Each row In tbl.ListRows
        If modCmn.GetRowValue(row, COL_IS_VALID) = True Then
            GetValidStagingCount = GetValidStagingCount + 1
        End If
    Next row
End Function

' �G���[���R�[�h���擾
Public Function GetErrorStagingCount() As Long
    On Error Resume Next
    
    Dim tbl As ListObject
    Dim row As ListRow
    
    GetErrorStagingCount = 0
    
    Set tbl = modData.GetStagingTable()
    If tbl Is Nothing Then Exit Function
    
    For Each row In tbl.ListRows
        If modCmn.GetRowValue(row, COL_IS_VALID) <> True Then
            GetErrorStagingCount = GetErrorStagingCount + 1
        End If
    Next row
End Function