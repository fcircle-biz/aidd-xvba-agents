Attribute VB_Name = "common"

'===================================================================
' ���i�݌ɊǗ��V�X�e�� - ���ʃ��W���[��
'===================================================================

' �O���[�o���ϐ�
Public Const PRODUCT_SHEET As String = "���i�}�X�^"
Public Const INVENTORY_SHEET As String = "�݌ɊǗ�"
Public Const TRANSACTION_SHEET As String = "�������"
Public Const REPORT_SHEET As String = "���|�[�g"

' ���i�Ǘ��֘A�̌^��`
Public Type ProductInfo
    ProductID As String
    ProductName As String
    Category As String
    Price As Double
    MinStock As Long
    MaxStock As Long
    CurrentStock As Long
    Supplier As String
    LastUpdated As Date
    IsActive As Boolean
End Type

Public Type InventoryTransaction
    TransactionID As String
    ProductID As String
    TransactionType As String  ' "����", "�o��", "�I��"
    Quantity As Long
    TransactionDate As Date
    UserName As String
    Notes As String
    ReferenceNo As String
End Type

'===================================================================
' ���ʃ��[�e�B���e�B�֐�
'===================================================================

' Xdebug�Ƃ̓���
Public Sub XLog(ByVal message As String, Optional ByVal level As String = "INFO")
    On Error Resume Next
    
    ' Xdebug�����p�\�ȏꍇ�͎g�p
    If IsXdebugAvailable() Then
        Dim xd As Object
        Set xd = CreateObject("Xdebug.Logger")
        xd.Log level & ": " & message
    End If
    
    ' ���O�t�@�C���ɂ��o��
    LogToFile message, level
End Sub

' Xdebug�̗��p�\���`�F�b�N
Public Function IsXdebugAvailable() As Boolean
    On Error GoTo ErrorHandler
    Dim xd As Object
    Set xd = CreateObject("Xdebug.Logger")
    IsXdebugAvailable = True
    Exit Function
    
ErrorHandler:
    IsXdebugAvailable = False
End Function

' �t�@�C�����O�o��
Public Sub LogToFile(ByVal message As String, Optional ByVal level As String = "INFO")
    On Error Resume Next
    
    Dim logPath As String
    logPath = ThisWorkbook.Path & "\logs\inventory_system.log"
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open logPath For Append As #fileNum
    Print #fileNum, Format(Now, "yyyy-mm-dd hh:nn:ss") & " [" & level & "] " & message
    Close #fileNum
End Sub

'===================================================================
' ���i�}�X�^�Ǘ��@�\
'===================================================================

' ���i����ǉ�/�X�V
Public Function AddOrUpdateProduct(product As ProductInfo) As Boolean
    On Error GoTo ErrorHandler
    
    XLog "���i���̒ǉ�/�X�V���J�n: " & product.ProductID
    
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(PRODUCT_SHEET)
    
    ' �w�b�_�[�s�̊m�F�E�쐬
    If ws.Cells(1, 1).Value = "" Then
        Call CreateProductMasterHeaders(ws)
    End If
    
    ' �������i�̃`�F�b�N
    Dim targetRow As Long
    targetRow = FindProductRow(product.ProductID)
    
    If targetRow = 0 Then
        ' �V�K�ǉ�
        targetRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        XLog "�V�K���i�Ƃ��Ēǉ�: �s" & targetRow
    Else
        XLog "�������i���X�V: �s" & targetRow
    End If
    
    ' �f�[�^�ݒ�
    With ws
        .Cells(targetRow, 1).Value = product.ProductID
        .Cells(targetRow, 2).Value = product.ProductName
        .Cells(targetRow, 3).Value = product.Category
        .Cells(targetRow, 4).Value = product.Price
        .Cells(targetRow, 5).Value = product.MinStock
        .Cells(targetRow, 6).Value = product.MaxStock
        .Cells(targetRow, 7).Value = product.CurrentStock
        .Cells(targetRow, 8).Value = product.Supplier
        .Cells(targetRow, 9).Value = product.LastUpdated
        .Cells(targetRow, 10).Value = product.IsActive
    End With
    
    AddOrUpdateProduct = True
    XLog "���i���̒ǉ�/�X�V���������܂���"
    Exit Function
    
ErrorHandler:
    XLog "���i���̒ǉ�/�X�V�ŃG���[������: " & Err.Description, "ERROR"
    AddOrUpdateProduct = False
End Function

' ���i�}�X�^�̃w�b�_�[�쐬
Public Sub CreateProductMasterHeaders(ws As Worksheet)
    With ws
        .Cells(1, 1).Value = "���iID"
        .Cells(1, 2).Value = "���i��"
        .Cells(1, 3).Value = "�J�e�S��"
        .Cells(1, 4).Value = "���i"
        .Cells(1, 5).Value = "�ŏ��݌�"
        .Cells(1, 6).Value = "�ő�݌�"
        .Cells(1, 7).Value = "���ݍ݌�"
        .Cells(1, 8).Value = "�d����"
        .Cells(1, 9).Value = "�X�V����"
        .Cells(1, 10).Value = "�L���t���O"
        
        ' �w�b�_�[�s�̏����ݒ�
        With .Range("A1:J1")
            .Font.Bold = True
            .Interior.Color = RGB(200, 200, 200)
            .Borders.LineStyle = xlContinuous
        End With
    End With
    
    XLog "���i�}�X�^�̃w�b�_�[���쐬���܂���"
End Sub

' ���iID����s�ԍ�������
Public Function FindProductRow(productID As String) As Long
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = Worksheets(PRODUCT_SHEET)
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = productID Then
            FindProductRow = i
            Exit Function
        End If
    Next i
    
    FindProductRow = 0
    Exit Function
    
ErrorHandler:
    FindProductRow = 0
End Function

' ���i�����擾
Public Function GetProduct(productID As String) As ProductInfo
    On Error GoTo ErrorHandler
    
    Dim product As ProductInfo
    Dim targetRow As Long
    targetRow = FindProductRow(productID)
    
    If targetRow > 0 Then
        Dim ws As Worksheet
        Set ws = Worksheets(PRODUCT_SHEET)
        
        With product
            .ProductID = ws.Cells(targetRow, 1).Value
            .ProductName = ws.Cells(targetRow, 2).Value
            .Category = ws.Cells(targetRow, 3).Value
            .Price = ws.Cells(targetRow, 4).Value
            .MinStock = ws.Cells(targetRow, 5).Value
            .MaxStock = ws.Cells(targetRow, 6).Value
            .CurrentStock = ws.Cells(targetRow, 7).Value
            .Supplier = ws.Cells(targetRow, 8).Value
            .LastUpdated = ws.Cells(targetRow, 9).Value
            .IsActive = ws.Cells(targetRow, 10).Value
        End With
        
        XLog "���i�����擾���܂���: " & productID
    Else
        XLog "���i��������܂���: " & productID, "WARNING"
    End If
    
    GetProduct = product
    Exit Function
    
ErrorHandler:
    XLog "���i���擾�ŃG���[������: " & Err.Description, "ERROR"
End Function

'===================================================================
' ���[�N�V�[�g�Ǘ�
'===================================================================

' �V�[�g���擾�܂��͍쐬
Public Function GetOrCreateSheet(sheetName As String) As Worksheet
    On Error Resume Next
    
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    
    If GetOrCreateSheet Is Nothing Then
        ' ������Sheet1-9���g�p���ăV�[�g����ύX
        Dim availableSheet As Worksheet
        Dim i As Integer
        For i = 1 To 9
            Set availableSheet = ThisWorkbook.Worksheets("Sheet" & i)
            If Not availableSheet Is Nothing And availableSheet.Cells(1, 1).Value = "" Then
                availableSheet.Name = sheetName
                Set GetOrCreateSheet = availableSheet
                XLog "Sheet" & i & "��" & sheetName & "�ɕύX���܂���"
                Exit For
            End If
        Next i
        
        ' ���p�\�ȃV�[�g��������Ȃ��ꍇ�͍ŏ��̃V�[�g���g�p
        If GetOrCreateSheet Is Nothing Then
            Set GetOrCreateSheet = ThisWorkbook.Worksheets("Sheet1")
            GetOrCreateSheet.Name = sheetName
            XLog "Sheet1��" & sheetName & "�ɕύX���܂���"
        End If
    End If
    
    On Error GoTo 0
End Function

' �f�[�^���؋@�\
Public Function ValidateProductData(product As ProductInfo) As String
    Dim errorMessage As String
    
    ' �K�{�t�B�[���h�̃`�F�b�N
    If Trim(product.ProductID) = "" Then
        errorMessage = errorMessage & "���iID�͕K�{�ł��B" & vbCrLf
    End If
    
    If Trim(product.ProductName) = "" Then
        errorMessage = errorMessage & "���i���͕K�{�ł��B" & vbCrLf
    End If
    
    ' ���l�f�[�^�̑Ó����`�F�b�N
    If product.Price < 0 Then
        errorMessage = errorMessage & "���i��0�ȏ�ł���K�v������܂��B" & vbCrLf
    End If
    
    If product.MinStock < 0 Then
        errorMessage = errorMessage & "�ŏ��݌ɂ�0�ȏ�ł���K�v������܂��B" & vbCrLf
    End If
    
    If product.MaxStock < product.MinStock Then
        errorMessage = errorMessage & "�ő�݌ɂ͍ŏ��݌Ɉȏ�ł���K�v������܂��B" & vbCrLf
    End If
    
    If product.CurrentStock < 0 Then
        errorMessage = errorMessage & "���ݍ݌ɂ�0�ȏ�ł���K�v������܂��B" & vbCrLf
    End If
    
    ValidateProductData = errorMessage
End Function

'===================================================================
' �݌ɊǗ��@�\
'===================================================================

' �݌Ɏ�����L�^
Public Function RecordInventoryTransaction(transaction As InventoryTransaction) As Boolean
    On Error GoTo ErrorHandler
    
    XLog "�݌Ɏ���̋L�^���J�n: " & transaction.TransactionID
    
    ' �f�[�^����
    Dim validationError As String
    validationError = ValidateTransactionData(transaction)
    If validationError <> "" Then
        MsgBox "�f�[�^�G���[:" & vbCrLf & validationError, vbCritical, "�݌ɊǗ��G���["
        RecordInventoryTransaction = False
        Exit Function
    End If
    
    ' �݌ɍX�V�̎��s
    If Not UpdateProductStock(transaction) Then
        RecordInventoryTransaction = False
        Exit Function
    End If
    
    ' ��������̋L�^
    If Not LogTransactionHistory(transaction) Then
        RecordInventoryTransaction = False
        Exit Function
    End If
    
    RecordInventoryTransaction = True
    XLog "�݌Ɏ���̋L�^���������܂���: " & transaction.TransactionID
    Exit Function
    
ErrorHandler:
    XLog "�݌Ɏ���L�^�ŃG���[������: " & Err.Description, "ERROR"
    RecordInventoryTransaction = False
End Function

' ���i�݌ɂ��X�V
Public Function UpdateProductStock(transaction As InventoryTransaction) As Boolean
    On Error GoTo ErrorHandler
    
    Dim productRow As Long
    productRow = FindProductRow(transaction.ProductID)
    
    If productRow = 0 Then
        MsgBox "���iID " & transaction.ProductID & " ��������܂���B", vbCritical, "�݌ɊǗ��G���["
        UpdateProductStock = False
        Exit Function
    End If
    
    Dim ws As Worksheet
    Set ws = Worksheets(PRODUCT_SHEET)
    
    Dim currentStock As Long
    currentStock = ws.Cells(productRow, 7).Value
    
    Dim newStock As Long
    Select Case transaction.TransactionType
        Case "����"
            newStock = currentStock + transaction.Quantity
        Case "�o��"
            newStock = currentStock - transaction.Quantity
            If newStock < 0 Then
                MsgBox "�݌ɕs���ł��B���ݍ݌�: " & currentStock & ", �v������: " & transaction.Quantity, vbCritical, "�݌ɊǗ��G���["
                UpdateProductStock = False
                Exit Function
            End If
        Case "�I��"
            newStock = transaction.Quantity ' �I���̏ꍇ�͐�Βl
        Case Else
            MsgBox "�����Ȏ����ʂł�: " & transaction.TransactionType, vbCritical, "�݌ɊǗ��G���["
            UpdateProductStock = False
            Exit Function
    End Select
    
    ' �݌ɐ��̍X�V
    ws.Cells(productRow, 7).Value = newStock
    ws.Cells(productRow, 9).Value = Now ' �X�V����
    
    ' �݌ɃA���[�g�̃`�F�b�N
    CheckStockAlerts transaction.ProductID, newStock
    
    UpdateProductStock = True
    XLog "�݌ɂ��X�V���܂��� - ���iID: " & transaction.ProductID & ", �V�݌�: " & newStock
    Exit Function
    
ErrorHandler:
    XLog "�݌ɍX�V�ŃG���[������: " & Err.Description, "ERROR"
    UpdateProductStock = False
End Function

' ��������̋L�^
Public Function LogTransactionHistory(transaction As InventoryTransaction) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(TRANSACTION_SHEET)
    
    ' �w�b�_�[�s�̊m�F�E�쐬
    If ws.Cells(1, 1).Value = "" Then
        Call CreateTransactionHistoryHeaders(ws)
    End If
    
    ' �V�����s�Ƀf�[�^��ǉ�
    Dim newRow As Long
    newRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    With ws
        .Cells(newRow, 1).Value = transaction.TransactionID
        .Cells(newRow, 2).Value = transaction.ProductID
        .Cells(newRow, 3).Value = transaction.TransactionType
        .Cells(newRow, 4).Value = transaction.Quantity
        .Cells(newRow, 5).Value = transaction.TransactionDate
        .Cells(newRow, 6).Value = transaction.UserName
        .Cells(newRow, 7).Value = transaction.Notes
        .Cells(newRow, 8).Value = transaction.ReferenceNo
    End With
    
    LogTransactionHistory = True
    XLog "����������L�^���܂���: " & transaction.TransactionID
    Exit Function
    
ErrorHandler:
    XLog "��������L�^�ŃG���[������: " & Err.Description, "ERROR"
    LogTransactionHistory = False
End Function

' ��������̃w�b�_�[�쐬
Public Sub CreateTransactionHistoryHeaders(ws As Worksheet)
    With ws
        .Cells(1, 1).Value = "���ID"
        .Cells(1, 2).Value = "���iID"
        .Cells(1, 3).Value = "������"
        .Cells(1, 4).Value = "����"
        .Cells(1, 5).Value = "�������"
        .Cells(1, 6).Value = "�S����"
        .Cells(1, 7).Value = "���l"
        .Cells(1, 8).Value = "�Q�Ɣԍ�"
        
        ' �w�b�_�[�s�̏����ݒ�
        With .Range("A1:H1")
            .Font.Bold = True
            .Interior.Color = RGB(180, 220, 180)
            .Borders.LineStyle = xlContinuous
        End With
    End With
    
    XLog "��������̃w�b�_�[���쐬���܂���"
End Sub

' �݌ɃA���[�g�̃`�F�b�N
Public Sub CheckStockAlerts(productID As String, currentStock As Long)
    On Error Resume Next
    
    Dim product As ProductInfo
    product = GetProduct(productID)
    
    If currentStock <= product.MinStock Then
        XLog "�݌ɃA���[�g: " & product.ProductName & " (ID: " & productID & ") �̍݌ɂ��ŏ��l�������܂����B���ݍ݌�: " & currentStock, "WARNING"
        
        ' �݌ɃA���[�g��ʂ̕\���i�I�v�V�����j
        ShowStockAlert product, currentStock
    ElseIf currentStock >= product.MaxStock Then
        XLog "�݌ɉߑ��A���[�g: " & product.ProductName & " (ID: " & productID & ") �̍݌ɂ��ő�l�𒴂��܂����B���ݍ݌�: " & currentStock, "WARNING"
    End If
End Sub

' �݌ɃA���[�g��ʂ̕\��
Public Sub ShowStockAlert(product As ProductInfo, currentStock As Long)
    Dim message As String
    message = "�݌ɃA���[�g" & vbCrLf & vbCrLf
    message = message & "���i��: " & product.ProductName & vbCrLf
    message = message & "���iID: " & product.ProductID & vbCrLf
    message = message & "���ݍ݌�: " & currentStock & vbCrLf
    message = message & "�ŏ��݌�: " & product.MinStock & vbCrLf
    message = message & "������������: " & (product.MaxStock - currentStock)
    
    MsgBox message, vbExclamation, "�݌ɊǗ��V�X�e��"
End Sub

' ����f�[�^�̌���
Public Function ValidateTransactionData(transaction As InventoryTransaction) As String
    Dim errorMessage As String
    
    ' �K�{�t�B�[���h�̃`�F�b�N
    If Trim(transaction.TransactionID) = "" Then
        errorMessage = errorMessage & "���ID�͕K�{�ł��B" & vbCrLf
    End If
    
    If Trim(transaction.ProductID) = "" Then
        errorMessage = errorMessage & "���iID�͕K�{�ł��B" & vbCrLf
    End If
    
    If Trim(transaction.TransactionType) = "" Then
        errorMessage = errorMessage & "�����ʂ͕K�{�ł��B" & vbCrLf
    ElseIf transaction.TransactionType <> "����" And transaction.TransactionType <> "�o��" And transaction.TransactionType <> "�I��" Then
        errorMessage = errorMessage & "�����ʂ́u���Ɂv�u�o�Ɂv�u�I���v�̂����ꂩ���w�肵�Ă��������B" & vbCrLf
    End If
    
    If transaction.Quantity <= 0 Then
        errorMessage = errorMessage & "���ʂ�1�ȏ���w�肵�Ă��������B" & vbCrLf
    End If
    
    If Trim(transaction.UserName) = "" Then
        errorMessage = errorMessage & "�S���Җ��͕K�{�ł��B" & vbCrLf
    End If
    
    ' ���i�̑��݃`�F�b�N
    If FindProductRow(transaction.ProductID) = 0 Then
        errorMessage = errorMessage & "�w�肳�ꂽ���iID�u" & transaction.ProductID & "�v�����݂��܂���B" & vbCrLf
    End If
    
    ValidateTransactionData = errorMessage
End Function

'===================================================================
' �����E�t�B���^�����O�@�\
'===================================================================

' ���i�����@�\
Public Function SearchProducts(searchTerm As String, Optional categoryFilter As String = "", Optional stockFilter As String = "") As Collection
    On Error GoTo ErrorHandler
    
    Set SearchProducts = New Collection
    
    Dim ws As Worksheet
    Set ws = Worksheets(PRODUCT_SHEET)
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 2 Then ' �w�b�_�[�s�݂̂̏ꍇ
        Exit Function
    End If
    
    Dim i As Long
    For i = 2 To lastRow
        Dim product As ProductInfo
        With product
            .ProductID = ws.Cells(i, 1).Value
            .ProductName = ws.Cells(i, 2).Value
            .Category = ws.Cells(i, 3).Value
            .Price = ws.Cells(i, 4).Value
            .MinStock = ws.Cells(i, 5).Value
            .MaxStock = ws.Cells(i, 6).Value
            .CurrentStock = ws.Cells(i, 7).Value
            .Supplier = ws.Cells(i, 8).Value
            .LastUpdated = ws.Cells(i, 9).Value
            .IsActive = ws.Cells(i, 10).Value
        End With
        
        ' ���������̃`�F�b�N
        Dim matchesSearch As Boolean
        matchesSearch = (searchTerm = "" Or _
                        InStr(1, product.ProductName, searchTerm, vbTextCompare) > 0 Or _
                        InStr(1, product.ProductID, searchTerm, vbTextCompare) > 0)
        
        Dim matchesCategory As Boolean
        matchesCategory = (categoryFilter = "" Or product.Category = categoryFilter)
        
        Dim matchesStock As Boolean
        Select Case stockFilter
            Case "��݌�"
                matchesStock = (product.CurrentStock <= product.MinStock)
            Case "�ߏ�݌�"
                matchesStock = (product.CurrentStock >= product.MaxStock)
            Case "�K���݌�"
                matchesStock = (product.CurrentStock > product.MinStock And product.CurrentStock < product.MaxStock)
            Case Else
                matchesStock = True
        End Select
        
        If matchesSearch And matchesCategory And matchesStock And product.IsActive Then
            SearchProducts.Add product, product.ProductID
        End If
    Next i
    
    XLog "���i���������s���܂����B������: " & searchTerm & ", ���ʌ���: " & SearchProducts.Count
    Exit Function
    
ErrorHandler:
    XLog "���i�����ŃG���[������: " & Err.Description, "ERROR"
    Set SearchProducts = New Collection
End Function

'===================================================================
' ���|�[�g�E�����@�\
'===================================================================

' �݌ɏ󋵃��|�[�g�̐���
Public Sub GenerateInventoryStatusReport()
    On Error GoTo ErrorHandler
    
    XLog "�݌ɏ󋵃��|�[�g�������J�n"
    
    Dim reportWs As Worksheet
    Set reportWs = GetOrCreateSheet(REPORT_SHEET)
    
    ' ���|�[�g�V�[�g���N���A
    reportWs.Cells.Clear
    
    ' ���|�[�g�w�b�_�[�̍쐬
    CreateInventoryReportHeaders reportWs
    
    ' �f�[�^�W�v�Əo��
    Dim ws As Worksheet
    Set ws = Worksheets(PRODUCT_SHEET)
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 2 Then
        reportWs.Cells(3, 1).Value = "�f�[�^������܂���"
        Exit Sub
    End If
    
    Dim reportRow As Long
    reportRow = 3
    
    Dim totalValue As Double
    Dim lowStockCount As Long
    Dim overStockCount As Long
    
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 10).Value Then ' �A�N�e�B�u�ȏ��i�̂�
            Dim stockStatus As String
            Dim currentStock As Long
            Dim minStock As Long
            Dim maxStock As Long
            Dim price As Double
            
            currentStock = ws.Cells(i, 7).Value
            minStock = ws.Cells(i, 5).Value
            maxStock = ws.Cells(i, 6).Value
            price = ws.Cells(i, 4).Value
            
            If currentStock <= minStock Then
                stockStatus = "��݌�"
                lowStockCount = lowStockCount + 1
            ElseIf currentStock >= maxStock Then
                stockStatus = "�ߏ�݌�"
                overStockCount = overStockCount + 1
            Else
                stockStatus = "�K��"
            End If
            
            totalValue = totalValue + (currentStock * price)
            
            ' ���|�[�g�s�Ƀf�[�^���o��
            With reportWs
                .Cells(reportRow, 1).Value = ws.Cells(i, 1).Value ' ���iID
                .Cells(reportRow, 2).Value = ws.Cells(i, 2).Value ' ���i��
                .Cells(reportRow, 3).Value = ws.Cells(i, 3).Value ' �J�e�S��
                .Cells(reportRow, 4).Value = currentStock
                .Cells(reportRow, 5).Value = minStock
                .Cells(reportRow, 6).Value = maxStock
                .Cells(reportRow, 7).Value = stockStatus
                .Cells(reportRow, 8).Value = currentStock * price ' �݌ɋ��z
                
                ' �݌ɏ󋵂ɉ������F����
                Select Case stockStatus
                    Case "��݌�"
                        .Cells(reportRow, 7).Interior.Color = RGB(255, 200, 200)
                    Case "�ߏ�݌�"
                        .Cells(reportRow, 7).Interior.Color = RGB(255, 255, 200)
                    Case "�K��"
                        .Cells(reportRow, 7).Interior.Color = RGB(200, 255, 200)
                End Select
            End With
            
            reportRow = reportRow + 1
        End If
    Next i
    
    ' �T�}���[���̒ǉ�
    CreateReportSummary reportWs, reportRow + 2, totalValue, lowStockCount, overStockCount
    
    ' ���|�[�g�̏����ݒ�
    FormatInventoryReport reportWs, reportRow - 1
    
    XLog "�݌ɏ󋵃��|�[�g�������������܂���"
    MsgBox "�݌ɏ󋵃��|�[�g���u" & REPORT_SHEET & "�v�V�[�g�ɐ�������܂����B", vbInformation, "�݌ɊǗ��V�X�e��"
    Exit Sub
    
ErrorHandler:
    XLog "���|�[�g�����ŃG���[������: " & Err.Description, "ERROR"
    MsgBox "���|�[�g�������ɃG���[���������܂���: " & Err.Description, vbCritical, "�݌ɊǗ��G���["
End Sub

' ���|�[�g�w�b�_�[�̍쐬
Public Sub CreateInventoryReportHeaders(ws As Worksheet)
    ws.Cells(1, 1).Value = "�݌ɏ󋵃��|�[�g - " & Format(Now, "yyyy/mm/dd hh:nn")
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 14
    
    With ws
        .Cells(2, 1).Value = "���iID"
        .Cells(2, 2).Value = "���i��"
        .Cells(2, 3).Value = "�J�e�S��"
        .Cells(2, 4).Value = "���ݍ݌�"
        .Cells(2, 5).Value = "�ŏ��݌�"
        .Cells(2, 6).Value = "�ő�݌�"
        .Cells(2, 7).Value = "�݌ɏ�"
        .Cells(2, 8).Value = "�݌ɋ��z"
        
        With .Range("A2:H2")
            .Font.Bold = True
            .Interior.Color = RGB(200, 200, 200)
            .Borders.LineStyle = xlContinuous
        End With
    End With
End Sub

' ���|�[�g�T�}���[�̍쐬
Public Sub CreateReportSummary(ws As Worksheet, startRow As Long, totalValue As Double, lowStockCount As Long, overStockCount As Long)
    With ws
        .Cells(startRow, 1).Value = "�T�}���["
        .Cells(startRow, 1).Font.Bold = True
        
        .Cells(startRow + 1, 1).Value = "���݌ɋ��z:"
        .Cells(startRow + 1, 2).Value = Format(totalValue, "#,##0")
        
        .Cells(startRow + 2, 1).Value = "��݌ɏ��i��:"
        .Cells(startRow + 2, 2).Value = lowStockCount
        
        .Cells(startRow + 3, 1).Value = "�ߏ�݌ɏ��i��:"
        .Cells(startRow + 3, 2).Value = overStockCount
        
        .Range(.Cells(startRow, 1), .Cells(startRow + 3, 2)).Font.Bold = True
    End With
End Sub

' ���|�[�g�̏����ݒ�
Public Sub FormatInventoryReport(ws As Worksheet, lastDataRow As Long)
    With ws
        ' �񕝂̎�������
        .Columns("A:H").AutoFit
        
        ' ���l��̏����ݒ�
        .Range("D3:F" & lastDataRow).NumberFormat = "#,##0"
        .Range("H3:H" & lastDataRow).NumberFormat = "#,##0"
        
        ' �S�f�[�^�͈͂ɘg����ݒ�
        With .Range("A2:H" & lastDataRow)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
    End With
End Sub

'===================================================================
' �T���v���f�[�^�����֐�
'===================================================================

' �T���v�����i�f�[�^�̍쐬
Public Sub CreateSampleData()
    On Error GoTo ErrorHandler
    
    XLog "�T���v���f�[�^�𐶐���"
    
    Dim products(1 To 10) As ProductInfo
    
    ' �T���v���f�[�^�̒�`
    With products(1)
        .ProductID = "P001"
        .ProductName = "�m�[�gPC ThinkPad X1"
        .Category = "PC�E���Ӌ@��"
        .Price = 180000
        .MinStock = 5
        .MaxStock = 20
        .CurrentStock = 12
        .Supplier = "���m�{�E�W���p��"
        .LastUpdated = Now
        .IsActive = True
    End With
    
    With products(2)
        .ProductID = "P002"
        .ProductName = "���C�����X�}�E�X MX Master 3"
        .Category = "PC�E���Ӌ@��"
        .Price = 12000
        .MinStock = 10
        .MaxStock = 50
        .CurrentStock = 3
        .Supplier = "���W�N�[��"
        .LastUpdated = Now
        .IsActive = True
    End With
    
    With products(3)
        .ProductID = "P003"
        .ProductName = "�I�t�B�X�`�F�A �G���S�q���[�}��"
        .Category = "�I�t�B�X�Ƌ�"
        .Price = 120000
        .MinStock = 2
        .MaxStock = 10
        .CurrentStock = 8
        .Supplier = "�I�J����"
        .LastUpdated = Now
        .IsActive = True
    End With
    
    With products(4)
        .ProductID = "P004"
        .ProductName = "�v�����^�����@ MFC-L3770CDW"
        .Category = "PC�E���Ӌ@��"
        .Price = 45000
        .MinStock = 3
        .MaxStock = 15
        .CurrentStock = 18
        .Supplier = "�u���U�[�H��"
        .LastUpdated = Now
        .IsActive = True
    End With
    
    With products(5)
        .ProductID = "P005"
        .ProductName = "���j�^�[ 27�C���` 4K"
        .Category = "PC�E���Ӌ@��"
        .Price = 85000
        .MinStock = 8
        .MaxStock = 25
        .CurrentStock = 15
        .Supplier = "�f��"
        .LastUpdated = Now
        .IsActive = True
    End With
    
    With products(6)
        .ProductID = "P006"
        .ProductName = "�z���C�g�{�[�h 1800�~900"
        .Category = "�I�t�B�X�p�i"
        .Price = 25000
        .MinStock = 5
        .MaxStock = 15
        .CurrentStock = 2
        .Supplier = "�R�N��"
        .LastUpdated = Now
        .IsActive = True
    End With
    
    With products(7)
        .ProductID = "P007"
        .ProductName = "�V�����b�_�[ GCS280i"
        .Category = "�I�t�B�X�@��"
        .Price = 180000
        .MinStock = 1
        .MaxStock = 5
        .CurrentStock = 3
        .Supplier = "�t�F���[�Y"
        .LastUpdated = Now
        .IsActive = True
    End With
    
    With products(8)
        .ProductID = "P008"
        .ProductName = "�R�s�[�p�� A4 500��"
        .Category = "���Օi"
        .Price = 800
        .MinStock = 100
        .MaxStock = 500
        .CurrentStock = 450
        .Supplier = "�R�N��"
        .LastUpdated = Now
        .IsActive = True
    End With
    
    With products(9)
        .ProductID = "P009"
        .ProductName = "�{�[���y�� JETSTREAM 0.5mm"
        .Category = "���[��"
        .Price = 120
        .MinStock = 50
        .MaxStock = 200
        .CurrentStock = 25
        .Supplier = "�O�H���M"
        .LastUpdated = Now
        .IsActive = True
    End With
    
    With products(10)
        .ProductID = "P010"
        .ProductName = "�f�X�N���C�g LED Z-80"
        .Category = "�I�t�B�X�Ƌ�"
        .Price = 35000
        .MinStock = 10
        .MaxStock = 30
        .CurrentStock = 35
        .Supplier = "�R�c�Ɩ�"
        .LastUpdated = Now
        .IsActive = True
    End With
    
    ' �T���v���f�[�^��o�^
    Dim i As Long
    For i = 1 To 10
        If Not AddOrUpdateProduct(products(i)) Then
            MsgBox "�T���v���f�[�^�̍쐬�Ɏ��s���܂���: " & products(i).ProductID, vbCritical
            Exit Sub
        End If
    Next i
    
    XLog "�T���v���f�[�^�i10���j���쐬���܂���"
    Exit Sub
    
ErrorHandler:
    MsgBox "�T���v���f�[�^�쐬���ɃG���[���������܂���: " & Err.Description, vbCritical, "�G���["
End Sub

' �T���v������f�[�^�̍쐬
Public Sub CreateSampleTransactions()
    On Error GoTo ErrorHandler
    
    XLog "�T���v������f�[�^�𐶐���"
    
    ' �T���v������f�[�^�̔z��
    Dim transactions(1 To 15) As InventoryTransaction
    
    ' ���Ɏ���̃T���v��
    With transactions(1)
        .TransactionID = "TXN" & Format(Now - 10, "yyyymmddhhnnss") & "001"
        .ProductID = "P001"
        .TransactionType = "����"
        .Quantity = 5
        .TransactionDate = Now - 10
        .UserName = "�c�����Y"
        .Notes = "�������"
        .ReferenceNo = "PO-2024-001"
    End With
    
    With transactions(2)
        .TransactionID = "TXN" & Format(Now - 9, "yyyymmddhhnnss") & "002"
        .ProductID = "P002"
        .TransactionType = "����"
        .Quantity = 25
        .TransactionDate = Now - 9
        .UserName = "�����Ԏq"
        .Notes = "�ً}��["
        .ReferenceNo = "PO-2024-002"
    End With
    
    ' �o�Ɏ���̃T���v��
    With transactions(3)
        .TransactionID = "TXN" & Format(Now - 8, "yyyymmddhhnnss") & "003"
        .ProductID = "P001"
        .TransactionType = "�o��"
        .Quantity = 2
        .TransactionDate = Now - 8
        .UserName = "�R�c��Y"
        .Notes = "�����ւ̔z�z"
        .ReferenceNo = "REQ-2024-001"
    End With
    
    With transactions(4)
        .TransactionID = "TXN" & Format(Now - 7, "yyyymmddhhnnss") & "004"
        .ProductID = "P003"
        .TransactionType = "�o��"
        .Quantity = 1
        .TransactionDate = Now - 7
        .UserName = "������Y"
        .Notes = "�V�K�]�ƈ��z�u"
        .ReferenceNo = "REQ-2024-002"
    End With
    
    ' �I������̃T���v��
    With transactions(5)
        .TransactionID = "TXN" & Format(Now - 6, "yyyymmddhhnnss") & "005"
        .ProductID = "P008"
        .TransactionType = "�I��"
        .Quantity = 450
        .TransactionDate = Now - 6
        .UserName = "�I���ӔC��"
        .Notes = "�����I��"
        .ReferenceNo = "INV-2024-001"
    End With
    
    ' �ǉ��̃T���v���f�[�^�i�c��10���j
    Dim i As Long
    For i = 6 To 15
        With transactions(i)
            .TransactionID = "TXN" & Format(Now - (15 - i), "yyyymmddhhnnss") & Format(i, "000")
            .ProductID = "P" & Format((i Mod 10) + 1, "000")
            Select Case (i Mod 3)
                Case 0: .TransactionType = "����"
                Case 1: .TransactionType = "�o��"
                Case 2: .TransactionType = "�I��"
            End Select
            .Quantity = (i Mod 50) + 1
            .TransactionDate = Now - (15 - i)
            .UserName = "�V�X�e���Ǘ���"
            .Notes = "�T���v���f�[�^"
            .ReferenceNo = "SAMPLE-" & i
        End With
    Next i
    
    ' �T���v���f�[�^��o�^
    For i = 1 To 15
        If Not RecordInventoryTransaction(transactions(i)) Then
            MsgBox "�T���v������f�[�^�̍쐬�Ɏ��s���܂���: " & transactions(i).TransactionID, vbCritical
            Exit Sub
        End If
    Next i
    
    XLog "�T���v������f�[�^�i15���j���쐬���܂���"
    Exit Sub
    
ErrorHandler:
    MsgBox "�T���v������f�[�^�쐬���ɃG���[���������܂���: " & Err.Description, vbCritical, "�G���["
End Sub
