VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'===================================================================
' 在庫管理シート - 入庫・出庫・棚卸管理
'===================================================================

Private Const FORM_START_ROW As Long = 2
Private Const FORM_START_COL As Long = 12 ' L列から開始

'===================================================================
' ワークシートイベント
'===================================================================

' シートアクティベート時の処理
Private Sub Worksheet_Activate()
    On Error Resume Next
    
    XLog "在庫管理シートがアクティベートされました"
    
    ' シート名を設定
    If Me.Name <> "在庫管理" Then
        Me.Name = "在庫管理"
    End If
    
    ' 入力フォームの初期化
    Call InitializeInventoryForm
    
    ' データの表示
    Call RefreshInventoryList
End Sub

' セル変更時の処理
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    
    Application.EnableEvents = False
    
    ' 取引履歴データ領域の変更をチェック
    If Not Intersect(Target, Range("A:H")) Is Nothing And Target.Row > 1 Then
        XLog "取引履歴データが変更されました: " & Target.Address
    End If
    
    Application.EnableEvents = True
End Sub

'===================================================================
' 入力フォーム機能
'===================================================================

' 在庫管理入力フォームの初期化
Private Sub InitializeInventoryForm()
    On Error Resume Next
    
    XLog "在庫管理入力フォームを初期化中"
    
    ' フォームヘッダーの作成
    With Me
        .Cells(1, FORM_START_COL).Value = "在庫取引入力フォーム"
        .Cells(1, FORM_START_COL).Font.Bold = True
        .Cells(1, FORM_START_COL).Font.Size = 12
        .Cells(1, FORM_START_COL).Interior.Color = RGB(200, 255, 200)
        
        ' 入力項目ラベルとテキストボックス
        .Cells(2, FORM_START_COL).Value = "取引ID:"
        .Cells(2, FORM_START_COL + 1).Name = "TransactionID_Input"
        .Cells(2, FORM_START_COL + 1).Value = "TXN" & Format(Now, "yyyymmddhhnnss")
        
        .Cells(3, FORM_START_COL).Value = "商品ID:"
        .Cells(3, FORM_START_COL + 1).Name = "ProductID_Input"
        
        .Cells(4, FORM_START_COL).Value = "取引種別:"
        .Cells(4, FORM_START_COL + 1).Name = "TransactionType_Input"
        
        .Cells(5, FORM_START_COL).Value = "数量:"
        .Cells(5, FORM_START_COL + 1).Name = "Quantity_Input"
        
        .Cells(6, FORM_START_COL).Value = "担当者:"
        .Cells(6, FORM_START_COL + 1).Name = "UserName_Input"
        .Cells(6, FORM_START_COL + 1).Value = Application.UserName
        
        .Cells(7, FORM_START_COL).Value = "備考:"
        .Cells(7, FORM_START_COL + 1).Name = "Notes_Input"
        
        .Cells(8, FORM_START_COL).Value = "参照番号:"
        .Cells(8, FORM_START_COL + 1).Name = "ReferenceNo_Input"
        
        ' 取引種別のドロップダウン設定
        With .Cells(4, FORM_START_COL + 1).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Formula1:="入庫,出庫,棚卸"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        .Cells(4, FORM_START_COL + 1).Value = "入庫"
        
        ' ボタンエリア
        .Cells(10, FORM_START_COL).Value = "取引実行"
        .Cells(10, FORM_START_COL).Name = "ExecuteTransactionButton"
        CreateInventoryButton "ExecuteTransactionButton", "Sheet2.ExecuteTransaction_Click"
        
        .Cells(10, FORM_START_COL + 1).Value = "クリア"
        .Cells(10, FORM_START_COL + 1).Name = "ClearInventoryButton"
        CreateInventoryButton "ClearInventoryButton", "Sheet2.ClearInventoryForm_Click"
        
        .Cells(10, FORM_START_COL + 2).Value = "商品検索"
        .Cells(10, FORM_START_COL + 2).Name = "ProductLookupButton"
        CreateInventoryButton "ProductLookupButton", "Sheet2.ProductLookup_Click"
        
        ' 現在在庫表示エリア
        .Cells(12, FORM_START_COL).Value = "現在在庫情報:"
        .Cells(12, FORM_START_COL).Font.Bold = True
        .Cells(12, FORM_START_COL).Interior.Color = RGB(255, 255, 200)
        
        .Cells(13, FORM_START_COL).Value = "商品名:"
        .Cells(13, FORM_START_COL + 1).Name = "CurrentProductName_Display"
        
        .Cells(14, FORM_START_COL).Value = "現在在庫:"
        .Cells(14, FORM_START_COL + 1).Name = "CurrentStock_Display"
        
        .Cells(15, FORM_START_COL).Value = "最小在庫:"
        .Cells(15, FORM_START_COL + 1).Name = "MinStock_Display"
        
        .Cells(16, FORM_START_COL).Value = "最大在庫:"
        .Cells(16, FORM_START_COL + 1).Name = "MaxStock_Display"
        
        ' 一括処理エリア
        .Cells(18, FORM_START_COL).Value = "一括処理:"
        .Cells(18, FORM_START_COL).Font.Bold = True
        .Cells(18, FORM_START_COL).Interior.Color = RGB(220, 220, 255)
        
        .Cells(19, FORM_START_COL).Value = "一括入庫"
        .Cells(19, FORM_START_COL).Name = "BulkReceiveButton"
        CreateInventoryButton "BulkReceiveButton", "Sheet2.BulkReceive_Click"
        
        .Cells(19, FORM_START_COL + 1).Value = "一括出庫"
        .Cells(19, FORM_START_COL + 1).Name = "BulkIssueButton"
        CreateInventoryButton "BulkIssueButton", "Sheet2.BulkIssue_Click"
        
        .Cells(19, FORM_START_COL + 2).Value = "一括棚卸"
        .Cells(19, FORM_START_COL + 2).Name = "BulkCountButton"
        CreateInventoryButton "BulkCountButton", "Sheet2.BulkCount_Click"
        
        ' フォーム範囲の書式設定
        With .Range(.Cells(1, FORM_START_COL), .Cells(21, FORM_START_COL + 3))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
        
        ' 入力欄の背景色
        With .Range(.Cells(2, FORM_START_COL + 1), .Cells(8, FORM_START_COL + 1))
            .Interior.Color = RGB(255, 255, 240)
        End With
        
        ' 表示欄の背景色
        With .Range(.Cells(13, FORM_START_COL + 1), .Cells(16, FORM_START_COL + 1))
            .Interior.Color = RGB(240, 255, 240)
        End With
    End With
    
    XLog "在庫管理入力フォームの初期化が完了しました"
End Sub

' ボタンの作成
Private Sub CreateInventoryButton(cellName As String, macroName As String)
    On Error Resume Next
    
    Dim btn As Button
    Dim targetRange As Range
    Set targetRange = Me.Range(cellName)
    
    ' 既存のボタンを削除
    Me.Buttons(cellName & "_Btn").Delete
    
    ' 新しいボタンを作成
    Set btn = Me.Buttons.Add(targetRange.Left, targetRange.Top, targetRange.Width, targetRange.Height)
    btn.Name = cellName & "_Btn"
    btn.Caption = targetRange.Value
    btn.OnAction = macroName
    
    ' ボタンの書式設定
    btn.Font.Bold = True
    btn.Font.Size = 9
End Sub

'===================================================================
' ボタンイベント処理
'===================================================================

' 取引実行ボタン
Public Sub ExecuteTransaction_Click()
    On Error GoTo ErrorHandler
    
    XLog "取引実行ボタンがクリックされました"
    
    ' 入力値の取得
    Dim transaction As InventoryTransaction
    With transaction
        .TransactionID = Trim(Range("TransactionID_Input").Value)
        .ProductID = Trim(Range("ProductID_Input").Value)
        .TransactionType = Trim(Range("TransactionType_Input").Value)
        .Quantity = CLng(Range("Quantity_Input").Value)
        .TransactionDate = Now
        .UserName = Trim(Range("UserName_Input").Value)
        .Notes = Trim(Range("Notes_Input").Value)
        .ReferenceNo = Trim(Range("ReferenceNo_Input").Value)
    End With
    
    ' 取引の実行
    If RecordInventoryTransaction(transaction) Then
        MsgBox "在庫取引を正常に実行しました。", vbInformation, "在庫管理"
        Call ClearInventoryForm_Click
        Call RefreshInventoryList
        Call UpdateCurrentStockDisplay(transaction.ProductID)
    Else
        MsgBox "在庫取引の実行に失敗しました。", vbCritical, "在庫管理"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "取引処理中にエラーが発生しました: " & Err.Description, vbCritical, "在庫管理エラー"
    XLog "取引実行でエラー: " & Err.Description, "ERROR"
End Sub

' フォームクリアボタン
Public Sub ClearInventoryForm_Click()
    On Error Resume Next
    
    With Me
        .Range("TransactionID_Input").Value = "TXN" & Format(Now, "yyyymmddhhnnss")
        .Range("ProductID_Input").Value = ""
        .Range("TransactionType_Input").Value = "入庫"
        .Range("Quantity_Input").Value = ""
        .Range("UserName_Input").Value = Application.UserName
        .Range("Notes_Input").Value = ""
        .Range("ReferenceNo_Input").Value = ""
        
        ' 現在在庫表示をクリア
        .Range("CurrentProductName_Display").Value = ""
        .Range("CurrentStock_Display").Value = ""
        .Range("MinStock_Display").Value = ""
        .Range("MaxStock_Display").Value = ""
    End With
    
    XLog "在庫管理入力フォームをクリアしました"
End Sub

' 商品検索ボタン
Public Sub ProductLookup_Click()
    On Error GoTo ErrorHandler
    
    Dim searchTerm As String
    searchTerm = InputBox("商品ID または 商品名を入力してください:", "商品検索")
    
    If searchTerm = "" Then Exit Sub
    
    ' 商品検索の実行
    Dim results As Collection
    Set results = SearchProducts(searchTerm)
    
    If results.Count = 0 Then
        MsgBox "該当する商品が見つかりませんでした。", vbInformation, "商品検索"
        Exit Sub
    ElseIf results.Count = 1 Then
        ' 1件の場合は直接設定
        Dim product As ProductInfo
        Set product = results(1)
        Range("ProductID_Input").Value = product.ProductID
        Call UpdateCurrentStockDisplay(product.ProductID)
    Else
        ' 複数件の場合は選択ダイアログを表示
        Call ShowProductSelectionDialog(results)
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "商品検索中にエラーが発生しました: " & Err.Description, vbCritical, "検索エラー"
End Sub

' 一括入庫ボタン
Public Sub BulkReceive_Click()
    On Error GoTo ErrorHandler
    
    Dim filePath As String
    filePath = Application.GetOpenFilename("CSV Files (*.csv), *.csv", , "一括入庫CSVファイルを選択")
    
    If filePath = "False" Then Exit Sub
    
    XLog "一括入庫処理を開始: " & filePath
    Call ProcessBulkInventoryFile(filePath, "入庫")
    
    Exit Sub
    
ErrorHandler:
    MsgBox "一括入庫処理中にエラーが発生しました: " & Err.Description, vbCritical, "一括処理エラー"
End Sub

' 一括出庫ボタン
Public Sub BulkIssue_Click()
    On Error GoTo ErrorHandler
    
    Dim filePath As String
    filePath = Application.GetOpenFilename("CSV Files (*.csv), *.csv", , "一括出庫CSVファイルを選択")
    
    If filePath = "False" Then Exit Sub
    
    XLog "一括出庫処理を開始: " & filePath
    Call ProcessBulkInventoryFile(filePath, "出庫")
    
    Exit Sub
    
ErrorHandler:
    MsgBox "一括出庫処理中にエラーが発生しました: " & Err.Description, vbCritical, "一括処理エラー"
End Sub

' 一括棚卸ボタン
Public Sub BulkCount_Click()
    On Error GoTo ErrorHandler
    
    Dim filePath As String
    filePath = Application.GetOpenFilename("CSV Files (*.csv), *.csv", , "一括棚卸CSVファイルを選択")
    
    If filePath = "False" Then Exit Sub
    
    XLog "一括棚卸処理を開始: " & filePath
    Call ProcessBulkInventoryFile(filePath, "棚卸")
    
    Exit Sub
    
ErrorHandler:
    MsgBox "一括棚卸処理中にエラーが発生しました: " & Err.Description, vbCritical, "一括処理エラー"
End Sub

'===================================================================
' データ表示・更新機能
'===================================================================

' 在庫履歴リストの更新
Private Sub RefreshInventoryList()
    On Error Resume Next
    
    XLog "在庫履歴リストを更新中"
    
    ' ヘッダーの設定
    If Cells(1, 1).Value = "" Then
        Call CreateTransactionHistoryHeaders(Me)
    End If
    
    ' データ領域の書式設定
    With Range("A1:H1")
        .Font.Bold = True
        .Interior.Color = RGB(180, 220, 180)
        .Borders.LineStyle = xlContinuous
    End With
    
    ' 列幅の調整
    Columns("A:H").AutoFit
    
    XLog "在庫履歴リストの更新が完了しました"
End Sub

' 現在在庫表示の更新
Private Sub UpdateCurrentStockDisplay(productID As String)
    On Error Resume Next
    
    If productID = "" Then Exit Sub
    
    Dim product As ProductInfo
    product = GetProduct(productID)
    
    If product.ProductID <> "" Then
        With Me
            .Range("CurrentProductName_Display").Value = product.ProductName
            .Range("CurrentStock_Display").Value = product.CurrentStock
            .Range("MinStock_Display").Value = product.MinStock
            .Range("MaxStock_Display").Value = product.MaxStock
            
            ' 在庫状況に応じた色分け
            If product.CurrentStock <= product.MinStock Then
                .Range("CurrentStock_Display").Interior.Color = RGB(255, 200, 200)
            ElseIf product.CurrentStock >= product.MaxStock Then
                .Range("CurrentStock_Display").Interior.Color = RGB(255, 255, 200)
            Else
                .Range("CurrentStock_Display").Interior.Color = RGB(200, 255, 200)
            End If
        End With
    End If
End Sub

' 商品選択ダイアログの表示
Private Sub ShowProductSelectionDialog(results As Collection)
    On Error Resume Next
    
    Dim selectionList As String
    Dim i As Long
    i = 1
    
    Dim product As ProductInfo
    For Each product In results
        selectionList = selectionList & i & ": " & product.ProductID & " - " & product.ProductName & vbCrLf
        i = i + 1
    Next product
    
    Dim selectedIndex As String
    selectedIndex = InputBox("選択してください（番号を入力）:" & vbCrLf & vbCrLf & selectionList, "商品選択")
    
    If IsNumeric(selectedIndex) And CLng(selectedIndex) >= 1 And CLng(selectedIndex) <= results.Count Then
        Set product = results(CLng(selectedIndex))
        Range("ProductID_Input").Value = product.ProductID
        Call UpdateCurrentStockDisplay(product.ProductID)
    End If
End Sub

' 一括処理ファイルの処理
Private Sub ProcessBulkInventoryFile(filePath As String, transactionType As String)
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open filePath For Input As #fileNum
    
    Dim lineText As String
    Dim lineNumber As Long
    Dim successCount As Long
    Dim errorCount As Long
    
    lineNumber = 0
    
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        lineNumber = lineNumber + 1
        
        ' ヘッダー行をスキップ
        If lineNumber = 1 Then
            GoTo ContinueLoop
        End If
        
        ' CSV行の解析
        Dim fields As Variant
        fields = Split(lineText, ",")
        
        If UBound(fields) >= 2 Then
            Dim transaction As InventoryTransaction
            With transaction
                .TransactionID = "BULK" & Format(Now, "yyyymmddhhnnss") & "_" & lineNumber
                .ProductID = Trim(fields(0))
                .TransactionType = transactionType
                .Quantity = CLng(Trim(fields(1)))
                .TransactionDate = Now
                .UserName = Application.UserName & " (一括処理)"
                .Notes = "一括" & transactionType & " - ファイル: " & Dir(filePath)
                If UBound(fields) >= 2 Then
                    .ReferenceNo = Trim(fields(2))
                End If
            End With
            
            If RecordInventoryTransaction(transaction) Then
                successCount = successCount + 1
            Else
                errorCount = errorCount + 1
                XLog "一括処理エラー - 行" & lineNumber & ": " & lineText, "ERROR"
            End If
        End If
        
ContinueLoop:
    Loop
    
    Close #fileNum
    
    Call RefreshInventoryList
    
    Dim message As String
    message = "一括" & transactionType & "処理が完了しました。" & vbCrLf & vbCrLf
    message = message & "成功: " & successCount & " 件" & vbCrLf
    message = message & "エラー: " & errorCount & " 件"
    
    MsgBox message, vbInformation, "一括処理結果"
    
    XLog "一括処理完了 - 成功: " & successCount & ", エラー: " & errorCount
    Exit Sub
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    MsgBox "一括処理ファイルの読み込み中にエラーが発生しました: " & Err.Description, vbCritical, "一括処理エラー"
    XLog "一括処理ファイルエラー: " & Err.Description, "ERROR"
End Sub

' ProductIDが変更された時の自動処理
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next
    
    ' ProductID_Input セルが選択されている場合
    If Target.Address = Range("ProductID_Input").Address Then
        Dim productID As String
        productID = Trim(Target.Value)
        
        If productID <> "" And Len(productID) >= 3 Then
            Call UpdateCurrentStockDisplay(productID)
        End If
    End If
End Sub

'===================================================================
' サンプル取引データ生成
'===================================================================

' サンプル取引データの作成
Public Sub CreateSampleTransactions()
    On Error GoTo ErrorHandler
    
    XLog "サンプル取引データを生成中"
    
    ' サンプル取引データの配列
    Dim transactions(1 To 15) As InventoryTransaction
    
    ' 入庫取引のサンプル
    With transactions(1)
        .TransactionID = "TXN" & Format(Now - 10, "yyyymmddhhnnss") & "001"
        .ProductID = "P001"
        .TransactionType = "入庫"
        .Quantity = 5
        .TransactionDate = Now - 10
        .UserName = "田中太郎"
        .Notes = "定期発注分"
        .ReferenceNo = "PO-2024-001"
    End With
    
    With transactions(2)
        .TransactionID = "TXN" & Format(Now - 9, "yyyymmddhhnnss") & "002"
        .ProductID = "P002"
        .TransactionType = "入庫"
        .Quantity = 25
        .TransactionDate = Now - 9
        .UserName = "佐藤花子"
        .Notes = "緊急発注分"
        .ReferenceNo = "PO-2024-002"
    End With
    
    ' 出庫取引のサンプル
    With transactions(3)
        .TransactionID = "TXN" & Format(Now - 8, "yyyymmddhhnnss") & "003"
        .ProductID = "P001"
        .TransactionType = "出庫"
        .Quantity = 2
        .TransactionDate = Now - 8
        .UserName = "山田次郎"
        .Notes = "営業部使用分"
        .ReferenceNo = "REQ-2024-001"
    End With
    
    With transactions(4)
        .TransactionID = "TXN" & Format(Now - 7, "yyyymmddhhnnss") & "004"
        .ProductID = "P003"
        .TransactionType = "出庫"
        .Quantity = 1
        .TransactionDate = Now - 7
        .UserName = "鈴木一郎"
        .Notes = "新入社員配備"
        .ReferenceNo = "REQ-2024-002"
    End With
    
    ' 棚卸取引のサンプル
    With transactions(5)
        .TransactionID = "TXN" & Format(Now - 6, "yyyymmddhhnnss") & "005"
        .ProductID = "P008"
        .TransactionType = "棚卸"
        .Quantity = 450
        .TransactionDate = Now - 6
        .UserName = "棚卸担当者"
        .Notes = "月次棚卸"
        .ReferenceNo = "INV-2024-001"
    End With
    
    ' 追加のサンプルデータ（残り10件）
    Dim i As Long
    For i = 6 To 15
        With transactions(i)
            .TransactionID = "TXN" & Format(Now - (15 - i), "yyyymmddhhnnss") & Format(i, "000")
            .ProductID = "P" & Format((i Mod 10) + 1, "000")
            Select Case (i Mod 3)
                Case 0: .TransactionType = "入庫"
                Case 1: .TransactionType = "出庫"
                Case 2: .TransactionType = "棚卸"
            End Select
            .Quantity = (i Mod 50) + 1
            .TransactionDate = Now - (15 - i)
            .UserName = "システム生成"
            .Notes = "サンプルデータ"
            .ReferenceNo = "SAMPLE-" & i
        End With
    Next i
    
    ' サンプルデータを登録
    For i = 1 To 15
        If Not RecordInventoryTransaction(transactions(i)) Then
            MsgBox "サンプル取引データの作成に失敗しました: " & transactions(i).TransactionID, vbCritical
            Exit Sub
        End If
    Next i
    
    Call RefreshInventoryList
    MsgBox "サンプル取引データ（15件）を作成しました。", vbInformation, "在庫管理"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "サンプル取引データ作成中にエラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub