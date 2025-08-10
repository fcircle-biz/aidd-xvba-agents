Attribute VB_Name = "common"

'===================================================================
' 商品在庫管理システム - 共通モジュール
'===================================================================

' グローバル変数
Public Const PRODUCT_SHEET As String = "商品マスタ"
Public Const INVENTORY_SHEET As String = "在庫管理"
Public Const TRANSACTION_SHEET As String = "取引履歴"
Public Const REPORT_SHEET As String = "レポート"

' 商品管理関連の型定義
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
    TransactionType As String  ' "入庫", "出庫", "棚卸"
    Quantity As Long
    TransactionDate As Date
    UserName As String
    Notes As String
    ReferenceNo As String
End Type

'===================================================================
' 共通ユーティリティ関数
'===================================================================

' Xdebugとの統合
Public Sub XLog(ByVal message As String, Optional ByVal level As String = "INFO")
    On Error Resume Next
    
    ' Xdebugが利用可能な場合は使用
    If IsXdebugAvailable() Then
        Dim xd As Object
        Set xd = CreateObject("Xdebug.Logger")
        xd.Log level & ": " & message
    End If
    
    ' ログファイルにも出力
    LogToFile message, level
End Sub

' Xdebugの利用可能性チェック
Public Function IsXdebugAvailable() As Boolean
    On Error GoTo ErrorHandler
    Dim xd As Object
    Set xd = CreateObject("Xdebug.Logger")
    IsXdebugAvailable = True
    Exit Function
    
ErrorHandler:
    IsXdebugAvailable = False
End Function

' ファイルログ出力
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
' 商品マスタ管理機能
'===================================================================

' 商品情報を追加/更新
Public Function AddOrUpdateProduct(product As ProductInfo) As Boolean
    On Error GoTo ErrorHandler
    
    XLog "商品情報の追加/更新を開始: " & product.ProductID
    
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(PRODUCT_SHEET)
    
    ' ヘッダー行の確認・作成
    If ws.Cells(1, 1).Value = "" Then
        Call CreateProductMasterHeaders(ws)
    End If
    
    ' 既存商品のチェック
    Dim targetRow As Long
    targetRow = FindProductRow(product.ProductID)
    
    If targetRow = 0 Then
        ' 新規追加
        targetRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        XLog "新規商品として追加: 行" & targetRow
    Else
        XLog "既存商品を更新: 行" & targetRow
    End If
    
    ' データ設定
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
    XLog "商品情報の追加/更新が完了しました"
    Exit Function
    
ErrorHandler:
    XLog "商品情報の追加/更新でエラーが発生: " & Err.Description, "ERROR"
    AddOrUpdateProduct = False
End Function

' 商品マスタのヘッダー作成
Public Sub CreateProductMasterHeaders(ws As Worksheet)
    With ws
        .Cells(1, 1).Value = "商品ID"
        .Cells(1, 2).Value = "商品名"
        .Cells(1, 3).Value = "カテゴリ"
        .Cells(1, 4).Value = "価格"
        .Cells(1, 5).Value = "最小在庫"
        .Cells(1, 6).Value = "最大在庫"
        .Cells(1, 7).Value = "現在在庫"
        .Cells(1, 8).Value = "仕入先"
        .Cells(1, 9).Value = "更新日時"
        .Cells(1, 10).Value = "有効フラグ"
        
        ' ヘッダー行の書式設定
        With .Range("A1:J1")
            .Font.Bold = True
            .Interior.Color = RGB(200, 200, 200)
            .Borders.LineStyle = xlContinuous
        End With
    End With
    
    XLog "商品マスタのヘッダーを作成しました"
End Sub

' 商品IDから行番号を検索
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

' 商品情報を取得
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
        
        XLog "商品情報を取得しました: " & productID
    Else
        XLog "商品が見つかりません: " & productID, "WARNING"
    End If
    
    GetProduct = product
    Exit Function
    
ErrorHandler:
    XLog "商品情報取得でエラーが発生: " & Err.Description, "ERROR"
End Function

'===================================================================
' ワークシート管理
'===================================================================

' シートを取得または作成
Public Function GetOrCreateSheet(sheetName As String) As Worksheet
    On Error Resume Next
    
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    
    If GetOrCreateSheet Is Nothing Then
        ' 既存のSheet1-9を使用してシート名を変更
        Dim availableSheet As Worksheet
        Dim i As Integer
        For i = 1 To 9
            Set availableSheet = ThisWorkbook.Worksheets("Sheet" & i)
            If Not availableSheet Is Nothing And availableSheet.Cells(1, 1).Value = "" Then
                availableSheet.Name = sheetName
                Set GetOrCreateSheet = availableSheet
                XLog "Sheet" & i & "を" & sheetName & "に変更しました"
                Exit For
            End If
        Next i
        
        ' 利用可能なシートが見つからない場合は最初のシートを使用
        If GetOrCreateSheet Is Nothing Then
            Set GetOrCreateSheet = ThisWorkbook.Worksheets("Sheet1")
            GetOrCreateSheet.Name = sheetName
            XLog "Sheet1を" & sheetName & "に変更しました"
        End If
    End If
    
    On Error GoTo 0
End Function

' データ検証機能
Public Function ValidateProductData(product As ProductInfo) As String
    Dim errorMessage As String
    
    ' 必須フィールドのチェック
    If Trim(product.ProductID) = "" Then
        errorMessage = errorMessage & "商品IDは必須です。" & vbCrLf
    End If
    
    If Trim(product.ProductName) = "" Then
        errorMessage = errorMessage & "商品名は必須です。" & vbCrLf
    End If
    
    ' 数値データの妥当性チェック
    If product.Price < 0 Then
        errorMessage = errorMessage & "価格は0以上である必要があります。" & vbCrLf
    End If
    
    If product.MinStock < 0 Then
        errorMessage = errorMessage & "最小在庫は0以上である必要があります。" & vbCrLf
    End If
    
    If product.MaxStock < product.MinStock Then
        errorMessage = errorMessage & "最大在庫は最小在庫以上である必要があります。" & vbCrLf
    End If
    
    If product.CurrentStock < 0 Then
        errorMessage = errorMessage & "現在在庫は0以上である必要があります。" & vbCrLf
    End If
    
    ValidateProductData = errorMessage
End Function

'===================================================================
' 在庫管理機能
'===================================================================

' 在庫取引を記録
Public Function RecordInventoryTransaction(transaction As InventoryTransaction) As Boolean
    On Error GoTo ErrorHandler
    
    XLog "在庫取引の記録を開始: " & transaction.TransactionID
    
    ' データ検証
    Dim validationError As String
    validationError = ValidateTransactionData(transaction)
    If validationError <> "" Then
        MsgBox "データエラー:" & vbCrLf & validationError, vbCritical, "在庫管理エラー"
        RecordInventoryTransaction = False
        Exit Function
    End If
    
    ' 在庫更新の実行
    If Not UpdateProductStock(transaction) Then
        RecordInventoryTransaction = False
        Exit Function
    End If
    
    ' 取引履歴の記録
    If Not LogTransactionHistory(transaction) Then
        RecordInventoryTransaction = False
        Exit Function
    End If
    
    RecordInventoryTransaction = True
    XLog "在庫取引の記録が完了しました: " & transaction.TransactionID
    Exit Function
    
ErrorHandler:
    XLog "在庫取引記録でエラーが発生: " & Err.Description, "ERROR"
    RecordInventoryTransaction = False
End Function

' 商品在庫を更新
Public Function UpdateProductStock(transaction As InventoryTransaction) As Boolean
    On Error GoTo ErrorHandler
    
    Dim productRow As Long
    productRow = FindProductRow(transaction.ProductID)
    
    If productRow = 0 Then
        MsgBox "商品ID " & transaction.ProductID & " が見つかりません。", vbCritical, "在庫管理エラー"
        UpdateProductStock = False
        Exit Function
    End If
    
    Dim ws As Worksheet
    Set ws = Worksheets(PRODUCT_SHEET)
    
    Dim currentStock As Long
    currentStock = ws.Cells(productRow, 7).Value
    
    Dim newStock As Long
    Select Case transaction.TransactionType
        Case "入庫"
            newStock = currentStock + transaction.Quantity
        Case "出庫"
            newStock = currentStock - transaction.Quantity
            If newStock < 0 Then
                MsgBox "在庫不足です。現在在庫: " & currentStock & ", 要求数量: " & transaction.Quantity, vbCritical, "在庫管理エラー"
                UpdateProductStock = False
                Exit Function
            End If
        Case "棚卸"
            newStock = transaction.Quantity ' 棚卸の場合は絶対値
        Case Else
            MsgBox "無効な取引種別です: " & transaction.TransactionType, vbCritical, "在庫管理エラー"
            UpdateProductStock = False
            Exit Function
    End Select
    
    ' 在庫数の更新
    ws.Cells(productRow, 7).Value = newStock
    ws.Cells(productRow, 9).Value = Now ' 更新日時
    
    ' 在庫アラートのチェック
    CheckStockAlerts transaction.ProductID, newStock
    
    UpdateProductStock = True
    XLog "在庫を更新しました - 商品ID: " & transaction.ProductID & ", 新在庫: " & newStock
    Exit Function
    
ErrorHandler:
    XLog "在庫更新でエラーが発生: " & Err.Description, "ERROR"
    UpdateProductStock = False
End Function

' 取引履歴の記録
Public Function LogTransactionHistory(transaction As InventoryTransaction) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(TRANSACTION_SHEET)
    
    ' ヘッダー行の確認・作成
    If ws.Cells(1, 1).Value = "" Then
        Call CreateTransactionHistoryHeaders(ws)
    End If
    
    ' 新しい行にデータを追加
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
    XLog "取引履歴を記録しました: " & transaction.TransactionID
    Exit Function
    
ErrorHandler:
    XLog "取引履歴記録でエラーが発生: " & Err.Description, "ERROR"
    LogTransactionHistory = False
End Function

' 取引履歴のヘッダー作成
Public Sub CreateTransactionHistoryHeaders(ws As Worksheet)
    With ws
        .Cells(1, 1).Value = "取引ID"
        .Cells(1, 2).Value = "商品ID"
        .Cells(1, 3).Value = "取引種別"
        .Cells(1, 4).Value = "数量"
        .Cells(1, 5).Value = "取引日時"
        .Cells(1, 6).Value = "担当者"
        .Cells(1, 7).Value = "備考"
        .Cells(1, 8).Value = "参照番号"
        
        ' ヘッダー行の書式設定
        With .Range("A1:H1")
            .Font.Bold = True
            .Interior.Color = RGB(180, 220, 180)
            .Borders.LineStyle = xlContinuous
        End With
    End With
    
    XLog "取引履歴のヘッダーを作成しました"
End Sub

' 在庫アラートのチェック
Public Sub CheckStockAlerts(productID As String, currentStock As Long)
    On Error Resume Next
    
    Dim product As ProductInfo
    product = GetProduct(productID)
    
    If currentStock <= product.MinStock Then
        XLog "在庫アラート: " & product.ProductName & " (ID: " & productID & ") の在庫が最小値を下回りました。現在在庫: " & currentStock, "WARNING"
        
        ' 在庫アラート画面の表示（オプション）
        ShowStockAlert product, currentStock
    ElseIf currentStock >= product.MaxStock Then
        XLog "在庫過多アラート: " & product.ProductName & " (ID: " & productID & ") の在庫が最大値を超えました。現在在庫: " & currentStock, "WARNING"
    End If
End Sub

' 在庫アラート画面の表示
Public Sub ShowStockAlert(product As ProductInfo, currentStock As Long)
    Dim message As String
    message = "在庫アラート" & vbCrLf & vbCrLf
    message = message & "商品名: " & product.ProductName & vbCrLf
    message = message & "商品ID: " & product.ProductID & vbCrLf
    message = message & "現在在庫: " & currentStock & vbCrLf
    message = message & "最小在庫: " & product.MinStock & vbCrLf
    message = message & "推奨発注数量: " & (product.MaxStock - currentStock)
    
    MsgBox message, vbExclamation, "在庫管理システム"
End Sub

' 取引データの検証
Public Function ValidateTransactionData(transaction As InventoryTransaction) As String
    Dim errorMessage As String
    
    ' 必須フィールドのチェック
    If Trim(transaction.TransactionID) = "" Then
        errorMessage = errorMessage & "取引IDは必須です。" & vbCrLf
    End If
    
    If Trim(transaction.ProductID) = "" Then
        errorMessage = errorMessage & "商品IDは必須です。" & vbCrLf
    End If
    
    If Trim(transaction.TransactionType) = "" Then
        errorMessage = errorMessage & "取引種別は必須です。" & vbCrLf
    ElseIf transaction.TransactionType <> "入庫" And transaction.TransactionType <> "出庫" And transaction.TransactionType <> "棚卸" Then
        errorMessage = errorMessage & "取引種別は「入庫」「出庫」「棚卸」のいずれかを指定してください。" & vbCrLf
    End If
    
    If transaction.Quantity <= 0 Then
        errorMessage = errorMessage & "数量は1以上を指定してください。" & vbCrLf
    End If
    
    If Trim(transaction.UserName) = "" Then
        errorMessage = errorMessage & "担当者名は必須です。" & vbCrLf
    End If
    
    ' 商品の存在チェック
    If FindProductRow(transaction.ProductID) = 0 Then
        errorMessage = errorMessage & "指定された商品ID「" & transaction.ProductID & "」が存在しません。" & vbCrLf
    End If
    
    ValidateTransactionData = errorMessage
End Function

'===================================================================
' 検索・フィルタリング機能
'===================================================================

' 商品検索機能
Public Function SearchProducts(searchTerm As String, Optional categoryFilter As String = "", Optional stockFilter As String = "") As Collection
    On Error GoTo ErrorHandler
    
    Set SearchProducts = New Collection
    
    Dim ws As Worksheet
    Set ws = Worksheets(PRODUCT_SHEET)
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 2 Then ' ヘッダー行のみの場合
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
        
        ' 検索条件のチェック
        Dim matchesSearch As Boolean
        matchesSearch = (searchTerm = "" Or _
                        InStr(1, product.ProductName, searchTerm, vbTextCompare) > 0 Or _
                        InStr(1, product.ProductID, searchTerm, vbTextCompare) > 0)
        
        Dim matchesCategory As Boolean
        matchesCategory = (categoryFilter = "" Or product.Category = categoryFilter)
        
        Dim matchesStock As Boolean
        Select Case stockFilter
            Case "低在庫"
                matchesStock = (product.CurrentStock <= product.MinStock)
            Case "過剰在庫"
                matchesStock = (product.CurrentStock >= product.MaxStock)
            Case "適正在庫"
                matchesStock = (product.CurrentStock > product.MinStock And product.CurrentStock < product.MaxStock)
            Case Else
                matchesStock = True
        End Select
        
        If matchesSearch And matchesCategory And matchesStock And product.IsActive Then
            SearchProducts.Add product, product.ProductID
        End If
    Next i
    
    XLog "商品検索を実行しました。検索語: " & searchTerm & ", 結果件数: " & SearchProducts.Count
    Exit Function
    
ErrorHandler:
    XLog "商品検索でエラーが発生: " & Err.Description, "ERROR"
    Set SearchProducts = New Collection
End Function

'===================================================================
' レポート・可視化機能
'===================================================================

' 在庫状況レポートの生成
Public Sub GenerateInventoryStatusReport()
    On Error GoTo ErrorHandler
    
    XLog "在庫状況レポート生成を開始"
    
    Dim reportWs As Worksheet
    Set reportWs = GetOrCreateSheet(REPORT_SHEET)
    
    ' レポートシートをクリア
    reportWs.Cells.Clear
    
    ' レポートヘッダーの作成
    CreateInventoryReportHeaders reportWs
    
    ' データ集計と出力
    Dim ws As Worksheet
    Set ws = Worksheets(PRODUCT_SHEET)
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 2 Then
        reportWs.Cells(3, 1).Value = "データがありません"
        Exit Sub
    End If
    
    Dim reportRow As Long
    reportRow = 3
    
    Dim totalValue As Double
    Dim lowStockCount As Long
    Dim overStockCount As Long
    
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 10).Value Then ' アクティブな商品のみ
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
                stockStatus = "低在庫"
                lowStockCount = lowStockCount + 1
            ElseIf currentStock >= maxStock Then
                stockStatus = "過剰在庫"
                overStockCount = overStockCount + 1
            Else
                stockStatus = "適正"
            End If
            
            totalValue = totalValue + (currentStock * price)
            
            ' レポート行にデータを出力
            With reportWs
                .Cells(reportRow, 1).Value = ws.Cells(i, 1).Value ' 商品ID
                .Cells(reportRow, 2).Value = ws.Cells(i, 2).Value ' 商品名
                .Cells(reportRow, 3).Value = ws.Cells(i, 3).Value ' カテゴリ
                .Cells(reportRow, 4).Value = currentStock
                .Cells(reportRow, 5).Value = minStock
                .Cells(reportRow, 6).Value = maxStock
                .Cells(reportRow, 7).Value = stockStatus
                .Cells(reportRow, 8).Value = currentStock * price ' 在庫金額
                
                ' 在庫状況に応じた色分け
                Select Case stockStatus
                    Case "低在庫"
                        .Cells(reportRow, 7).Interior.Color = RGB(255, 200, 200)
                    Case "過剰在庫"
                        .Cells(reportRow, 7).Interior.Color = RGB(255, 255, 200)
                    Case "適正"
                        .Cells(reportRow, 7).Interior.Color = RGB(200, 255, 200)
                End Select
            End With
            
            reportRow = reportRow + 1
        End If
    Next i
    
    ' サマリー情報の追加
    CreateReportSummary reportWs, reportRow + 2, totalValue, lowStockCount, overStockCount
    
    ' レポートの書式設定
    FormatInventoryReport reportWs, reportRow - 1
    
    XLog "在庫状況レポート生成が完了しました"
    MsgBox "在庫状況レポートが「" & REPORT_SHEET & "」シートに生成されました。", vbInformation, "在庫管理システム"
    Exit Sub
    
ErrorHandler:
    XLog "レポート生成でエラーが発生: " & Err.Description, "ERROR"
    MsgBox "レポート生成中にエラーが発生しました: " & Err.Description, vbCritical, "在庫管理エラー"
End Sub

' レポートヘッダーの作成
Public Sub CreateInventoryReportHeaders(ws As Worksheet)
    ws.Cells(1, 1).Value = "在庫状況レポート - " & Format(Now, "yyyy/mm/dd hh:nn")
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 14
    
    With ws
        .Cells(2, 1).Value = "商品ID"
        .Cells(2, 2).Value = "商品名"
        .Cells(2, 3).Value = "カテゴリ"
        .Cells(2, 4).Value = "現在在庫"
        .Cells(2, 5).Value = "最小在庫"
        .Cells(2, 6).Value = "最大在庫"
        .Cells(2, 7).Value = "在庫状況"
        .Cells(2, 8).Value = "在庫金額"
        
        With .Range("A2:H2")
            .Font.Bold = True
            .Interior.Color = RGB(200, 200, 200)
            .Borders.LineStyle = xlContinuous
        End With
    End With
End Sub

' レポートサマリーの作成
Public Sub CreateReportSummary(ws As Worksheet, startRow As Long, totalValue As Double, lowStockCount As Long, overStockCount As Long)
    With ws
        .Cells(startRow, 1).Value = "サマリー"
        .Cells(startRow, 1).Font.Bold = True
        
        .Cells(startRow + 1, 1).Value = "総在庫金額:"
        .Cells(startRow + 1, 2).Value = Format(totalValue, "#,##0")
        
        .Cells(startRow + 2, 1).Value = "低在庫商品数:"
        .Cells(startRow + 2, 2).Value = lowStockCount
        
        .Cells(startRow + 3, 1).Value = "過剰在庫商品数:"
        .Cells(startRow + 3, 2).Value = overStockCount
        
        .Range(.Cells(startRow, 1), .Cells(startRow + 3, 2)).Font.Bold = True
    End With
End Sub

' レポートの書式設定
Public Sub FormatInventoryReport(ws As Worksheet, lastDataRow As Long)
    With ws
        ' 列幅の自動調整
        .Columns("A:H").AutoFit
        
        ' 数値列の書式設定
        .Range("D3:F" & lastDataRow).NumberFormat = "#,##0"
        .Range("H3:H" & lastDataRow).NumberFormat = "#,##0"
        
        ' 全データ範囲に枠線を設定
        With .Range("A2:H" & lastDataRow)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
    End With
End Sub

'===================================================================
' サンプルデータ生成関数
'===================================================================

' サンプル商品データの作成
Public Sub CreateSampleData()
    On Error GoTo ErrorHandler
    
    XLog "サンプルデータを生成中"
    
    Dim products(1 To 10) As ProductInfo
    
    ' サンプルデータの定義
    With products(1)
        .ProductID = "P001"
        .ProductName = "ノートPC ThinkPad X1"
        .Category = "PC・周辺機器"
        .Price = 180000
        .MinStock = 5
        .MaxStock = 20
        .CurrentStock = 12
        .Supplier = "レノボ・ジャパン"
        .LastUpdated = Now
        .IsActive = True
    End With
    
    With products(2)
        .ProductID = "P002"
        .ProductName = "ワイヤレスマウス MX Master 3"
        .Category = "PC・周辺機器"
        .Price = 12000
        .MinStock = 10
        .MaxStock = 50
        .CurrentStock = 3
        .Supplier = "ロジクール"
        .LastUpdated = Now
        .IsActive = True
    End With
    
    With products(3)
        .ProductID = "P003"
        .ProductName = "オフィスチェア エルゴヒューマン"
        .Category = "オフィス家具"
        .Price = 120000
        .MinStock = 2
        .MaxStock = 10
        .CurrentStock = 8
        .Supplier = "オカムラ"
        .LastUpdated = Now
        .IsActive = True
    End With
    
    With products(4)
        .ProductID = "P004"
        .ProductName = "プリンタ複合機 MFC-L3770CDW"
        .Category = "PC・周辺機器"
        .Price = 45000
        .MinStock = 3
        .MaxStock = 15
        .CurrentStock = 18
        .Supplier = "ブラザー工業"
        .LastUpdated = Now
        .IsActive = True
    End With
    
    With products(5)
        .ProductID = "P005"
        .ProductName = "モニター 27インチ 4K"
        .Category = "PC・周辺機器"
        .Price = 85000
        .MinStock = 8
        .MaxStock = 25
        .CurrentStock = 15
        .Supplier = "デル"
        .LastUpdated = Now
        .IsActive = True
    End With
    
    With products(6)
        .ProductID = "P006"
        .ProductName = "ホワイトボード 1800×900"
        .Category = "オフィス用品"
        .Price = 25000
        .MinStock = 5
        .MaxStock = 15
        .CurrentStock = 2
        .Supplier = "コクヨ"
        .LastUpdated = Now
        .IsActive = True
    End With
    
    With products(7)
        .ProductID = "P007"
        .ProductName = "シュレッダー GCS280i"
        .Category = "オフィス機器"
        .Price = 180000
        .MinStock = 1
        .MaxStock = 5
        .CurrentStock = 3
        .Supplier = "フェローズ"
        .LastUpdated = Now
        .IsActive = True
    End With
    
    With products(8)
        .ProductID = "P008"
        .ProductName = "コピー用紙 A4 500枚"
        .Category = "消耗品"
        .Price = 800
        .MinStock = 100
        .MaxStock = 500
        .CurrentStock = 450
        .Supplier = "コクヨ"
        .LastUpdated = Now
        .IsActive = True
    End With
    
    With products(9)
        .ProductID = "P009"
        .ProductName = "ボールペン JETSTREAM 0.5mm"
        .Category = "文房具"
        .Price = 120
        .MinStock = 50
        .MaxStock = 200
        .CurrentStock = 25
        .Supplier = "三菱鉛筆"
        .LastUpdated = Now
        .IsActive = True
    End With
    
    With products(10)
        .ProductID = "P010"
        .ProductName = "デスクライト LED Z-80"
        .Category = "オフィス家具"
        .Price = 35000
        .MinStock = 10
        .MaxStock = 30
        .CurrentStock = 35
        .Supplier = "山田照明"
        .LastUpdated = Now
        .IsActive = True
    End With
    
    ' サンプルデータを登録
    Dim i As Long
    For i = 1 To 10
        If Not AddOrUpdateProduct(products(i)) Then
            MsgBox "サンプルデータの作成に失敗しました: " & products(i).ProductID, vbCritical
            Exit Sub
        End If
    Next i
    
    XLog "サンプルデータ（10件）を作成しました"
    Exit Sub
    
ErrorHandler:
    MsgBox "サンプルデータ作成中にエラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

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
        .Notes = "定期入荷"
        .ReferenceNo = "PO-2024-001"
    End With
    
    With transactions(2)
        .TransactionID = "TXN" & Format(Now - 9, "yyyymmddhhnnss") & "002"
        .ProductID = "P002"
        .TransactionType = "入庫"
        .Quantity = 25
        .TransactionDate = Now - 9
        .UserName = "佐藤花子"
        .Notes = "緊急補充"
        .ReferenceNo = "PO-2024-002"
    End With
    
    ' 出庫取引のサンプル
    With transactions(3)
        .TransactionID = "TXN" & Format(Now - 8, "yyyymmddhhnnss") & "003"
        .ProductID = "P001"
        .TransactionType = "出庫"
        .Quantity = 2
        .TransactionDate = Now - 8
        .UserName = "山田一郎"
        .Notes = "部署への配布"
        .ReferenceNo = "REQ-2024-001"
    End With
    
    With transactions(4)
        .TransactionID = "TXN" & Format(Now - 7, "yyyymmddhhnnss") & "004"
        .ProductID = "P003"
        .TransactionType = "出庫"
        .Quantity = 1
        .TransactionDate = Now - 7
        .UserName = "高橋二郎"
        .Notes = "新規従業員配置"
        .ReferenceNo = "REQ-2024-002"
    End With
    
    ' 棚卸取引のサンプル
    With transactions(5)
        .TransactionID = "TXN" & Format(Now - 6, "yyyymmddhhnnss") & "005"
        .ProductID = "P008"
        .TransactionType = "棚卸"
        .Quantity = 450
        .TransactionDate = Now - 6
        .UserName = "棚卸責任者"
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
            .UserName = "システム管理者"
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
    
    XLog "サンプル取引データ（15件）を作成しました"
    Exit Sub
    
ErrorHandler:
    MsgBox "サンプル取引データ作成中にエラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub
