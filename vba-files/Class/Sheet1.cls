VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'===================================================================
' 商品マスタ管理シート - イベントハンドラ
'===================================================================

Private Const FORM_START_ROW As Long = 2
Private Const FORM_START_COL As Long = 12 ' L列から開始

'===================================================================
' ワークシートイベント
'===================================================================

' シートアクティベート時の処理
Private Sub Worksheet_Activate()
    On Error Resume Next
    
    XLog "商品マスタシートがアクティベートされました"
    
    ' シート名を設定
    If Me.Name <> "商品マスタ" Then
        Me.Name = "商品マスタ"
    End If
    
    ' 入力フォームの初期化
    Call InitializeProductForm
    
    ' データの表示
    Call RefreshProductList
End Sub

' セル変更時の処理
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    
    Application.EnableEvents = False
    
    ' 商品データ領域の変更をチェック
    If Not Intersect(Target, Range("A:J")) Is Nothing And Target.Row > 1 Then
        XLog "商品データが変更されました: " & Target.Address
        
        ' 更新日時を自動設定
        If Target.Column <> 9 Then ' 更新日時列以外の変更の場合
            Cells(Target.Row, 9).Value = Now
        End If
    End If
    
    Application.EnableEvents = True
End Sub

' ダブルクリック時の処理
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    On Error Resume Next
    
    ' 商品データ行をダブルクリックした場合
    If Not Intersect(Target, Range("A:J")) Is Nothing And Target.Row > 1 Then
        Cancel = True
        Call EditProduct(Target.Row)
    End If
End Sub

'===================================================================
' 入力フォーム機能
'===================================================================

' 商品入力フォームの初期化
Private Sub InitializeProductForm()
    On Error Resume Next
    
    XLog "商品入力フォームを初期化中"
    
    ' フォームヘッダーの作成
    With Me
        .Cells(1, FORM_START_COL).Value = "商品情報入力フォーム"
        .Cells(1, FORM_START_COL).Font.Bold = True
        .Cells(1, FORM_START_COL).Font.Size = 12
        .Cells(1, FORM_START_COL).Interior.Color = RGB(200, 220, 255)
        
        ' 入力項目ラベルとテキストボックス
        .Cells(2, FORM_START_COL).Value = "商品ID:"
        .Cells(2, FORM_START_COL + 1).Name = "ProductID_Input"
        
        .Cells(3, FORM_START_COL).Value = "商品名:"
        .Cells(3, FORM_START_COL + 1).Name = "ProductName_Input"
        
        .Cells(4, FORM_START_COL).Value = "カテゴリ:"
        .Cells(4, FORM_START_COL + 1).Name = "Category_Input"
        
        .Cells(5, FORM_START_COL).Value = "価格:"
        .Cells(5, FORM_START_COL + 1).Name = "Price_Input"
        
        .Cells(6, FORM_START_COL).Value = "最小在庫:"
        .Cells(6, FORM_START_COL + 1).Name = "MinStock_Input"
        
        .Cells(7, FORM_START_COL).Value = "最大在庫:"
        .Cells(7, FORM_START_COL + 1).Name = "MaxStock_Input"
        
        .Cells(8, FORM_START_COL).Value = "現在在庫:"
        .Cells(8, FORM_START_COL + 1).Name = "CurrentStock_Input"
        
        .Cells(9, FORM_START_COL).Value = "仕入先:"
        .Cells(9, FORM_START_COL + 1).Name = "Supplier_Input"
        
        .Cells(10, FORM_START_COL).Value = "有効フラグ:"
        .Cells(10, FORM_START_COL + 1).Name = "IsActive_Input"
        .Cells(10, FORM_START_COL + 1).Value = True
        
        ' ボタンエリア
        .Cells(12, FORM_START_COL).Value = "追加/更新"
        .Cells(12, FORM_START_COL).Name = "AddUpdateButton"
        CreateButton "AddUpdateButton", "Sheet1.AddOrUpdateProduct_Click"
        
        .Cells(12, FORM_START_COL + 1).Value = "クリア"
        .Cells(12, FORM_START_COL + 1).Name = "ClearButton"
        CreateButton "ClearButton", "Sheet1.ClearForm_Click"
        
        .Cells(12, FORM_START_COL + 2).Value = "削除"
        .Cells(12, FORM_START_COL + 2).Name = "DeleteButton"
        CreateButton "DeleteButton", "Sheet1.DeleteProduct_Click"
        
        ' 検索エリア
        .Cells(14, FORM_START_COL).Value = "検索:"
        .Cells(14, FORM_START_COL + 1).Name = "Search_Input"
        .Cells(14, FORM_START_COL + 2).Value = "検索実行"
        .Cells(14, FORM_START_COL + 2).Name = "SearchButton"
        CreateButton "SearchButton", "Sheet1.SearchProducts_Click"
        
        ' カテゴリフィルタ
        .Cells(15, FORM_START_COL).Value = "カテゴリ:"
        .Cells(15, FORM_START_COL + 1).Name = "CategoryFilter_Input"
        
        ' 在庫状況フィルタ
        .Cells(16, FORM_START_COL).Value = "在庫状況:"
        .Cells(16, FORM_START_COL + 1).Name = "StockFilter_Input"
        
        ' レポート生成ボタン
        .Cells(18, FORM_START_COL).Value = "レポート生成"
        .Cells(18, FORM_START_COL).Name = "ReportButton"
        CreateButton "ReportButton", "Sheet1.GenerateReport_Click"
        
        ' フォーム範囲の書式設定
        With .Range(.Cells(1, FORM_START_COL), .Cells(20, FORM_START_COL + 3))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
        
        ' 入力欄の背景色
        With .Range(.Cells(2, FORM_START_COL + 1), .Cells(10, FORM_START_COL + 1))
            .Interior.Color = RGB(255, 255, 240)
        End With
    End With
    
    XLog "商品入力フォームの初期化が完了しました"
End Sub

' ボタンの作成
Private Sub CreateButton(cellName As String, macroName As String)
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

' 商品追加/更新ボタン
Public Sub AddOrUpdateProduct_Click()
    On Error GoTo ErrorHandler
    
    XLog "商品追加/更新ボタンがクリックされました"
    
    ' 入力値の取得
    Dim product As ProductInfo
    With product
        .ProductID = Trim(Range("ProductID_Input").Value)
        .ProductName = Trim(Range("ProductName_Input").Value)
        .Category = Trim(Range("Category_Input").Value)
        .Price = CDbl(Range("Price_Input").Value)
        .MinStock = CLng(Range("MinStock_Input").Value)
        .MaxStock = CLng(Range("MaxStock_Input").Value)
        .CurrentStock = CLng(Range("CurrentStock_Input").Value)
        .Supplier = Trim(Range("Supplier_Input").Value)
        .LastUpdated = Now
        .IsActive = CBool(Range("IsActive_Input").Value)
    End With
    
    ' データ検証
    Dim validationError As String
    validationError = ValidateProductData(product)
    If validationError <> "" Then
        MsgBox "入力エラー:" & vbCrLf & validationError, vbCritical, "商品マスタ"
        Exit Sub
    End If
    
    ' 商品の追加/更新
    If AddOrUpdateProduct(product) Then
        MsgBox "商品情報を正常に保存しました。", vbInformation, "商品マスタ"
        Call ClearForm_Click
        Call RefreshProductList
    Else
        MsgBox "商品情報の保存に失敗しました。", vbCritical, "商品マスタ"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "処理中にエラーが発生しました: " & Err.Description, vbCritical, "商品マスタエラー"
    XLog "商品追加/更新でエラー: " & Err.Description, "ERROR"
End Sub

' フォームクリアボタン
Public Sub ClearForm_Click()
    On Error Resume Next
    
    With Me
        .Range("ProductID_Input").Value = ""
        .Range("ProductName_Input").Value = ""
        .Range("Category_Input").Value = ""
        .Range("Price_Input").Value = ""
        .Range("MinStock_Input").Value = ""
        .Range("MaxStock_Input").Value = ""
        .Range("CurrentStock_Input").Value = ""
        .Range("Supplier_Input").Value = ""
        .Range("IsActive_Input").Value = True
    End With
    
    XLog "入力フォームをクリアしました"
End Sub

' 商品削除ボタン
Public Sub DeleteProduct_Click()
    On Error GoTo ErrorHandler
    
    Dim productID As String
    productID = Trim(Range("ProductID_Input").Value)
    
    If productID = "" Then
        MsgBox "削除する商品IDを入力してください。", vbExclamation, "商品マスタ"
        Exit Sub
    End If
    
    If MsgBox("商品ID「" & productID & "」を削除しますか？", vbYesNo + vbQuestion, "商品削除確認") = vbYes Then
        Dim targetRow As Long
        targetRow = FindProductRow(productID)
        
        If targetRow > 0 Then
            ' 論理削除（有効フラグをFalseに設定）
            Cells(targetRow, 10).Value = False
            Cells(targetRow, 9).Value = Now
            
            MsgBox "商品を削除しました。", vbInformation, "商品マスタ"
            Call ClearForm_Click
            Call RefreshProductList
        Else
            MsgBox "指定された商品IDが見つかりません。", vbExclamation, "商品マスタ"
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "削除処理中にエラーが発生しました: " & Err.Description, vbCritical, "商品マスタエラー"
End Sub

' 商品検索ボタン
Public Sub SearchProducts_Click()
    On Error GoTo ErrorHandler
    
    Dim searchTerm As String
    Dim categoryFilter As String
    Dim stockFilter As String
    
    searchTerm = Trim(Range("Search_Input").Value)
    categoryFilter = Trim(Range("CategoryFilter_Input").Value)
    stockFilter = Trim(Range("StockFilter_Input").Value)
    
    XLog "商品検索を実行: " & searchTerm
    
    Dim results As Collection
    Set results = SearchProducts(searchTerm, categoryFilter, stockFilter)
    
    ' 検索結果の表示
    Call DisplaySearchResults(results)
    
    MsgBox "検索が完了しました。結果: " & results.Count & "件", vbInformation, "商品検索"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "検索処理中にエラーが発生しました: " & Err.Description, vbCritical, "検索エラー"
End Sub

' レポート生成ボタン
Public Sub GenerateReport_Click()
    On Error Resume Next
    
    XLog "レポート生成ボタンがクリックされました"
    Call GenerateInventoryStatusReport
End Sub

'===================================================================
' データ表示機能
'===================================================================

' 商品リストの更新
Private Sub RefreshProductList()
    On Error Resume Next
    
    XLog "商品リストを更新中"
    
    ' データ領域の書式設定
    With Range("A1:J1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
        .Borders.LineStyle = xlContinuous
    End With
    
    ' 列幅の調整
    Columns("A:J").AutoFit
    
    XLog "商品リストの更新が完了しました"
End Sub

' 商品編集
Private Sub EditProduct(targetRow As Long)
    On Error Resume Next
    
    XLog "商品編集: 行" & targetRow
    
    ' 選択された商品の情報を入力フォームに設定
    With Me
        .Range("ProductID_Input").Value = .Cells(targetRow, 1).Value
        .Range("ProductName_Input").Value = .Cells(targetRow, 2).Value
        .Range("Category_Input").Value = .Cells(targetRow, 3).Value
        .Range("Price_Input").Value = .Cells(targetRow, 4).Value
        .Range("MinStock_Input").Value = .Cells(targetRow, 5).Value
        .Range("MaxStock_Input").Value = .Cells(targetRow, 6).Value
        .Range("CurrentStock_Input").Value = .Cells(targetRow, 7).Value
        .Range("Supplier_Input").Value = .Cells(targetRow, 8).Value
        .Range("IsActive_Input").Value = .Cells(targetRow, 10).Value
    End With
End Sub

' 検索結果の表示
Private Sub DisplaySearchResults(results As Collection)
    On Error Resume Next
    
    ' データエリアをクリア（ヘッダー行は残す）
    If Cells(2, 1).Value <> "" Then
        Range("A2:J" & Cells(Rows.Count, 1).End(xlUp).Row).Clear
    End If
    
    ' 検索結果を表示
    Dim i As Long
    i = 2
    
    Dim product As ProductInfo
    For Each product In results
        With Me
            .Cells(i, 1).Value = product.ProductID
            .Cells(i, 2).Value = product.ProductName
            .Cells(i, 3).Value = product.Category
            .Cells(i, 4).Value = product.Price
            .Cells(i, 5).Value = product.MinStock
            .Cells(i, 6).Value = product.MaxStock
            .Cells(i, 7).Value = product.CurrentStock
            .Cells(i, 8).Value = product.Supplier
            .Cells(i, 9).Value = product.LastUpdated
            .Cells(i, 10).Value = product.IsActive
            
            ' 在庫状況に応じた行の色分け
            If product.CurrentStock <= product.MinStock Then
                .Range(.Cells(i, 1), .Cells(i, 10)).Interior.Color = RGB(255, 220, 220)
            ElseIf product.CurrentStock >= product.MaxStock Then
                .Range(.Cells(i, 1), .Cells(i, 10)).Interior.Color = RGB(255, 255, 220)
            End If
        End With
        
        i = i + 1
    Next product
    
    Call RefreshProductList
End Sub

'===================================================================
' サンプルデータ生成
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
    
    ' 残りのサンプルデータ
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
    
    Call RefreshProductList
    MsgBox "サンプルデータ（10件）を作成しました。", vbInformation, "商品マスタ"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "サンプルデータ作成中にエラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub
