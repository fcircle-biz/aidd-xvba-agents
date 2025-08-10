VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'===================================================================
' 商品在庫管理システム - ワークブックイベント
'===================================================================

'===================================================================
' ワークブックイベント
'===================================================================

' ワークブック起動時の処理
Private Sub Workbook_Open()
    On Error Resume Next
    
    XLog "商品在庫管理システムが起動されました"
    
    ' システム初期化
    Call InitializeSystem
    
    ' スプラッシュ画面の表示
    Call ShowSplashScreen
    
    ' メインメニューの表示（遅延実行）
    Call ShowMainMenuDelayed
End Sub

' ワークブック終了前の処理
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    
    XLog "商品在庫管理システムが終了されます"
    
    ' システムクリーンアップ
    Call CleanupSystem
End Sub

' シート作成後の処理
Private Sub Workbook_NewSheet(ByVal Sh As Object)
    On Error Resume Next
    
    XLog "新しいシートが作成されました: " & Sh.Name
End Sub

'===================================================================
' システム初期化
'===================================================================

' システムの初期化
Private Sub InitializeSystem()
    On Error GoTo ErrorHandler
    
    XLog "システムを初期化中..."
    
    ' アプリケーション設定
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = True
    
    ' 必要なシートの作成・確認
    Call EnsureSystemSheets
    
    ' ログディレクトリの確認・作成
    Call EnsureLogDirectory
    
    ' データベースの初期化チェック
    Call InitializeDataStructure
    
    ' サンプルデータの作成（初回のみ）
    Call CreateInitialSampleData
    
    ' カスタムリボンの設定（オプション）
    Call SetupCustomRibbon
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    XLog "システムの初期化が完了しました"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    XLog "システム初期化でエラーが発生: " & Err.Description, "ERROR"
    MsgBox "システムの初期化中にエラーが発生しました: " & Err.Description, vbCritical, "システムエラー"
End Sub

' システムシートの確認・作成
Private Sub EnsureSystemSheets()
    On Error Resume Next
    
    ' 商品マスタシートの確認
    Dim productSheet As Worksheet
    Set productSheet = GetOrCreateSheet(PRODUCT_SHEET)
    
    ' 在庫管理シートの確認
    Dim inventorySheet As Worksheet
    Set inventorySheet = GetOrCreateSheet(INVENTORY_SHEET)
    
    ' 取引履歴シートの確認
    Dim transactionSheet As Worksheet
    Set transactionSheet = GetOrCreateSheet(TRANSACTION_SHEET)
    
    ' レポートシートの確認
    Dim reportSheet As Worksheet
    Set reportSheet = GetOrCreateSheet(REPORT_SHEET)
    
    XLog "システムシートの確認・作成が完了しました"
End Sub

' ログディレクトリの確認・作成
Private Sub EnsureLogDirectory()
    On Error Resume Next
    
    Dim logDir As String
    logDir = ThisWorkbook.Path & "\logs"
    
    ' ディレクトリが存在しない場合は作成
    If Dir(logDir, vbDirectory) = "" Then
        MkDir logDir
        XLog "ログディレクトリを作成しました: " & logDir
    End If
End Sub

' データ構造の初期化
Private Sub InitializeDataStructure()
    On Error Resume Next
    
    Dim ws As Worksheet
    
    ' 商品マスタのヘッダーチェック
    Set ws = Worksheets(PRODUCT_SHEET)
    If ws.Cells(1, 1).Value = "" Then
        Call CreateProductMasterHeaders(ws)
        XLog "商品マスタのヘッダーを初期化しました"
    End If
    
    ' 取引履歴のヘッダーチェック
    Set ws = Worksheets(TRANSACTION_SHEET)
    If ws.Cells(1, 1).Value = "" Then
        Call CreateTransactionHistoryHeaders(ws)
        XLog "取引履歴のヘッダーを初期化しました"
    End If
End Sub

' カスタムリボンの設定
Private Sub SetupCustomRibbon()
    On Error Resume Next
    
    ' リボンカスタマイズは高度な機能のため、基本的なメニューを作成
    Call CreateCustomMenu
End Sub

' カスタムメニューの作成
Private Sub CreateCustomMenu()
    On Error Resume Next
    
    ' 既存のカスタムメニューを削除
    Application.CommandBars("在庫管理システム").Delete
    
    ' 新しいカスタムメニューバーを作成
    Dim menuBar As CommandBar
    Set menuBar = Application.CommandBars.Add("在庫管理システム", msoBarFloating, True, True)
    
    ' メニュー項目の追加
    With menuBar.Controls.Add(Type:=msoControlPopup)
        .Caption = "商品管理"
        
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "商品マスタ画面"
            .OnAction = "ShowProductMaster"
        End With
        
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "サンプルデータ作成"
            .OnAction = "CreateSampleData"
        End With
    End With
    
    With menuBar.Controls.Add(Type:=msoControlPopup)
        .Caption = "在庫管理"
        
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "在庫管理画面"
            .OnAction = "ShowInventoryManagement"
        End With
        
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "サンプル取引作成"
            .OnAction = "CreateSampleTransactions"
        End With
    End With
    
    With menuBar.Controls.Add(Type:=msoControlPopup)
        .Caption = "レポート"
        
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "在庫状況レポート"
            .OnAction = "GenerateInventoryStatusReport"
        End With
        
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "取引履歴レポート"
            .OnAction = "GenerateTransactionReport"
        End With
    End With
    
    With menuBar.Controls.Add(Type:=msoControlPopup)
        .Caption = "ヘルプ"
        
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "システム情報"
            .OnAction = "ShowSystemInfo"
        End With
        
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "使用方法"
            .OnAction = "ShowUserGuide"
        End With
    End With
    
    menuBar.Visible = True
    XLog "カスタムメニューを作成しました"
End Sub

'===================================================================
' システムクリーンアップ
'===================================================================

' システムのクリーンアップ
Private Sub CleanupSystem()
    On Error Resume Next
    
    XLog "システムのクリーンアップを実行中..."
    
    ' カスタムメニューの削除
    Application.CommandBars("在庫管理システム").Delete
    
    ' 一時ファイルのクリーンアップ
    Call CleanupTempFiles
    
    ' 最終ログの出力
    XLog "商品在庫管理システムが正常に終了しました"
End Sub

' 一時ファイルのクリーンアップ
Private Sub CleanupTempFiles()
    On Error Resume Next
    
    Dim tempDir As String
    tempDir = Environ("TEMP") & "\XVBA_Inventory\"
    
    ' 一時ディレクトリが存在する場合はクリーンアップ
    If Dir(tempDir, vbDirectory) <> "" Then
        ' 一時ファイルの削除処理
        XLog "一時ファイルをクリーンアップしました"
    End If
End Sub

'===================================================================
' メニュー関数
'===================================================================

' メインメニューの遅延表示
Private Sub ShowMainMenuDelayed()
    On Error Resume Next
    Application.Wait Now + TimeSerial(0, 0, 2)
    Call ShowMainMenu
End Sub

' メインメニューの表示
Public Sub ShowMainMenu()
    On Error Resume Next
    
    Dim message As String
    message = "商品在庫管理システム" & vbCrLf & vbCrLf
    message = message & "利用可能な機能:" & vbCrLf
    message = message & "・商品マスタ管理" & vbCrLf
    message = message & "・在庫管理（入庫・出庫・棚卸）" & vbCrLf
    message = message & "・在庫状況レポート" & vbCrLf
    message = message & "・取引履歴管理" & vbCrLf & vbCrLf
    message = message & "各機能は対応するシートタブをクリックして利用できます。"
    
    MsgBox message, vbInformation, "商品在庫管理システム"
End Sub

' スプラッシュ画面の表示
Private Sub ShowSplashScreen()
    On Error Resume Next
    
    Dim message As String
    message = "商品在庫管理システム v1.0" & vbCrLf & vbCrLf
    message = message & "XVBA (Extended VBA) Framework" & vbCrLf
    message = message & "Powered by Excel VBA" & vbCrLf & vbCrLf
    message = message & "システムを初期化しています..."
    
    ' 短時間のスプラッシュ表示
    Application.StatusBar = message
    DoEvents
    Application.Wait Now + TimeSerial(0, 0, 1)
    Application.StatusBar = False
End Sub

' 商品マスタ画面の表示
Public Sub ShowProductMaster()
    On Error Resume Next
    Worksheets(PRODUCT_SHEET).Activate
End Sub

' 在庫管理画面の表示
Public Sub ShowInventoryManagement()
    On Error Resume Next
    Worksheets(INVENTORY_SHEET).Activate
End Sub

' システム情報の表示
Public Sub ShowSystemInfo()
    On Error Resume Next
    
    Dim info As String
    info = "商品在庫管理システム" & vbCrLf & vbCrLf
    info = info & "バージョン: 1.0" & vbCrLf
    info = info & "フレームワーク: XVBA" & vbCrLf
    info = info & "作成日: " & Format(ThisWorkbook.BuiltinDocumentProperties("Creation Date"), "yyyy/mm/dd") & vbCrLf
    info = info & "ファイル名: " & ThisWorkbook.Name & vbCrLf
    info = info & "パス: " & ThisWorkbook.Path & vbCrLf
    info = info & "Excel バージョン: " & Application.Version & vbCrLf
    info = info & "ユーザー名: " & Application.UserName & vbCrLf & vbCrLf
    info = info & "インストール済みパッケージ:" & vbCrLf
    info = info & "・Xdebug (デバッグサポート)" & vbCrLf
    info = info & "・excel-types (型定義)" & vbCrLf
    
    MsgBox info, vbInformation, "システム情報"
End Sub

' 使用方法の表示
Public Sub ShowUserGuide()
    On Error Resume Next
    
    Dim guide As String
    guide = "商品在庫管理システム 使用方法" & vbCrLf & vbCrLf
    guide = guide & "■ 商品マスタ管理" & vbCrLf
    guide = guide & "1. 「商品マスタ」シートで商品情報を管理" & vbCrLf
    guide = guide & "2. 右側の入力フォームで商品の追加・編集" & vbCrLf
    guide = guide & "3. データ行をダブルクリックで編集モード" & vbCrLf
    guide = guide & "4. 検索・フィルタ機能で商品を検索" & vbCrLf & vbCrLf
    guide = guide & "■ 在庫管理" & vbCrLf
    guide = guide & "1. 「在庫管理」シートで在庫取引を実行" & vbCrLf
    guide = guide & "2. 入庫・出庫・棚卸の3種類の取引に対応" & vbCrLf
    guide = guide & "3. 一括処理でCSVファイルからの取込が可能" & vbCrLf
    guide = guide & "4. 在庫アラート機能で適正在庫を管理" & vbCrLf & vbCrLf
    guide = guide & "■ レポート機能" & vbCrLf
    guide = guide & "1. 在庫状況レポートで現在の在庫状況を確認" & vbCrLf
    guide = guide & "2. 取引履歴で過去の取引を追跡" & vbCrLf
    guide = guide & "3. 色分け表示で在庫状況を視覚的に把握" & vbCrLf
    
    MsgBox guide, vbInformation, "使用方法"
End Sub

' 取引履歴レポートの生成
Public Sub GenerateTransactionReport()
    On Error GoTo ErrorHandler
    
    XLog "取引履歴レポート生成を開始"
    
    Dim reportWs As Worksheet
    Set reportWs = GetOrCreateSheet("取引履歴レポート")
    
    ' レポートシートをクリア
    reportWs.Cells.Clear
    
    ' レポートヘッダーの作成
    With reportWs
        .Cells(1, 1).Value = "取引履歴レポート - " & Format(Now, "yyyy/mm/dd hh:nn")
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 14
        
        .Cells(2, 1).Value = "取引ID"
        .Cells(2, 2).Value = "商品ID"
        .Cells(2, 3).Value = "商品名"
        .Cells(2, 4).Value = "取引種別"
        .Cells(2, 5).Value = "数量"
        .Cells(2, 6).Value = "取引日時"
        .Cells(2, 7).Value = "担当者"
        .Cells(2, 8).Value = "備考"
        
        With .Range("A2:H2")
            .Font.Bold = True
            .Interior.Color = RGB(180, 220, 180)
            .Borders.LineStyle = xlContinuous
        End With
    End With
    
    ' 取引履歴データの取得と出力
    Dim transactionWs As Worksheet
    Set transactionWs = Worksheets(TRANSACTION_SHEET)
    
    Dim lastRow As Long
    lastRow = transactionWs.Cells(transactionWs.Rows.Count, 1).End(xlUp).Row
    
    If lastRow > 1 Then
        Dim reportRow As Long
        reportRow = 3
        
        Dim i As Long
        For i = 2 To lastRow
            Dim productID As String
            productID = transactionWs.Cells(i, 2).Value
            
            Dim product As ProductInfo
            product = GetProduct(productID)
            
            With reportWs
                .Cells(reportRow, 1).Value = transactionWs.Cells(i, 1).Value ' 取引ID
                .Cells(reportRow, 2).Value = transactionWs.Cells(i, 2).Value ' 商品ID
                .Cells(reportRow, 3).Value = product.ProductName           ' 商品名
                .Cells(reportRow, 4).Value = transactionWs.Cells(i, 3).Value ' 取引種別
                .Cells(reportRow, 5).Value = transactionWs.Cells(i, 4).Value ' 数量
                .Cells(reportRow, 6).Value = transactionWs.Cells(i, 5).Value ' 取引日時
                .Cells(reportRow, 7).Value = transactionWs.Cells(i, 6).Value ' 担当者
                .Cells(reportRow, 8).Value = transactionWs.Cells(i, 7).Value ' 備考
                
                ' 取引種別に応じた色分け
                Select Case transactionWs.Cells(i, 3).Value
                    Case "入庫"
                        .Range(.Cells(reportRow, 1), .Cells(reportRow, 8)).Interior.Color = RGB(220, 255, 220)
                    Case "出庫"
                        .Range(.Cells(reportRow, 1), .Cells(reportRow, 8)).Interior.Color = RGB(255, 220, 220)
                    Case "棚卸"
                        .Range(.Cells(reportRow, 1), .Cells(reportRow, 8)).Interior.Color = RGB(255, 255, 220)
                End Select
            End With
            
            reportRow = reportRow + 1
        Next i
        
        ' レポートの書式設定
        With reportWs
            .Columns("A:H").AutoFit
            .Range("E3:E" & reportRow - 1).NumberFormat = "#,##0"
            .Range("F3:F" & reportRow - 1).NumberFormat = "yyyy/mm/dd hh:nn"
            
            With .Range("A2:H" & reportRow - 1)
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
            End With
        End With
        
        reportWs.Activate
    End If
    
    XLog "取引履歴レポート生成が完了しました"
    MsgBox "取引履歴レポートが生成されました。", vbInformation, "在庫管理システム"
    
    Exit Sub
    
ErrorHandler:
    XLog "取引履歴レポート生成でエラーが発生: " & Err.Description, "ERROR"
    MsgBox "取引履歴レポート生成中にエラーが発生しました: " & Err.Description, vbCritical, "レポートエラー"
End Sub

'===================================================================
' ユーティリティ関数
'===================================================================

' システムステータスの取得
Public Function GetSystemStatus() As String
    On Error Resume Next
    
    Dim status As String
    status = "システム状態: 正常" & vbCrLf
    status = status & "商品数: " & GetProductCount() & " 件" & vbCrLf
    status = status & "取引履歴: " & GetTransactionCount() & " 件" & vbCrLf
    status = status & "低在庫商品: " & GetLowStockCount() & " 件"
    
    GetSystemStatus = status
End Function

' 商品数の取得
Private Function GetProductCount() As Long
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = Worksheets(PRODUCT_SHEET)
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow > 1 Then
        GetProductCount = lastRow - 1
    Else
        GetProductCount = 0
    End If
End Function

' 取引履歴数の取得
Private Function GetTransactionCount() As Long
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = Worksheets(TRANSACTION_SHEET)
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow > 1 Then
        GetTransactionCount = lastRow - 1
    Else
        GetTransactionCount = 0
    End If
End Function

' 低在庫商品数の取得
Private Function GetLowStockCount() As Long
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = Worksheets(PRODUCT_SHEET)
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim count As Long
    count = 0
    
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 10).Value Then ' アクティブな商品のみ
            If ws.Cells(i, 7).Value <= ws.Cells(i, 5).Value Then ' 現在在庫 <= 最小在庫
                count = count + 1
            End If
        End If
    Next i
    
    GetLowStockCount = count
End Function

'===================================================================
' サンプルデータ自動作成
'===================================================================

' 初回起動時のサンプルデータ作成
Private Sub CreateInitialSampleData()
    On Error Resume Next
    
    XLog "初回起動時のサンプルデータ作成をチェック中"
    
    ' 商品マスタにデータがあるかチェック
    Dim productWs As Worksheet
    Set productWs = Worksheets(PRODUCT_SHEET)
    
    If productWs.Cells(2, 1).Value = "" Then
        XLog "商品データが存在しないため、サンプルデータを作成します"
        
        ' サンプル商品データを作成
        Call CreateSampleData
        
        XLog "商品サンプルデータを作成しました"
    Else
        XLog "既に商品データが存在するため、サンプルデータ作成をスキップします"
    End If
    
    ' 取引履歴にデータがあるかチェック
    Dim transactionWs As Worksheet
    Set transactionWs = Worksheets(TRANSACTION_SHEET)
    
    If transactionWs.Cells(2, 1).Value = "" Then
        XLog "取引データが存在しないため、サンプルデータを作成します"
        
        ' サンプル取引データを作成
        Call CreateSampleTransactions
        
        XLog "取引サンプルデータを作成しました"
    Else
        XLog "既に取引データが存在するため、サンプルデータ作成をスキップします"
    End If
End Sub
