Attribute VB_Name = "modUI"
Option Explicit

' =============================================================================
' 備品貸出管理システム - UI操作関数（フォーマット、ボタン設定等）
' =============================================================================

' 標準テーブルフォーマット適用
Public Sub ApplyStandardTableFormat(tbl As ListObject)
    On Error GoTo ErrHandler
    
    If tbl Is Nothing Then Exit Sub
    
    ' ヘッダー行のフォーマット
    With tbl.HeaderRowRange
        .Font.Bold = True
        .Interior.Color = COLOR_HEADER
        .Font.Color = vbWhite
        .Font.Size = 11
        .RowHeight = 25
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' データ行の交互色設定
    If Not tbl.DataBodyRange Is Nothing Then
        Call ApplyAlternatingRowColors(tbl.DataBodyRange)
    End If
    
    ' 枠線の設定
    With tbl.Range.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(128, 128, 128)
    End With
    
    ' 列幅の自動調整
    tbl.Range.Columns.AutoFit
    
    Exit Sub
    
ErrHandler:
    Call LogError("ApplyStandardTableFormat", Err.Number, Err.Description)
End Sub

' 交互行色設定関数
Private Sub ApplyAlternatingRowColors(dataRange As Range)
    On Error GoTo ErrHandler
    
    Dim i As Long
    For i = 1 To dataRange.Rows.Count
        If i Mod 2 = 0 Then
            dataRange.Rows(i).Interior.Color = COLOR_ALTERNATE
        Else
            dataRange.Rows(i).Interior.Color = COLOR_NORMAL
        End If
    Next i
    
    Exit Sub
    
ErrHandler:
    Call LogError("ApplyAlternatingRowColors", Err.Number, Err.Description)
End Sub

' 条件付き色分け関数
Public Sub ApplyConditionalFormatting(targetRange As Range, status As String)
    On Error GoTo ErrHandler
    
    Select Case status
        Case "期限超過", "エラー", "失敗"
            With targetRange
                .Interior.Color = COLOR_OVERDUE
                .Font.Color = vbWhite
                .Font.Bold = True
            End With
        Case "期限間近", "警告", "注意"
            With targetRange
                .Interior.Color = COLOR_WARNING
                .Font.Color = vbBlack
                .Font.Bold = True
            End With
        Case "完了", "成功", "返却済"
            With targetRange
                .Interior.Color = COLOR_SUCCESS
                .Font.Color = vbWhite
            End With
        Case Else
            With targetRange
                .Interior.Color = COLOR_NORMAL
                .Font.Color = vbBlack
            End With
    End Select
    
    Exit Sub
    
ErrHandler:
    Call LogError("ApplyConditionalFormatting", Err.Number, Err.Description)
End Sub

' ダッシュボードレイアウト作成
Public Sub CreateDashboardLayout()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = GetWorksheet(SHEET_DASHBOARD)
    If ws Is Nothing Then
        Call LogError("CreateDashboardLayout", 9, "Dashboard sheet not found")
        Exit Sub
    End If
    
    ' 既存の内容をクリア
    ws.Cells.Clear
    
    ' タイトル部分
    With ws.Range("A1:L1")
        .Merge
        .Value = "備品貸出管理システム - ダッシュボード"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 35
    End With
    
    ' KPIサマリーセクション
    Call CreateKPISummarySection(ws)
    
    ' ボタン配置
    Call CreateActionButtons(ws)
    
    ' データ表示エリアのラベル
    ws.Range("A7").Value = "■ 貸出中一覧"
    ws.Range("A7").Font.Bold = True
    ws.Range("A7").Font.Size = 12
    
    ws.Range("H7").Value = "■ 在庫状況"
    ws.Range("H7").Font.Bold = True
    ws.Range("H7").Font.Size = 12
    
    ws.Range("A21").Value = "■ 期限超過一覧"
    ws.Range("A21").Font.Bold = True
    ws.Range("A21").Font.Size = 12
    ws.Range("A21").Font.Color = COLOR_OVERDUE
    
    ' 初期データ更新
    Call UpdateDashboard
    
    Exit Sub
    
ErrHandler:
    Call LogError("CreateDashboardLayout", Err.Number, Err.Description)
End Sub

' KPIサマリーセクション作成（内部関数）
Private Sub CreateKPISummarySection(ws As Worksheet)
    On Error Resume Next
    
    ' KPIラベル
    ws.Range("A3").Value = "総備品数:"
    ws.Range("A3").Font.Bold = True
    
    ws.Range("C3").Value = "貸出中:"
    ws.Range("C3").Font.Bold = True
    
    ws.Range("E3").Value = "期限超過:"
    ws.Range("E3").Font.Bold = True
    
    ws.Range("G3").Value = "利用可能:"
    ws.Range("G3").Font.Bold = True
    
    ' KPI値セル（計算結果が入る）
    With ws.Range("B3")
        .Interior.Color = COLOR_SUCCESS
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    With ws.Range("D3")
        .Interior.Color = COLOR_WARNING
        .Font.Color = vbBlack
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    With ws.Range("F3")
        .Interior.Color = COLOR_OVERDUE
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    With ws.Range("H3")
        .Interior.Color = COLOR_SUCCESS
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    On Error GoTo 0
End Sub

' アクションボタン作成（内部関数）
Private Sub CreateActionButtons(ws As Worksheet)
    On Error GoTo ErrHandler
    
    ' ダッシュボード更新ボタン
    Dim btnUpdate As Button
    Set btnUpdate = ws.Buttons.Add(ws.Range("J3").Left, ws.Range("J3").Top, 80, 25)
    btnUpdate.Caption = "更新"
    btnUpdate.OnAction = "modDashboard.UpdateDashboard"
    
    ' 貸出登録ボタン
    Dim btnLend As Button
    Set btnLend = ws.Buttons.Add(ws.Range("A5").Left, ws.Range("A5").Top, 100, 25)
    btnLend.Caption = "貸出登録"
    btnLend.OnAction = "modLending.RegisterLending"
    
    ' 返却登録ボタン
    Dim btnReturn As Button
    Set btnReturn = ws.Buttons.Add(ws.Range("C5").Left, ws.Range("C5").Top, 100, 25)
    btnReturn.Caption = "返却登録"
    btnReturn.OnAction = "modLending.RegisterReturn"
    
    ' 入力画面表示ボタン
    Dim btnInput As Button
    Set btnInput = ws.Buttons.Add(ws.Range("E5").Left, ws.Range("E5").Top, 100, 25)
    btnInput.Caption = "入力画面"
    btnInput.OnAction = "modUI.ShowInputSheet"
    
    ' テストデータ作成ボタン
    Dim btnTest As Button
    Set btnTest = ws.Buttons.Add(ws.Range("G5").Left, ws.Range("G5").Top, 120, 25)
    btnTest.Caption = "テストデータ作成"
    btnTest.OnAction = "modTestData.CreateAllTestData"
    
    Exit Sub
    
ErrHandler:
    Call LogError("CreateActionButtons", Err.Number, Err.Description)
End Sub

' 入力シートレイアウト作成
Public Sub CreateInputLayout()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = GetWorksheet(SHEET_INPUT)
    If ws Is Nothing Then
        Call LogError("CreateInputLayout", 9, "Input sheet not found")
        Exit Sub
    End If
    
    ' 既存の内容をクリア
    ws.Cells.Clear
    
    ' タイトル
    With ws.Range("A1:E1")
        .Merge
        .Value = "備品貸出・返却入力フォーム"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 30
    End With
    
    ' 入力項目ラベル
    ws.Range("A3").Value = "備品ID:"
    ws.Range("A4").Value = "借用者:"
    ws.Range("A5").Value = "貸出日:"
    ws.Range("A6").Value = "貸出期間（日）:"
    ws.Range("A7").Value = "返却日:"
    
    ' ラベルのフォーマット
    With ws.Range("A3:A7")
        .Font.Bold = True
        .VerticalAlignment = xlCenter
    End With
    
    ' 入力セルのフォーマット
    With ws.Range("B3:B7")
        .Interior.Color = RGB(255, 255, 204) ' 薄い黄色
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
    End With
    
    ' 説明テキスト
    ws.Range("D3").Value = "例: 1001"
    ws.Range("D4").Value = "例: 田中太郎"
    ws.Range("D5").Value = "例: 2024/1/15 (空白=今日)"
    ws.Range("D6").Value = "例: 7 (空白=7日)"
    ws.Range("D7").Value = "例: 2024/1/22 (返却時のみ)"
    
    With ws.Range("D3:D7")
        .Font.Color = RGB(128, 128, 128)
        .Font.Italic = True
    End With
    
    ' 貸出・返却説明
    ws.Range("A9").Value = "■ 貸出登録手順:"
    ws.Range("A10").Value = "1. 備品ID、借用者、貸出日、貸出期間を入力"
    ws.Range("A11").Value = "2. ダッシュボードの「貸出登録」ボタンをクリック"
    
    ws.Range("A13").Value = "■ 返却登録手順:"
    ws.Range("A14").Value = "1. 備品ID、借用者、返却日を入力"
    ws.Range("A15").Value = "2. ダッシュボードの「返却登録」ボタンをクリック"
    
    With ws.Range("A9,A13")
        .Font.Bold = True
        .Font.Color = RGB(68, 114, 196)
    End With
    
    ' ボタン作成
    Call CreateInputButtons(ws)
    
    Exit Sub
    
ErrHandler:
    Call LogError("CreateInputLayout", Err.Number, Err.Description)
End Sub

' 入力シートボタン作成（内部関数）
Private Sub CreateInputButtons(ws As Worksheet)
    On Error GoTo ErrHandler
    
    ' ダッシュボードに戻るボタン
    Dim btnDashboard As Button
    Set btnDashboard = ws.Buttons.Add(ws.Range("A17").Left, ws.Range("A17").Top, 120, 25)
    btnDashboard.Caption = "ダッシュボードへ"
    btnDashboard.OnAction = "modUI.ShowDashboard"
    
    ' 入力クリアボタン
    Dim btnClear As Button
    Set btnClear = ws.Buttons.Add(ws.Range("C17").Left, ws.Range("C17").Top, 100, 25)
    btnClear.Caption = "入力クリア"
    btnClear.OnAction = "modUI.ClearInputForm"
    
    Exit Sub
    
ErrHandler:
    Call LogError("CreateInputButtons", Err.Number, Err.Description)
End Sub

' 備品マスタシートレイアウト作成
Public Sub CreateItemsLayout()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = GetWorksheet(SHEET_ITEMS)
    If ws Is Nothing Then
        Call LogError("CreateItemsLayout", 9, "Items sheet not found")
        Exit Sub
    End If
    
    ' 既存の内容をクリア
    ws.Cells.Clear
    
    ' タイトル
    With ws.Range("A1:E1")
        .Merge
        .Value = "備品マスタ"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 30
    End With
    
    ' テーブル作成
    Call CreateItemsTable(ws)
    
    Exit Sub
    
ErrHandler:
    Call LogError("CreateItemsLayout", Err.Number, Err.Description)
End Sub

' 備品テーブル作成（内部関数）
Private Sub CreateItemsTable(ws As Worksheet)
    On Error GoTo ErrHandler
    
    ' ヘッダー作成
    ws.Range("A3").Value = COL_ITEM_ID
    ws.Range("B3").Value = COL_ITEM_NAME
    ws.Range("C3").Value = COL_CATEGORY
    ws.Range("D3").Value = COL_LOCATION
    ws.Range("E3").Value = COL_QUANTITY
    
    ' テーブル化
    Dim rng As Range
    Set rng = ws.Range("A3:E3")
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.Name = TABLE_ITEMS
    
    ' フォーマット適用
    Call ApplyStandardTableFormat(tbl)
    
    Exit Sub
    
ErrHandler:
    Call LogError("CreateItemsTable", Err.Number, Err.Description)
End Sub

' 貸出履歴シートレイアウト作成
Public Sub CreateLendingLayout()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = GetWorksheet(SHEET_LENDING)
    If ws Is Nothing Then
        Call LogError("CreateLendingLayout", 9, "Lending sheet not found")
        Exit Sub
    End If
    
    ' 既存の内容をクリア
    ws.Cells.Clear
    
    ' タイトル
    With ws.Range("A1:I1")
        .Merge
        .Value = "貸出・返却履歴"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 30
    End With
    
    ' テーブル作成
    Call CreateLendingTable(ws)
    
    Exit Sub
    
ErrHandler:
    Call LogError("CreateLendingLayout", Err.Number, Err.Description)
End Sub

' 貸出テーブル作成（内部関数）
Private Sub CreateLendingTable(ws As Worksheet)
    On Error GoTo ErrHandler
    
    ' ヘッダー作成
    ws.Range("A3").Value = COL_RECORD_ID
    ws.Range("B3").Value = COL_LENDING_ITEM_ID
    ws.Range("C3").Value = COL_LENDING_ITEM_NAME
    ws.Range("D3").Value = COL_BORROWER
    ws.Range("E3").Value = COL_LEND_DATE
    ws.Range("F3").Value = COL_DUE_DATE
    ws.Range("G3").Value = COL_RETURN_DATE
    ws.Range("H3").Value = COL_STATUS
    ws.Range("I3").Value = COL_REMARKS
    
    ' テーブル化
    Dim rng As Range
    Set rng = ws.Range("A3:I3")
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.Name = TABLE_LENDING
    
    ' フォーマット適用
    Call ApplyStandardTableFormat(tbl)
    
    Exit Sub
    
ErrHandler:
    Call LogError("CreateLendingTable", Err.Number, Err.Description)
End Sub

' シート表示切り替え関数群
Public Sub ShowDashboard()
    On Error Resume Next
    ThisWorkbook.Worksheets(SHEET_DASHBOARD).Activate
    On Error GoTo 0
End Sub

Public Sub ShowInputSheet()
    On Error Resume Next
    ThisWorkbook.Worksheets(SHEET_INPUT).Activate
    On Error GoTo 0
End Sub

Public Sub ShowItemsSheet()
    On Error Resume Next
    ThisWorkbook.Worksheets(SHEET_ITEMS).Activate
    On Error GoTo 0
End Sub

Public Sub ShowLendingSheet()
    On Error Resume Next
    ThisWorkbook.Worksheets(SHEET_LENDING).Activate
    On Error GoTo 0
End Sub

' 入力フォームクリア
Public Sub ClearInputForm()
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = GetWorksheet(SHEET_INPUT)
    If Not ws Is Nothing Then
        ws.Range(INPUT_ITEM_ID).Value = ""
        ws.Range(INPUT_BORROWER).Value = ""
        ws.Range(INPUT_LEND_DATE).Value = ""
        ws.Range(INPUT_LENDING_DAYS).Value = ""
        ws.Range(INPUT_RETURN_DATE).Value = ""
        MsgBox "入力フォームをクリアしました。", vbInformation
    End If
    
    On Error GoTo 0
End Sub

' 全シートレイアウト初期化
Public Sub InitializeAllLayouts()
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    
    Call CreateDashboardLayout
    Call CreateInputLayout
    Call CreateItemsLayout
    Call CreateLendingLayout
    
    ' ダッシュボードを表示
    Call ShowDashboard
    
    Application.ScreenUpdating = True
    MsgBox "全シートのレイアウトを初期化しました。", vbInformation
    
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Call LogError("InitializeAllLayouts", Err.Number, Err.Description)
    MsgBox "レイアウト初期化中にエラーが発生しました: " & Err.Description, vbCritical
End Sub