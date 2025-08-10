VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' ===============================================
' 従業員管理システム - ThisWorkbookクラス
' ===============================================

Option Explicit

' ===============================================
' ワークブックイベント
' ===============================================

Private Sub Workbook_Open()
    ' ワークブックが開かれた時の初期化処理
    On Error GoTo ErrorHandler
    
    Call Xdebug.printx("従業員管理システムを開始します...")
    
    ' システムの初期化
    Call InitializeEmployeeManagementSystem
    
    ' ウェルカムメッセージの表示
    Call ShowWelcomeMessage
    
    ' 従業員管理シートをアクティブにする
    Call ActivateEmployeeSheet
    
    Call Xdebug.printx("従業員管理システムの起動が完了しました")
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("Workbook_Open", Err.Description)
    MsgBox "システムの初期化中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' ワークブック終了前の処理
    On Error GoTo ErrorHandler
    
    Call Xdebug.printx("従業員管理システムを終了中...")
    
    ' 変更されたデータの保存確認
    If Me.Saved = False Then
        Dim response As VbMsgBoxResult
        response = MsgBox("従業員データに変更があります。保存しますか？", vbYesNoCancel + vbQuestion, "保存確認")
        
        Select Case response
            Case vbYes
                Me.Save
                Call Xdebug.printx("データを保存しました")
            Case vbNo
                Me.Saved = True ' 保存しないで終了
                Call Xdebug.printx("データを保存せずに終了します")
            Case vbCancel
                Cancel = True ' 終了をキャンセル
                Call Xdebug.printx("終了をキャンセルしました")
                Exit Sub
        End Select
    End If
    
    Call Xdebug.printx("従業員管理システムを終了しました")
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("Workbook_BeforeClose", Err.Description)
End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    ' 保存後の処理
    On Error GoTo ErrorHandler
    
    If Success Then
        Call Xdebug.printx("従業員データを正常に保存しました")
    Else
        Call Xdebug.printError("Workbook_AfterSave", "データの保存に失敗しました")
        MsgBox "データの保存に失敗しました。", vbCritical
    End If
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("Workbook_AfterSave", Err.Description)
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    ' シートがアクティブになった時の処理
    On Error GoTo ErrorHandler
    
    Call Xdebug.printx("シートがアクティブになりました: " & Sh.Name)
    
    ' 従業員管理シート以外がアクティブになった場合の警告
    If Sh.Name <> "従業員管理" And Sh.Name <> "Sheet1" Then
        Call Xdebug.printx("従業員管理シート以外がアクティブになりました")
    End If
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("Workbook_SheetActivate", Err.Description)
End Sub

' ===============================================
' 初期化関連
' ===============================================

Private Sub ShowWelcomeMessage()
    ' ウェルカムメッセージの表示
    On Error GoTo ErrorHandler
    
    Dim welcomeMsg As String
    welcomeMsg = "従業員管理システムへようこそ！" & vbCrLf & vbCrLf
    welcomeMsg = welcomeMsg & "主な機能:" & vbCrLf
    welcomeMsg = welcomeMsg & "• 従業員データの表示・編集" & vbCrLf
    welcomeMsg = welcomeMsg & "• 新規従業員の追加" & vbCrLf
    welcomeMsg = welcomeMsg & "• 従業員の削除" & vbCrLf
    welcomeMsg = welcomeMsg & "• 検索・フィルタ機能" & vbCrLf
    welcomeMsg = welcomeMsg & "• データバリデーション" & vbCrLf & vbCrLf
    welcomeMsg = welcomeMsg & "使用方法:" & vbCrLf
    welcomeMsg = welcomeMsg & "• データをダブルクリックで詳細表示" & vbCrLf
    welcomeMsg = welcomeMsg & "• 上部のボタンで操作" & vbCrLf
    welcomeMsg = welcomeMsg & "• 検索ボックスで従業員検索"
    
    MsgBox welcomeMsg, vbInformation, "従業員管理システム v1.0"
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("ShowWelcomeMessage", Err.Description)
End Sub

Private Sub ActivateEmployeeSheet()
    ' 従業員管理シートをアクティブにする
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    
    ' Sheet1を従業員管理シートとして使用
    Set ws = Me.Worksheets("Sheet1")
    ws.Activate
    
    ' A1セルを選択（開始位置）
    ws.Range("A1").Select
    
    Call Xdebug.printx("従業員管理シートをアクティブにしました")
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("ActivateEmployeeSheet", Err.Description)
End Sub

' ===============================================
' システム管理機能
' ===============================================

Public Sub ResetEmployeeManagementSystem()
    ' システム全体のリセット
    On Error GoTo ErrorHandler
    
    Dim response As VbMsgBoxResult
    response = MsgBox("従業員管理システムをリセットしますか？" & vbCrLf & "すべてのデータが初期状態に戻ります。", vbYesNo + vbQuestion, "システムリセット")
    
    If response = vbYes Then
        Dim ws As Worksheet
        Set ws = Me.Worksheets("Sheet1")
        
        ' データをクリア
        ws.Cells.Clear
        
        ' システムを再初期化
        Call InitializeEmployeeManagementSystem
        
        MsgBox "システムをリセットしました。", vbInformation
        Call Xdebug.printx("従業員管理システムをリセットしました")
    End If
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("ResetEmployeeManagementSystem", Err.Description)
End Sub

Public Sub ExportEmployeeData()
    ' 従業員データのエクスポート機能（簡易版）
    On Error GoTo ErrorHandler
    
    Call Xdebug.printx("従業員データのエクスポートを開始します...")
    
    Dim ws As Worksheet
    Set ws = Me.Worksheets("Sheet1")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_EMPLOYEE_ID).End(xlUp).row
    
    If lastRow < DATA_START_ROW Then
        MsgBox "エクスポートするデータがありません。", vbInformation
        Exit Sub
    End If
    
    ' 簡易エクスポート（クリップボードにコピー）
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, COL_EMAIL)).Copy
    
    MsgBox "従業員データをクリップボードにコピーしました。" & vbCrLf & "他のアプリケーションに貼り付けてください。", vbInformation
    
    Call Xdebug.printx("従業員データのエクスポートが完了しました")
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("ExportEmployeeData", Err.Description)
End Sub

Public Sub ShowSystemInfo()
    ' システム情報の表示
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = Me.Worksheets("Sheet1")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_EMPLOYEE_ID).End(xlUp).row
    
    Dim employeeCount As Long
    employeeCount = lastRow - 1 ' ヘッダー行を除く
    
    Dim systemInfo As String
    systemInfo = "従業員管理システム情報" & vbCrLf & vbCrLf
    systemInfo = systemInfo & "バージョン: 1.0" & vbCrLf
    systemInfo = systemInfo & "開発環境: XVBA" & vbCrLf
    systemInfo = systemInfo & "登録従業員数: " & employeeCount & "名" & vbCrLf
    systemInfo = systemInfo & "最終更新: " & Format(Now, "yyyy/mm/dd hh:nn:ss") & vbCrLf & vbCrLf
    systemInfo = systemInfo & "機能:" & vbCrLf
    systemInfo = systemInfo & "• CRUD操作対応" & vbCrLf
    systemInfo = systemInfo & "• リアルタイムバリデーション" & vbCrLf
    systemInfo = systemInfo & "• 検索・フィルタ機能" & vbCrLf
    systemInfo = systemInfo & "• デバッグサポート (Xdebug)"
    
    MsgBox systemInfo, vbInformation, "システム情報"
    
    Call Xdebug.printx("システム情報を表示しました")
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("ShowSystemInfo", Err.Description)
End Sub
