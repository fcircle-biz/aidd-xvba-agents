VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' ===============================================
' 従業員管理シート - Sheet1クラス
' ===============================================

Option Explicit

' ===============================================
' ワークシートイベント
' ===============================================

Private Sub Worksheet_Activate()
    ' シートがアクティブになった時の処理
    On Error GoTo ErrorHandler
    
    Call Xdebug.printx("従業員管理シートがアクティブになりました")
    
    ' シート名を従業員管理に変更（まだされていない場合）
    If Me.Name <> "従業員管理" Then
        Me.Name = "従業員管理"
    End If
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("Worksheet_Activate", Err.Description)
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    ' セルが変更された時のバリデーション
    On Error GoTo ErrorHandler
    
    ' データ行の範囲内での変更のみ処理
    If Target.row >= DATA_START_ROW And Target.Column <= COL_EMAIL Then
        Call ValidateChangedData(Target)
    End If
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("Worksheet_Change", Err.Description)
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' セル選択が変更された時の処理
    On Error GoTo ErrorHandler
    
    ' 選択された行がデータ行の場合、行全体をハイライト
    If Target.row >= DATA_START_ROW And Target.Cells.Count = 1 Then
        Call HighlightSelectedEmployeeRow(Target.row)
    End If
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("Worksheet_SelectionChange", Err.Description)
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' ダブルクリック時の詳細編集
    On Error GoTo ErrorHandler
    
    ' データ行のダブルクリックで詳細編集画面を表示
    If Target.row >= DATA_START_ROW And Target.Column <= COL_EMAIL Then
        If Me.Cells(Target.row, COL_EMPLOYEE_ID).Value <> "" Then
            Call ShowEmployeeDetailDialog(Target.row)
            Cancel = True ' デフォルトの編集モードをキャンセル
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("Worksheet_BeforeDoubleClick", Err.Description)
End Sub

' ===============================================
' データバリデーション
' ===============================================

Private Sub ValidateChangedData(Target As Range)
    ' 変更されたデータのリアルタイムバリデーション
    On Error GoTo ErrorHandler
    
    Dim isValid As Boolean
    isValid = True
    
    ' 列に応じたバリデーション
    Select Case Target.Column
        Case COL_EMPLOYEE_ID
            ' 従業員IDの重複チェック
            If CheckDuplicateEmployeeID(Target.Value, Target.row) Then
                MsgBox "従業員ID「" & Target.Value & "」は既に使用されています。", vbExclamation
                Application.Undo
                Exit Sub
            End If
            
        Case COL_SALARY
            ' 給与の数値チェック
            If Target.Value <> "" And (Not IsNumeric(Target.Value) Or Target.Value < 0) Then
                MsgBox "給与は正の数値で入力してください。", vbExclamation
                Application.Undo
                Exit Sub
            End If
            
        Case COL_EMAIL
            ' メールアドレスの形式チェック
            If Target.Value <> "" And InStr(Target.Value, "@") = 0 Then
                MsgBox "メールアドレスの形式が正しくありません。", vbExclamation
                Application.Undo
                Exit Sub
            End If
            
        Case COL_HIRE_DATE
            ' 入社日の日付チェック
            If Target.Value <> "" And Not IsDate(Target.Value) Then
                MsgBox "入社日は正しい日付形式で入力してください。", vbExclamation
                Application.Undo
                Exit Sub
            End If
    End Select
    
    ' データが変更されたことをログに記録
    Call Xdebug.printx("データが変更されました - 行:" & Target.row & " 列:" & Target.Column & " 値:" & Target.Value)
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("ValidateChangedData", Err.Description)
End Sub

Private Function CheckDuplicateEmployeeID(employeeID As String, currentRow As Long) As Boolean
    ' 従業員IDの重複チェック
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = Me.Cells(Me.Rows.Count, COL_EMPLOYEE_ID).End(xlUp).row
    
    For i = DATA_START_ROW To lastRow
        If i <> currentRow And Me.Cells(i, COL_EMPLOYEE_ID).Value = employeeID Then
            CheckDuplicateEmployeeID = True
            Exit Function
        End If
    Next i
    
    CheckDuplicateEmployeeID = False
    Exit Function
    
ErrorHandler:
    Call Xdebug.printError("CheckDuplicateEmployeeID", Err.Description)
    CheckDuplicateEmployeeID = False
End Function

' ===============================================
' UI関連機能
' ===============================================

Private Sub HighlightSelectedEmployeeRow(selectedRow As Long)
    ' 選択された従業員行をハイライト
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    lastRow = Me.Cells(Me.Rows.Count, COL_EMPLOYEE_ID).End(xlUp).row
    
    ' すべての行の背景色をリセット
    For i = DATA_START_ROW To lastRow
        If i Mod 2 = 0 Then
            Me.Range(Me.Cells(i, 1), Me.Cells(i, COL_EMAIL)).Interior.Color = RGB(245, 245, 245)
        Else
            Me.Range(Me.Cells(i, 1), Me.Cells(i, COL_EMAIL)).Interior.ColorIndex = xlNone
        End If
    Next i
    
    ' 選択行をハイライト
    If selectedRow >= DATA_START_ROW And selectedRow <= lastRow Then
        Me.Range(Me.Cells(selectedRow, 1), Me.Cells(selectedRow, COL_EMAIL)).Interior.Color = RGB(255, 255, 150)
    End If
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("HighlightSelectedEmployeeRow", Err.Description)
End Sub

Private Sub ShowEmployeeDetailDialog(employeeRow As Long)
    ' 従業員詳細ダイアログの表示（簡易版）
    On Error GoTo ErrorHandler
    
    Dim employeeInfo As String
    employeeInfo = "従業員詳細情報" & vbCrLf & vbCrLf
    employeeInfo = employeeInfo & "ID: " & Me.Cells(employeeRow, COL_EMPLOYEE_ID).Value & vbCrLf
    employeeInfo = employeeInfo & "氏名: " & Me.Cells(employeeRow, COL_NAME).Value & vbCrLf
    employeeInfo = employeeInfo & "部署: " & Me.Cells(employeeRow, COL_DEPARTMENT).Value & vbCrLf
    employeeInfo = employeeInfo & "役職: " & Me.Cells(employeeRow, COL_POSITION).Value & vbCrLf
    employeeInfo = employeeInfo & "入社日: " & Me.Cells(employeeRow, COL_HIRE_DATE).Value & vbCrLf
    employeeInfo = employeeInfo & "給与: " & Format(Me.Cells(employeeRow, COL_SALARY).Value, "#,##0") & "円" & vbCrLf
    employeeInfo = employeeInfo & "電話: " & Me.Cells(employeeRow, COL_PHONE).Value & vbCrLf
    employeeInfo = employeeInfo & "Email: " & Me.Cells(employeeRow, COL_EMAIL).Value
    
    MsgBox employeeInfo, vbInformation, "従業員詳細"
    
    Call Xdebug.printx("従業員詳細を表示: " & Me.Cells(employeeRow, COL_NAME).Value)
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("ShowEmployeeDetailDialog", Err.Description)
End Sub

' ===============================================
' データ操作支援機能
' ===============================================

Public Sub RefreshEmployeeView()
    ' 従業員ビューの再描画
    On Error GoTo ErrorHandler
    
    Call Xdebug.printx("従業員ビューを更新中...")
    
    ' データ範囲の再フォーマット
    Call FormatEmployeeDataRange(Me)
    
    ' 列幅の自動調整
    Me.Columns("A:H").AutoFit
    
    ' シートを先頭に移動
    Me.Cells(1, 1).Select
    
    Call Xdebug.printx("従業員ビューの更新が完了しました")
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("RefreshEmployeeView", Err.Description)
End Sub

Public Function GetSelectedEmployeeID() As String
    ' 選択された従業員のIDを取得
    On Error GoTo ErrorHandler
    
    Dim selectedRow As Long
    selectedRow = ActiveCell.row
    
    If selectedRow >= DATA_START_ROW Then
        GetSelectedEmployeeID = Me.Cells(selectedRow, COL_EMPLOYEE_ID).Value
    Else
        GetSelectedEmployeeID = ""
    End If
    
    Exit Function
    
ErrorHandler:
    Call Xdebug.printError("GetSelectedEmployeeID", Err.Description)
    GetSelectedEmployeeID = ""
End Function
