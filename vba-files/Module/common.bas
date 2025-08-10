Attribute VB_Name = "common"

' ===============================================
' 従業員管理システム - 共通モジュール
' ===============================================

Option Explicit

' 従業員データの列番号定数
Public Const COL_EMPLOYEE_ID As Integer = 1
Public Const COL_NAME As Integer = 2
Public Const COL_DEPARTMENT As Integer = 3
Public Const COL_POSITION As Integer = 4
Public Const COL_HIRE_DATE As Integer = 5
Public Const COL_SALARY As Integer = 6
Public Const COL_PHONE As Integer = 7
Public Const COL_EMAIL As Integer = 8

' データ行の開始位置（ヘッダー行の次）
Public Const DATA_START_ROW As Integer = 2

' ===============================================
' システム初期化関連
' ===============================================

Public Sub InitializeEmployeeManagementSystem()
    ' 従業員管理システムの初期化
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    
    Call Xdebug.printx("従業員管理システムを初期化中...")
    
    ' ヘッダーの設定
    Call SetupEmployeeHeaders(ws)
    
    ' 既存データがない場合のみサンプルデータを作成
    If IsEmployeeDataEmpty(ws) Then
        Call CreateSampleEmployeeData(ws)
        Call Xdebug.printx("サンプル従業員データを作成しました")
    Else
        Call Xdebug.printx("既存の従業員データが見つかりました")
    End If
    
    ' UI要素の設定
    Call SetupEmployeeUI(ws)
    
    Call Xdebug.printx("従業員管理システムの初期化が完了しました")
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("InitializeEmployeeManagementSystem", Err.Description)
End Sub

Public Sub SetupEmployeeHeaders(ws As Worksheet)
    ' ヘッダー行の設定
    On Error GoTo ErrorHandler
    
    With ws
        .Cells(1, COL_EMPLOYEE_ID).Value = "従業員ID"
        .Cells(1, COL_NAME).Value = "氏名"
        .Cells(1, COL_DEPARTMENT).Value = "部署"
        .Cells(1, COL_POSITION).Value = "役職"
        .Cells(1, COL_HIRE_DATE).Value = "入社日"
        .Cells(1, COL_SALARY).Value = "給与"
        .Cells(1, COL_PHONE).Value = "電話番号"
        .Cells(1, COL_EMAIL).Value = "メールアドレス"
        
        ' ヘッダー行のスタイル設定
        With .Range(.Cells(1, 1), .Cells(1, COL_EMAIL))
            .Font.Bold = True
            .Interior.Color = RGB(200, 220, 240)
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
        End With
        
        ' 列幅の自動調整
        .Columns("A:H").AutoFit
    End With
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("SetupEmployeeHeaders", Err.Description)
End Sub

Public Function IsEmployeeDataEmpty(ws As Worksheet) As Boolean
    ' 従業員データが空かどうかをチェック
    On Error GoTo ErrorHandler
    
    IsEmployeeDataEmpty = (ws.Cells(DATA_START_ROW, COL_EMPLOYEE_ID).Value = "")
    Exit Function
    
ErrorHandler:
    Call Xdebug.printError("IsEmployeeDataEmpty", Err.Description)
    IsEmployeeDataEmpty = True
End Function

' ===============================================
' サンプルデータ生成
' ===============================================

Public Sub CreateSampleEmployeeData(ws As Worksheet)
    ' 30件のリアルな従業員サンプルデータを作成
    On Error GoTo ErrorHandler
    
    ' サンプル従業員データを個別に追加する方式に変更
    Call AddEmployeeRecord(ws, "EMP001", "山田 太郎", "営業部", "部長", "2015-04-01", 8500000, "03-1234-5678", "yamada.taro@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP002", "佐藤 花子", "人事部", "課長", "2016-10-15", 7200000, "03-1234-5679", "sato.hanako@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP003", "田中 次郎", "技術部", "主任", "2017-08-20", 6500000, "03-1234-5680", "tanaka.jiro@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP004", "高橋 美咲", "経理部", "係長", "2018-03-10", 5800000, "03-1234-5681", "takahashi.misaki@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP005", "伊藤 和也", "営業部", "主任", "2018-07-01", 6200000, "03-1234-5682", "ito.kazuya@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP006", "渡辺 聡美", "総務部", "課長", "2014-09-15", 7000000, "03-1234-5683", "watanabe.satomi@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP007", "小林 健太", "技術部", "エンジニア", "2019-04-01", 5500000, "03-1234-5684", "kobayashi.kenta@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP008", "加藤 由美", "マーケティング部", "スペシャリスト", "2020-01-20", 6800000, "03-1234-5685", "kato.yumi@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP009", "吉田 雄一", "営業部", "係長", "2017-12-01", 6000000, "03-1234-5686", "yoshida.yuichi@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP010", "中村 麻衣", "人事部", "アシスタント", "2021-03-15", 4800000, "03-1234-5687", "nakamura.mai@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP011", "林 大輔", "技術部", "シニアエンジニア", "2016-05-10", 7500000, "03-1234-5688", "hayashi.daisuke@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP012", "森 智子", "経理部", "主任", "2019-08-25", 5700000, "03-1234-5689", "mori.tomoko@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP013", "池田 慎一", "営業部", "マネージャー", "2013-11-01", 8200000, "03-1234-5690", "ikeda.shinichi@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP014", "橋本 恵子", "総務部", "アシスタント", "2022-04-01", 4500000, "03-1234-5691", "hashimoto.keiko@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP015", "石川 浩二", "技術部", "テクニカルリード", "2015-07-15", 8000000, "03-1234-5692", "ishikawa.koji@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP016", "前田 結香", "マーケティング部", "アナリスト", "2020-09-01", 5900000, "03-1234-5693", "maeda.yuka@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP017", "岡田 竜也", "営業部", "営業担当", "2021-06-10", 5200000, "03-1234-5694", "okada.tatsuya@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP018", "村上 理恵", "人事部", "採用担当", "2018-12-01", 6300000, "03-1234-5695", "murakami.rie@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP019", "清水 拓海", "技術部", "エンジニア", "2022-01-15", 5400000, "03-1234-5696", "shimizu.takumi@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP020", "山口 愛", "経理部", "経理担当", "2019-11-20", 5600000, "03-1234-5697", "yamaguchi.ai@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP021", "松本 正樹", "営業部", "シニアセールス", "2016-02-28", 7300000, "03-1234-5698", "matsumoto.masaki@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP022", "井上 千春", "総務部", "総務担当", "2020-05-15", 5100000, "03-1234-5699", "inoue.chiharu@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP023", "木村 健", "技術部", "プロジェクトマネージャー", "2014-08-01", 8800000, "03-1234-5700", "kimura.ken@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP024", "斉藤 美穂", "マーケティング部", "マネージャー", "2017-03-20", 7800000, "03-1234-5701", "saito.miho@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP025", "中島 博之", "営業部", "営業担当", "2021-09-01", 5300000, "03-1234-5702", "nakajima.hiroyuki@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP026", "原田 さやか", "人事部", "人事担当", "2019-06-10", 5800000, "03-1234-5703", "harada.sayaka@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP027", "坂本 直樹", "技術部", "システムアーキテクト", "2015-12-01", 9200000, "03-1234-5704", "sakamoto.naoki@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP028", "青木 優香", "経理部", "財務アナリスト", "2020-08-15", 6700000, "03-1234-5705", "aoki.yuka@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP029", "藤田 光男", "営業部", "営業担当", "2022-02-01", 5000000, "03-1234-5706", "fujita.mitsuo@company.co.jp")
    Call AddEmployeeRecord(ws, "EMP030", "西村 あかね", "総務部", "秘書", "2021-11-15", 4700000, "03-1234-5707", "nishimura.akane@company.co.jp")
    
    ' データ範囲の書式設定
    Call FormatEmployeeDataRange(ws)
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("CreateSampleEmployeeData", Err.Description)
End Sub

Public Sub AddEmployeeRecord(ws As Worksheet, empID As String, empName As String, dept As String, pos As String, hireDate As String, salary As Long, phone As String, email As String)
    ' 従業員レコードを追加するヘルパー関数
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_EMPLOYEE_ID).End(xlUp).Row + 1
    
    ws.Cells(lastRow, COL_EMPLOYEE_ID).Value = empID
    ws.Cells(lastRow, COL_NAME).Value = empName
    ws.Cells(lastRow, COL_DEPARTMENT).Value = dept
    ws.Cells(lastRow, COL_POSITION).Value = pos
    ws.Cells(lastRow, COL_HIRE_DATE).Value = CDate(hireDate)
    ws.Cells(lastRow, COL_SALARY).Value = salary
    ws.Cells(lastRow, COL_PHONE).Value = phone
    ws.Cells(lastRow, COL_EMAIL).Value = email
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("AddEmployeeRecord", Err.Description)
End Sub

Public Sub FormatEmployeeDataRange(ws As Worksheet)
    ' データ範囲の書式設定
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    Dim i As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_EMPLOYEE_ID).End(xlUp).row
    
    With ws.Range(ws.Cells(DATA_START_ROW, 1), ws.Cells(lastRow, COL_EMAIL))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        
        ' 給与列の書式設定（カンマ区切り）
        ws.Range(ws.Cells(DATA_START_ROW, COL_SALARY), ws.Cells(lastRow, COL_SALARY)).NumberFormat = "#,##0"
        
        ' 入社日列の書式設定
        ws.Range(ws.Cells(DATA_START_ROW, COL_HIRE_DATE), ws.Cells(lastRow, COL_HIRE_DATE)).NumberFormat = "yyyy/mm/dd"
        
        ' 交互の行の色付け
        For i = DATA_START_ROW To lastRow
            If i Mod 2 = 0 Then
                ws.Range(ws.Cells(i, 1), ws.Cells(i, COL_EMAIL)).Interior.Color = RGB(245, 245, 245)
            End If
        Next i
    End With
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("FormatEmployeeDataRange", Err.Description)
End Sub

' ===============================================
' UI要素設定
' ===============================================

Public Sub SetupEmployeeUI(ws As Worksheet)
    ' 従業員管理用のUI要素（ボタンなど）を設定
    On Error GoTo ErrorHandler
    
    Call Xdebug.printx("従業員管理UIを設定中...")
    
    ' 既存のボタンを削除
    Call ClearExistingButtons(ws)
    
    ' 新規従業員追加ボタン
    Call CreateButton(ws, "新規追加", 10, 10, 80, 25, "common.AddNewEmployee_Click")
    
    ' 削除ボタン
    Call CreateButton(ws, "選択行削除", 100, 10, 80, 25, "common.DeleteSelectedEmployee_Click")
    
    ' 検索ボタン
    Call CreateButton(ws, "検索", 190, 10, 60, 25, "common.SearchEmployee_Click")
    
    ' リセットボタン
    Call CreateButton(ws, "リセット", 260, 10, 60, 25, "common.ResetView_Click")
    
    ' 検索ボックス
    Call CreateSearchBox(ws)
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("SetupEmployeeUI", Err.Description)
End Sub

Public Sub ClearExistingButtons(ws As Worksheet)
    ' 既存のボタンを削除
    On Error Resume Next
    Dim obj As Object
    For Each obj In ws.Shapes
        If obj.Type = 1 Then ' msoFormControl
            obj.Delete
        End If
    Next obj
    On Error GoTo 0
End Sub

Public Sub CreateButton(ws As Worksheet, buttonText As String, left As Double, top As Double, width As Double, height As Double, macroName As String)
    ' ボタンを作成
    On Error GoTo ErrorHandler
    
    Dim btn As Object
    Set btn = ws.Shapes.AddFormControl(xlButtonControl, left, top, width, height)
    btn.TextFrame.Characters.Text = buttonText
    btn.OnAction = macroName
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("CreateButton", Err.Description)
End Sub

Public Sub CreateSearchBox(ws As Worksheet)
    ' 検索用テキストボックスを作成
    On Error GoTo ErrorHandler
    
    ' 検索ラベル
    ws.Cells(1, 10).Value = "検索:"
    ws.Cells(1, 10).Font.Bold = True
    
    ' 検索ボックス（セル）
    ws.Cells(1, 11).Value = ""
    ws.Cells(1, 11).Interior.Color = RGB(255, 255, 200)
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("CreateSearchBox", Err.Description)
End Sub

' ===============================================
' データ操作関数（CRUD）
' ===============================================

Public Sub AddNewEmployee_Click()
    ' 新規従業員追加
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_EMPLOYEE_ID).End(xlUp).row + 1
    
    ' 新規従業員IDを生成
    Dim newEmployeeID As String
    newEmployeeID = "EMP" & Format(lastRow - 1, "000")
    
    ' 新しい行にデフォルト値を設定
    ws.Cells(lastRow, COL_EMPLOYEE_ID).Value = newEmployeeID
    ws.Cells(lastRow, COL_NAME).Value = "新規従業員"
    ws.Cells(lastRow, COL_DEPARTMENT).Value = "未設定"
    ws.Cells(lastRow, COL_POSITION).Value = "未設定"
    ws.Cells(lastRow, COL_HIRE_DATE).Value = Date
    ws.Cells(lastRow, COL_SALARY).Value = 0
    ws.Cells(lastRow, COL_PHONE).Value = ""
    ws.Cells(lastRow, COL_EMAIL).Value = ""
    
    ' 新しい行を選択
    ws.Cells(lastRow, COL_NAME).Select
    
    Call Xdebug.printx("新規従業員を追加しました: " & newEmployeeID)
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("AddNewEmployee_Click", Err.Description)
End Sub

Public Sub DeleteSelectedEmployee_Click()
    ' 選択された従業員を削除
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim selectedRow As Long
    selectedRow = Selection.row
    
    ' ヘッダー行や空行の削除を防ぐ
    If selectedRow < DATA_START_ROW Then
        MsgBox "ヘッダー行は削除できません。", vbExclamation
        Exit Sub
    End If
    
    If ws.Cells(selectedRow, COL_EMPLOYEE_ID).Value = "" Then
        MsgBox "削除する従業員が選択されていません。", vbExclamation
        Exit Sub
    End If
    
    ' 確認メッセージ
    Dim employeeName As String
    employeeName = ws.Cells(selectedRow, COL_NAME).Value
    
    If MsgBox("従業員「" & employeeName & "」を削除しますか？", vbYesNo + vbQuestion) = vbYes Then
        ws.Rows(selectedRow).Delete
        Call Xdebug.printx("従業員を削除しました: " & employeeName)
        MsgBox "従業員を削除しました。", vbInformation
    End If
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("DeleteSelectedEmployee_Click", Err.Description)
End Sub

' ===============================================
' 検索・フィルタ機能
' ===============================================

Public Sub SearchEmployee_Click()
    ' 従業員検索機能
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim searchTerm As String
    searchTerm = ws.Cells(1, 11).Value ' 検索ボックスの値を取得
    
    If searchTerm = "" Then
        MsgBox "検索キーワードを入力してください。", vbExclamation
        Exit Sub
    End If
    
    Call FilterEmployeeData(ws, searchTerm)
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("SearchEmployee_Click", Err.Description)
End Sub

Public Sub FilterEmployeeData(ws As Worksheet, searchTerm As String)
    ' データをフィルタリング
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_EMPLOYEE_ID).End(xlUp).row
    
    ' オートフィルタを適用
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
    
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, COL_EMAIL)).AutoFilter
    
    ' 複数列での検索（名前、部署、役職）
    Dim foundMatch As Boolean
    foundMatch = False
    
    ' 名前での検索
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, COL_EMAIL)).AutoFilter Field:=COL_NAME, Criteria1:="*" & searchTerm & "*"
    If CountVisibleRows(ws) > 1 Then
        foundMatch = True
    Else
        ' 部署での検索
        ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, COL_EMAIL)).AutoFilter Field:=COL_NAME
        ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, COL_EMAIL)).AutoFilter Field:=COL_DEPARTMENT, Criteria1:="*" & searchTerm & "*"
        If CountVisibleRows(ws) > 1 Then
            foundMatch = True
        Else
            ' 役職での検索
            ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, COL_EMAIL)).AutoFilter Field:=COL_DEPARTMENT
            ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, COL_EMAIL)).AutoFilter Field:=COL_POSITION, Criteria1:="*" & searchTerm & "*"
            If CountVisibleRows(ws) > 1 Then
                foundMatch = True
            End If
        End If
    End If
    
    If Not foundMatch Then
        MsgBox "「" & searchTerm & "」に一致する従業員が見つかりませんでした。", vbInformation
    End If
    
    Call Xdebug.printx("検索実行: " & searchTerm)
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("FilterEmployeeData", Err.Description)
End Sub

Public Function CountVisibleRows(ws As Worksheet) As Long
    ' 可視行数をカウント
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    Dim visibleCount As Long
    Dim i As Long
    
    lastRow = ws.Cells(ws.Rows.Count, COL_EMPLOYEE_ID).End(xlUp).row
    visibleCount = 0
    
    For i = 1 To lastRow
        If Not ws.Rows(i).Hidden Then
            visibleCount = visibleCount + 1
        End If
    Next i
    
    CountVisibleRows = visibleCount
    Exit Function
    
ErrorHandler:
    Call Xdebug.printError("CountVisibleRows", Err.Description)
    CountVisibleRows = 0
End Function

Public Sub ResetView_Click()
    ' ビューをリセット（フィルタを解除）
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' フィルタを解除
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
    
    ' 検索ボックスをクリア
    ws.Cells(1, 11).Value = ""
    
    Call Xdebug.printx("ビューをリセットしました")
    MsgBox "フィルタを解除しました。", vbInformation
    
    Exit Sub
    
ErrorHandler:
    Call Xdebug.printError("ResetView_Click", Err.Description)
End Sub

' ===============================================
' バリデーション関数
' ===============================================

Public Function ValidateEmployeeData(ws As Worksheet, row As Long) As Boolean
    ' 従業員データのバリデーション
    On Error GoTo ErrorHandler
    
    Dim isValid As Boolean
    isValid = True
    Dim errorMsg As String
    errorMsg = ""
    
    ' 従業員IDの検証
    If ws.Cells(row, COL_EMPLOYEE_ID).Value = "" Then
        isValid = False
        errorMsg = errorMsg & "従業員IDが入力されていません。" & vbCrLf
    End If
    
    ' 氏名の検証
    If ws.Cells(row, COL_NAME).Value = "" Then
        isValid = False
        errorMsg = errorMsg & "氏名が入力されていません。" & vbCrLf
    End If
    
    ' 部署の検証
    If ws.Cells(row, COL_DEPARTMENT).Value = "" Then
        isValid = False
        errorMsg = errorMsg & "部署が入力されていません。" & vbCrLf
    End If
    
    ' 給与の検証
    If Not IsNumeric(ws.Cells(row, COL_SALARY).Value) Or ws.Cells(row, COL_SALARY).Value < 0 Then
        isValid = False
        errorMsg = errorMsg & "給与は正の数値で入力してください。" & vbCrLf
    End If
    
    ' メールアドレスの簡単な検証
    Dim email As String
    email = ws.Cells(row, COL_EMAIL).Value
    If email <> "" And InStr(email, "@") = 0 Then
        isValid = False
        errorMsg = errorMsg & "メールアドレスの形式が正しくありません。" & vbCrLf
    End If
    
    If Not isValid Then
        MsgBox "データの検証に失敗しました:" & vbCrLf & errorMsg, vbExclamation
    End If
    
    ValidateEmployeeData = isValid
    Exit Function
    
ErrorHandler:
    Call Xdebug.printError("ValidateEmployeeData", Err.Description)
    ValidateEmployeeData = False
End Function
