VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' =============================================================================
' 備品貸出管理システム - Inputシートイベントハンドリング
' =============================================================================

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    
    ' 入力値の自動検証・補正処理
    If Target.Address = Range(INPUT_ITEM_ID).Address Then
        ' 備品IDが入力された場合の処理
        If IsNumeric(Target.Value) Then
            Dim itemName As String
            itemName = GetItemName(CLng(Target.Value))
            If itemName <> "" Then
                ' 備品名をヒント表示（隣接セル等）
                Range("C2").Value = "(" & itemName & ")"
                Range("C2").Font.Color = RGB(0, 128, 0)
            Else
                Range("C2").Value = "(備品が見つかりません)"
                Range("C2").Font.Color = RGB(255, 0, 0)
            End If
        Else
            Range("C2").Value = ""
        End If
    End If
    
    On Error GoTo 0
End Sub

Private Sub Worksheet_Activate()
    On Error Resume Next
    
    ' 入力シートがアクティブになったとき、入力ガイドを表示
    Application.StatusBar = "備品IDと借用者を入力して、ダッシュボードの登録ボタンを押してください"
    
    On Error GoTo 0
End Sub

Private Sub Worksheet_Deactivate()
    On Error Resume Next
    
    ' 他シートに移動時、ステータスバーをクリア
    Application.StatusBar = False
    
    On Error GoTo 0
End Sub
