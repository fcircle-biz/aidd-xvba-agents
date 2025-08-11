VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' =============================================================================
' 備品貸出管理システム - _Configシートイベントハンドリング
' =============================================================================

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrHandler
    
    ' 設定値の変更をログ記録
    Call LogAudit("システム設定変更", "セル: " & Target.Address & ", 値: " & Target.Value)
    
    Exit Sub
    
ErrHandler:
    Call LogError("Sheet5_Worksheet_Change", Err.Number, Err.Description)
End Sub

Private Sub Worksheet_Activate()
    On Error Resume Next
    
    ' 設定シートのレイアウト確認・作成
    If Range("A1").Value = "" Then
        Call CreateConfigLayout
    End If
    
    On Error GoTo 0
End Sub

' 設定シートレイアウト作成
Private Sub CreateConfigLayout()
    On Error GoTo ErrHandler
    
    ' タイトル
    With Range("A1:D1")
        .Merge
        .Value = "システム設定"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 設定項目
    Range("A3").Value = "デフォルト貸出期間（日）:"
    Range("B3").Value = DEFAULT_LENDING_DAYS
    
    Range("A4").Value = "最大貸出期間（日）:"
    Range("B4").Value = MAX_LENDING_DAYS
    
    Range("A5").Value = "期限警告日数:"
    Range("B5").Value = WARNING_DAYS_BEFORE
    
    Range("A7").Value = "システム情報:"
    Range("A8").Value = "バージョン:"
    Range("B8").Value = "1.0.0"
    
    Range("A9").Value = "最終更新:"
    Range("B9").Value = Format(Now, "yyyy/mm/dd hh:mm:ss")
    
    ' ラベルを太字に
    Range("A3:A9").Font.Bold = True
    
    ' 設定値セルに背景色
    With Range("B3:B5")
        .Interior.Color = RGB(255, 255, 204) ' 薄い黄色
        .Borders.LineStyle = xlContinuous
    End With
    
    Exit Sub
    
ErrHandler:
    Call LogError("CreateConfigLayout", Err.Number, Err.Description)
End Sub
