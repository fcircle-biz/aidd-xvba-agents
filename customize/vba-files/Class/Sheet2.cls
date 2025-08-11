VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' =============================================================================
' 備品貸出管理システム - Itemsシートイベントハンドリング
' =============================================================================

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrHandler
    
    ' 備品マスタの変更を監視し、必要に応じてダッシュボードを更新
    Dim tbl As ListObject
    Set tbl = GetItemsTable()
    
    If Not tbl Is Nothing Then
        If Not Intersect(Target, tbl.DataBodyRange) Is Nothing Then
            ' データ変更をログ記録
            Call LogAudit("備品マスタ更新", "Row: " & Target.Row & ", Column: " & Target.Column & ", Value: " & Target.Value)
            
            ' ダッシュボード更新
            Call UpdateDashboard
        End If
    End If
    
    Exit Sub
    
ErrHandler:
    Call LogError("Sheet2_Worksheet_Change", Err.Number, Err.Description)
End Sub
