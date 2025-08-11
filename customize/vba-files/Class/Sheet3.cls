VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' =============================================================================
' 備品貸出管理システム - Lendingシートイベントハンドリング
' =============================================================================

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrHandler
    
    ' 貸出履歴の変更を監視し、必要に応じてダッシュボードを更新
    Dim tbl As ListObject
    Set tbl = GetLendingTable()
    
    If Not tbl Is Nothing Then
        If Not Intersect(Target, tbl.DataBodyRange) Is Nothing Then
            ' データ変更をログ記録
            Call LogAudit("貸出履歴更新", "Row: " & Target.Row & ", Column: " & Target.Column & ", Value: " & Target.Value)
            
            ' ダッシュボード更新
            Call UpdateDashboard
        End If
    End If
    
    Exit Sub
    
ErrHandler:
    Call LogError("Sheet3_Worksheet_Change", Err.Number, Err.Description)
End Sub

Private Sub Worksheet_Activate()
    On Error Resume Next
    
    ' 貸出履歴シートがアクティブになったとき、条件付き書式を適用
    Call ApplyLendingStatusFormatting
    
    On Error GoTo 0
End Sub

' 貸出履歴のステータス色分け適用
Private Sub ApplyLendingStatusFormatting()
    On Error GoTo ErrHandler
    
    Dim tbl As ListObject
    Set tbl = GetLendingTable()
    
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    
    Dim statusCol As Long, dueDateCol As Long
    statusCol = GetColumnIndex(tbl, COL_STATUS)
    dueDateCol = GetColumnIndex(tbl, COL_DUE_DATE)
    
    If statusCol = 0 Or dueDateCol = 0 Then Exit Sub
    
    Dim i As Long
    For i = 1 To tbl.DataBodyRange.Rows.Count
        Dim status As String, dueDate As Date
        status = tbl.DataBodyRange.Cells(i, statusCol).Value
        
        If status = STATUS_LENDING Then
            If IsDate(tbl.DataBodyRange.Cells(i, dueDateCol).Value) Then
                dueDate = CDate(tbl.DataBodyRange.Cells(i, dueDateCol).Value)
                If dueDate < Date Then
                    ' 期限超過は赤
                    tbl.DataBodyRange.Rows(i).Interior.Color = COLOR_OVERDUE
                    tbl.DataBodyRange.Rows(i).Font.Color = vbWhite
                ElseIf dueDate <= Date + WARNING_DAYS_BEFORE Then
                    ' 期限間近は黄
                    tbl.DataBodyRange.Rows(i).Interior.Color = COLOR_WARNING
                    tbl.DataBodyRange.Rows(i).Font.Color = vbBlack
                End If
            End If
        ElseIf status = STATUS_RETURNED Then
            ' 返却済みは緑
            tbl.DataBodyRange.Rows(i).Interior.Color = COLOR_SUCCESS
            tbl.DataBodyRange.Rows(i).Font.Color = vbWhite
        End If
    Next i
    
    Exit Sub
    
ErrHandler:
    Call LogError("ApplyLendingStatusFormatting", Err.Number, Err.Description)
End Sub
