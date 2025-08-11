VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' =============================================================================
' 備品貸出管理システム - Dashboardシートイベントハンドリング
' =============================================================================

Private Sub Worksheet_Activate()
    On Error Resume Next
    
    ' ダッシュボードを最新状態に更新
    Call UpdateDashboard
    
    On Error GoTo 0
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next
    
    ' 必要に応じてリアルタイム更新処理を追加
    
    On Error GoTo 0
End Sub
