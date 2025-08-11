VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'=============================================================================
' ThisWorkbook.cls - ワークブッククラス（システム制御）
'=============================================================================
' 概要:
'   ワークブック全体の制御、システム初期化、開閉時イベント処理
'   セキュリティ要件に従いWorkbook_Open以外の自動イベントは禁止
'=============================================================================
Option Explicit

'=============================================================================
' ワークブックイベント処理（セキュリティ制限により最小限）
'=============================================================================

' ワークブック開時処理（唯一許可されたワークブックイベント）
Private Sub Workbook_Open()
    On Error GoTo ErrHandler
    
    ' システム初期化実行
    Call modCustomerSystem.InitializeSystem()
    
    ' Dashboardシートをアクティブ化
    Call ActivateDashboard()
    
    Exit Sub
    
ErrHandler:
    Call modCmn.LogError("Workbook_Open", "ワークブック開時エラー: " & Err.Description)
    
    ' エラーが発生してもシステムを起動
    MsgBox "システム初期化中にエラーが発生しました。" & vbCrLf & _
           "一部機能が正常に動作しない可能性があります。" & vbCrLf & vbCrLf & _
           "エラー内容: " & Err.Description, vbExclamation, "初期化警告"
End Sub

' Dashboardシートアクティブ化
Private Sub ActivateDashboard()
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_DASHBOARD)
    If Not ws Is Nothing Then
        ws.Activate
        ws.Range("A1").Select
    End If
End Sub

'=============================================================================
' セキュリティ制限により以下のイベントは実装禁止
' - Workbook_BeforeClose
' - Workbook_BeforeSave
' - その他のワークブックイベント
'=============================================================================

'=============================================================================
' パブリック制御メソッド（ボタンから呼び出し用）
'=============================================================================

' システム強制初期化（緊急時用）
Public Sub ForceInitializeSystem()
    On Error Resume Next
    Call modCustomerSystem.InitializeSystem()
End Sub

' システム情報表示
Public Sub ShowSystemInfo()
    On Error Resume Next
    Call modCustomerSystem.ShowSystemInformation()
End Sub

' システム健全性チェック
Public Sub CheckSystemHealth()
    On Error Resume Next
    Call modCustomerSystem.ExecuteSystemHealthCheck()
End Sub

' 緊急停止
Public Sub EmergencyStop()
    On Error Resume Next
    Call modCustomerSystem.EmergencyStop()
End Sub
