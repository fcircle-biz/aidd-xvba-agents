VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'=============================================================================
' Sheet1.cls - Dashboardシートクラス
'=============================================================================
' 概要:
'   Dashboardシートのボタンイベント処理、UI操作制御
'   セキュリティ要件に従いセル変更・シート選択イベントは禁止、ボタンのみで操作
'=============================================================================
Option Explicit

'=============================================================================
' ボタンイベントハンドラー（Dashboardシート用）
'=============================================================================

' メインボタン: CSV一括取込・更新実行
Public Sub ExecuteFullImportProcess()
    On Error Resume Next
    Call modCustomerSystem.ExecuteFullImportProcess()
End Sub

' サブボタン1: Stagingデータクリア
Public Sub ClearStagingData()
    On Error Resume Next
    
    If modDashboard.ConfirmOperation("Stagingデータをクリアします。") Then
        Call modData.ClearStagingData()
        Call modDashboard.RefreshKPI()
        Call modDashboard.ShowStatusMessage("Stagingデータをクリアしました")
    End If
End Sub

' サブボタン2: 設定画面を開く
Public Sub OpenConfigSheet()
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_CONFIG)
    If Not ws Is Nothing Then
        ws.Activate
        ws.Range("A1").Select
        Call modDashboard.ShowStatusMessage("設定画面を開きました")
    Else
        MsgBox "設定シートが見つかりません。", vbExclamation
    End If
End Sub

' サブボタン3: 差分レポート出力
Public Sub ExportDifferenceReport()
    On Error Resume Next
    Call modDashboard.GenerateDifferenceReport()
End Sub

' サブボタン4: KPI表示更新
Public Sub RefreshKPI()
    On Error Resume Next
    Call modDashboard.RefreshKPI()
    Call modDashboard.ShowStatusMessage("KPI表示を更新しました")
End Sub

' サブボタン5: 期限切れ顧客無効化
Public Sub InactivateStaleCustomers()
    On Error Resume Next
    
    If modDashboard.ConfirmOperation("期限切れ顧客データを無効化します。") Then
        Dim inactivatedCount As Long
        inactivatedCount = modUpsert.InactivateStaleCustomers()
        
        If inactivatedCount >= 0 Then
            Call modDashboard.RefreshKPI()
            Call modDashboard.ShowStatusMessage("期限切れ顧客 " & inactivatedCount & " 件を無効化しました")
        Else
            Call modDashboard.ShowStatusMessage("無効化処理でエラーが発生しました", isError:=True)
        End If
    End If
End Sub

'=============================================================================
' 追加操作ボタン（必要に応じて）
'=============================================================================

' CSV取込のみ実行
Public Sub ExecuteImportOnly()
    On Error Resume Next
    Call modCustomerSystem.ExecuteImportOnly()
End Sub

' データ検証のみ実行
Public Sub ExecuteValidationOnly()
    On Error Resume Next
    Call modCustomerSystem.ExecuteValidationOnly()
End Sub

' アップサートのみ実行
Public Sub ExecuteUpsertOnly()
    On Error Resume Next
    Call modCustomerSystem.ExecuteUpsertOnly()
End Sub

' 顧客データエクスポート
Public Sub ExportCustomerData()
    On Error Resume Next
    Call modCustomerSystem.ExportCustomerData()
End Sub

' システム健全性チェック
Public Sub CheckSystemHealth()
    On Error Resume Next
    Call modCustomerSystem.ExecuteSystemHealthCheck()
End Sub

' システム情報表示
Public Sub ShowSystemInfo()
    On Error Resume Next
    Call modCustomerSystem.ShowSystemInformation()
End Sub

' データベース最適化
Public Sub OptimizeDatabase()
    On Error Resume Next
    Call modCustomerSystem.OptimizeDatabase()
End Sub

' 全データクリア（開発・テスト用）
Public Sub ClearAllData()
    On Error Resume Next
    Call modCustomerSystem.ClearAllData()
End Sub

'=============================================================================
' セキュリティ制限により以下のイベントは実装禁止
' - Worksheet_Change
' - Worksheet_SelectionChange
' - その他のワークシートイベント
'=============================================================================
