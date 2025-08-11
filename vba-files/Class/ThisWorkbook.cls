VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' =============================================================================
' 備品貸出管理システム - ワークブックイベントハンドリング
' =============================================================================

Private Sub Workbook_Open()
    On Error GoTo ErrHandler
    
    ' システム初期化
    Call InitializeSystem
    
    ' 監査ログ記録
    Call LogAudit("システム起動", "Workbook opened by " & Application.UserName)
    
    Exit Sub
    
ErrHandler:
    Call LogError("Workbook_Open", Err.Number, Err.Description)
    MsgBox "システム初期化中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' システム初期化処理
Private Sub InitializeSystem()
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    
    ' シート名の初期化（既存のSheet1-5を業務用途にリネーム）
    Call InitializeSheetNames
    
    ' 必要なシートレイアウトが存在しない場合は作成
    Call CheckAndCreateLayouts
    
    ' ダッシュボードを表示
    Call ShowDashboard
    
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Call LogError("InitializeSystem", Err.Number, Err.Description)
End Sub

' シート名初期化（既存シートをリネーム）
Private Sub InitializeSheetNames()
    On Error Resume Next
    
    ' 既存シートを業務用途にリネーム
    If Not GetWorksheet("Dashboard") Is Nothing Then
        ' 既にリネーム済みの場合はスキップ
        Exit Sub
    End If
    
    ' Sheet1-5を業務シート名にリネーム
    ThisWorkbook.Worksheets("Sheet1").Name = SHEET_DASHBOARD
    ThisWorkbook.Worksheets("Sheet2").Name = SHEET_ITEMS
    ThisWorkbook.Worksheets("Sheet3").Name = SHEET_LENDING
    ThisWorkbook.Worksheets("Sheet4").Name = SHEET_INPUT
    ThisWorkbook.Worksheets("Sheet5").Name = SHEET_CONFIG
    
    On Error GoTo 0
End Sub

' レイアウト存在チェックと作成
Private Sub CheckAndCreateLayouts()
    On Error GoTo ErrHandler
    
    Dim needsLayout As Boolean
    needsLayout = False
    
    ' ダッシュボードのレイアウトチェック
    Dim ws As Worksheet
    Set ws = GetWorksheet(SHEET_DASHBOARD)
    If Not ws Is Nothing Then
        If ws.Range("A1").Value = "" Then needsLayout = True
    End If
    
    ' レイアウトが必要な場合は作成
    If needsLayout Then
        Call InitializeAllLayouts
    End If
    
    Exit Sub
    
ErrHandler:
    Call LogError("CheckAndCreateLayouts", Err.Number, Err.Description)
End Sub

' ?? セキュリティ重要事項 ??
' Workbook_BeforeClose と Workbook_BeforeSave イベントは実装禁止
' 理由：自動保存やファイル操作の阻害、セキュリティリスクのため
