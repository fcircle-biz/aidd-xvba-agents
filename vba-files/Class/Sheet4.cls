VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'=============================================================================
' Sheet4.cls - _Configシートクラス
'=============================================================================
' 概要:
'   _Configシート（設定管理）の基本機能
'   セキュリティ要件に従い自動イベントは禁止
'=============================================================================
Option Explicit

'=============================================================================
' セキュリティ制限により以下のイベントは実装禁止
' - Worksheet_Change
' - Worksheet_SelectionChange
' - その他のワークシートイベント
'=============================================================================
