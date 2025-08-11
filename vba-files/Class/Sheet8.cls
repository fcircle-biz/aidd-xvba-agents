VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'=============================================================================
' Sheet8.cls - 予備シートクラス
'=============================================================================
' 概要:
'   予備シート（将来拡張用）
'   セキュリティ要件に従い自動イベントは禁止
'=============================================================================
Option Explicit

'=============================================================================
' セキュリティ制限により以下のイベントは実装禁止
' - Worksheet_Change
' - Worksheet_SelectionChange
' - その他のワークシートイベント
'=============================================================================
