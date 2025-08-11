Attribute VB_Name = "modConstants"
Option Explicit

' =============================================================================
' 備品貸出管理システム - システム定数定義
' =============================================================================

' シート名定数
Public Const SHEET_DASHBOARD As String = "Dashboard"
Public Const SHEET_ITEMS As String = "Items"
Public Const SHEET_LENDING As String = "Lending" 
Public Const SHEET_INPUT As String = "Input"
Public Const SHEET_CONFIG As String = "_Config"

' テーブル名定数
Public Const TABLE_ITEMS As String = "ItemsTable"
Public Const TABLE_LENDING As String = "LendingTable"

' 列名定数 - Items（備品マスタ）
Public Const COL_ITEM_ID As String = "ItemID"
Public Const COL_ITEM_NAME As String = "ItemName"
Public Const COL_CATEGORY As String = "Category"
Public Const COL_LOCATION As String = "Location"
Public Const COL_QUANTITY As String = "Quantity"

' 列名定数 - Lending（貸出履歴）
Public Const COL_RECORD_ID As String = "RecordID"
Public Const COL_LENDING_ITEM_ID As String = "ItemID"
Public Const COL_LENDING_ITEM_NAME As String = "ItemName"
Public Const COL_BORROWER As String = "Borrower"
Public Const COL_LEND_DATE As String = "LendDate"
Public Const COL_DUE_DATE As String = "DueDate"
Public Const COL_RETURN_DATE As String = "ReturnDate"
Public Const COL_STATUS As String = "Status"
Public Const COL_REMARKS As String = "Remarks"

' ステータス値定数
Public Const STATUS_LENDING As String = "貸出中"
Public Const STATUS_RETURNED As String = "返却済"

' 色定数（RGB値）
Public Const COLOR_HEADER As Long = 12632256        ' 薄い青（ヘッダー用）
Public Const COLOR_OVERDUE As Long = 16711680       ' 赤（期限超過）
Public Const COLOR_WARNING As Long = 65535          ' 黄色（期限間近）
Public Const COLOR_NORMAL As Long = 16777215        ' 白（通常）
Public Const COLOR_SUCCESS As Long = 65280          ' 緑（成功・完了）
Public Const COLOR_ALTERNATE As Long = 15790320     ' 薄いグレー（交互行）

' 設定値定数
Public Const DEFAULT_LENDING_DAYS As Long = 7       ' デフォルト貸出期間（日）
Public Const WARNING_DAYS_BEFORE As Long = 1        ' 期限何日前に警告を出すか
Public Const MAX_LENDING_DAYS As Long = 30          ' 最大貸出期間（日）

' セル範囲定数（Dashboard用）
Public Const RANGE_TOTAL_ITEMS As String = "B2"
Public Const RANGE_LENDING_COUNT As String = "B3"
Public Const RANGE_OVERDUE_COUNT As String = "B4"
Public Const RANGE_AVAILABLE_COUNT As String = "B5"

' 入力シート セル範囲定数
Public Const INPUT_ITEM_ID As String = "B2"
Public Const INPUT_BORROWER As String = "B3"
Public Const INPUT_LEND_DATE As String = "B4"
Public Const INPUT_LENDING_DAYS As String = "B5"
Public Const INPUT_RETURN_DATE As String = "B6"

' カテゴリ選択肢
Public Const CATEGORY_PC As String = "PC・ノートPC"
Public Const CATEGORY_AV As String = "AV機器"
Public Const CATEGORY_STATIONERY As String = "文具・事務用品"
Public Const CATEGORY_TOOL As String = "工具・計測器"
Public Const CATEGORY_OTHER As String = "その他"

' 保管場所選択肢
Public Const LOCATION_OFFICE_1F As String = "事務所1F"
Public Const LOCATION_OFFICE_2F As String = "事務所2F"
Public Const LOCATION_WAREHOUSE As String = "倉庫"
Public Const LOCATION_MEETING_ROOM As String = "会議室"

' エラーメッセージ定数
Public Const MSG_ITEM_NOT_FOUND As String = "指定された備品IDが見つかりません。"
Public Const MSG_ALREADY_RETURNED As String = "この備品は既に返却済みです。"
Public Const MSG_NO_LENDING_RECORD As String = "該当する貸出記録が見つかりません。"
Public Const MSG_INSUFFICIENT_STOCK As String = "在庫が不足しています。"
Public Const MSG_REQUIRED_FIELD As String = "必須項目が入力されていません。"
Public Const MSG_INVALID_DATE As String = "日付の形式が正しくありません。"
Public Const MSG_INVALID_ITEM_ID As String = "備品IDは数値で入力してください。"

' ログレベル定数
Public Const LOG_LEVEL_INFO As String = "INFO"
Public Const LOG_LEVEL_WARNING As String = "WARNING"
Public Const LOG_LEVEL_ERROR As String = "ERROR"
Public Const LOG_LEVEL_AUDIT As String = "AUDIT"