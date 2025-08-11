Attribute VB_Name = "modConstants"
'=============================================================================
' modConstants.bas - システム定数定義モジュール
'=============================================================================
' 概要:
'   顧客データ管理システムで使用するすべての定数を管理
'   シート名、インデックス、テーブル名、フォルダパス、ファイル名パターンなど
'=============================================================================
Option Explicit

'=============================================================================
' シート構成定数
'=============================================================================

' シート名定数（リネーム用）
Public Const SHEET_DASHBOARD As String = "Dashboard"
Public Const SHEET_CUSTOMERS As String = "Customers"
Public Const SHEET_STAGING As String = "Staging"
Public Const SHEET_CONFIG As String = "_Config"
Public Const SHEET_LOGS As String = "Logs"
Public Const SHEET_CODEBOOK As String = "Codebook"

' シートインデックス定数（GetWorksheetByIndex用）
Public Const SHEET_INDEX_DASHBOARD As Integer = 1
Public Const SHEET_INDEX_CUSTOMERS As Integer = 2
Public Const SHEET_INDEX_STAGING As Integer = 3
Public Const SHEET_INDEX_CONFIG As Integer = 4
Public Const SHEET_INDEX_LOGS As Integer = 5
Public Const SHEET_INDEX_CODEBOOK As Integer = 6

'=============================================================================
' テーブル名定数
'=============================================================================
Public Const TABLE_CUSTOMERS As String = "tblCustomers"
Public Const TABLE_STAGING As String = "tblStaging"
Public Const TABLE_CONFIG As String = "tblConfig"
Public Const TABLE_LOGS As String = "tblLogs"
Public Const TABLE_CODEBOOK As String = "tblCodebook"

'=============================================================================
' 列名定数（Customersテーブル）
'=============================================================================
Public Const COL_CUSTOMER_ID As String = "CustomerID"
Public Const COL_CUSTOMER_NAME As String = "CustomerName"
Public Const COL_EMAIL As String = "Email"
Public Const COL_PHONE As String = "Phone"
Public Const COL_ZIP As String = "Zip"
Public Const COL_ADDRESS1 As String = "Address1"
Public Const COL_ADDRESS2 As String = "Address2"
Public Const COL_CATEGORY As String = "Category"
Public Const COL_STATUS As String = "Status"
Public Const COL_CREATED_AT As String = "CreatedAt"
Public Const COL_UPDATED_AT As String = "UpdatedAt"
Public Const COL_SOURCE_FILE As String = "SourceFile"
Public Const COL_NOTES As String = "Notes"

'=============================================================================
' Staging専用列名定数（正規化・検証用）
'=============================================================================
Public Const COL_EMAIL_NORM As String = "Email_norm"
Public Const COL_PHONE_NORM As String = "Phone_norm"
Public Const COL_ZIP_NORM As String = "Zip_norm"
Public Const COL_KEY_CANDIDATE As String = "Key_Candidate"
Public Const COL_IS_VALID As String = "IsValid"
Public Const COL_ERROR_MESSAGE As String = "ErrorMessage"

'=============================================================================
' 設定キー定数（_Configシート用）
'=============================================================================
Public Const CONFIG_CSV_DIR As String = "CSV_DIR"
Public Const CONFIG_CSV_FILE As String = "CSV_FILE"
Public Const CONFIG_PRIMARY_KEY As String = "PRIMARY_KEY"
Public Const CONFIG_ALT_KEY As String = "ALT_KEY"
Public Const CONFIG_REQUIRED As String = "REQUIRED"
Public Const CONFIG_INACTIVATE_DAYS As String = "INACTIVATE_DAYS"
Public Const CONFIG_EMAIL_REGEX As String = "EMAIL_REGEX"
Public Const CONFIG_ZIP_REGEX As String = "ZIP_REGEX"
Public Const CONFIG_PHONE_REGEX As String = "PHONE_REGEX"
Public Const CONFIG_BACKUP_ENABLED As String = "BACKUP_ENABLED"
Public Const CONFIG_BACKUP_DIR As String = "BACKUP_DIR"

'=============================================================================
' デフォルト設定値定数
'=============================================================================
Public Const DEFAULT_CSV_DIR As String = "C:\Data\Import\"
Public Const DEFAULT_CSV_FILE As String = "customers_*.csv"
Public Const DEFAULT_BACKUP_DIR As String = "C:\Data\Backup\"
Public Const DEFAULT_PRIMARY_KEY As String = "CustomerID"
Public Const DEFAULT_ALT_KEY As String = "Email+CustomerName"
Public Const DEFAULT_REQUIRED As String = "CustomerID,CustomerName,Status"
Public Const DEFAULT_INACTIVATE_DAYS As Integer = 180

'=============================================================================
' ステータス値定数
'=============================================================================
Public Const STATUS_ACTIVE As String = "有効"
Public Const STATUS_INACTIVE As String = "無効"
Public Const STATUS_SUSPENDED As String = "停止"

'=============================================================================
' カテゴリ値定数
'=============================================================================
Public Const CATEGORY_B2B As String = "B2B"
Public Const CATEGORY_B2C As String = "B2C"
Public Const CATEGORY_PARTNER As String = "代理店"
Public Const CATEGORY_RESELLER As String = "パートナー"

'=============================================================================
' ログレベル定数
'=============================================================================
Public Const LOG_LEVEL_INFO As String = "INFO"
Public Const LOG_LEVEL_WARN As String = "WARN"
Public Const LOG_LEVEL_ERROR As String = "ERROR"
Public Const LOG_LEVEL_DEBUG As String = "DEBUG"

'=============================================================================
' エラーメッセージ定数
'=============================================================================
Public Const ERR_CSV_NOT_FOUND As String = "CSVファイルが見つかりません"
Public Const ERR_INVALID_CSV_FORMAT As String = "CSVファイルの形式が不正です"
Public Const ERR_DUPLICATE_CUSTOMER As String = "重複する顧客データが存在します"
Public Const ERR_REQUIRED_FIELD_MISSING As String = "必須フィールドが不足しています"
Public Const ERR_INVALID_EMAIL_FORMAT As String = "メールアドレスの形式が不正です"
Public Const ERR_INVALID_PHONE_FORMAT As String = "電話番号の形式が不正です"
Public Const ERR_INVALID_ZIP_FORMAT As String = "郵便番号の形式が不正です"
Public Const ERR_CUSTOMER_NOT_FOUND As String = "顧客データが見つかりません"
Public Const ERR_UPDATE_FAILED As String = "データの更新に失敗しました"
Public Const ERR_DELETE_FAILED As String = "データの削除に失敗しました"

'=============================================================================
' UI関連定数
'=============================================================================
Public Const BTN_IMPORT_CSV As String = "btnImportCsv"
Public Const BTN_CLEAR_STAGING As String = "btnClearStaging"
Public Const BTN_EXPORT_REPORT As String = "btnExportReport"
Public Const BTN_OPEN_CONFIG As String = "btnOpenConfig"
Public Const BTN_REFRESH_KPI As String = "btnRefreshKpi"

' KPI表示位置定数（Dashboardシート）
Public Const KPI_TOTAL_CUSTOMERS_CELL As String = "D5"
Public Const KPI_ADDED_COUNT_CELL As String = "D6"
Public Const KPI_UPDATED_COUNT_CELL As String = "D7"
Public Const KPI_DUPLICATE_COUNT_CELL As String = "D8"
Public Const KPI_ERROR_COUNT_CELL As String = "D9"
Public Const KPI_INACTIVE_COUNT_CELL As String = "D10"
Public Const KPI_LAST_IMPORT_CELL As String = "D11"
Public Const KPI_PROCESS_TIME_CELL As String = "D12"

'=============================================================================
' フォーマット・スタイル定数
'=============================================================================
Public Const DATE_FORMAT_DISPLAY As String = "yyyy/mm/dd hh:mm"
Public Const DATE_FORMAT_FILE As String = "yyyymmdd_hhnnss"
Public Const NUMBER_FORMAT_COUNT As String = "#,##0"

'=============================================================================
' CSV処理関連定数
'=============================================================================
Public Const CSV_ENCODING_UTF8 As String = "UTF-8"
Public Const CSV_DELIMITER As String = ","
Public Const CSV_QUOTE_CHAR As String = """"
Public Const CSV_MAX_FIELD_COUNT As Integer = 20
Public Const CSV_MAX_RECORDS As Long = 100000

'=============================================================================
' バックアップ・ログ関連定数
'=============================================================================
Public Const BACKUP_FILE_PREFIX As String = "customers_backup_"
Public Const LOG_FILE_PREFIX As String = "customer_system_"
Public Const MAX_LOG_FILE_SIZE As Long = 10485760  ' 10MB
Public Const MAX_BACKUP_FILES As Integer = 30

'=============================================================================
' パフォーマンス関連定数
'=============================================================================
Public Const BATCH_SIZE_CSV_IMPORT As Integer = 1000
Public Const BATCH_SIZE_VALIDATION As Integer = 500
Public Const BATCH_SIZE_UPSERT As Integer = 200

'=============================================================================
' 正規表現パターン定数（詳細版）
'=============================================================================
Public Const REGEX_EMAIL_STRICT As String = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
Public Const REGEX_PHONE_JAPAN As String = "^0[0-9]{1,4}-[0-9]{1,4}-[0-9]{4}$"
Public Const REGEX_ZIP_JAPAN As String = "^\d{3}-\d{4}$"
Public Const REGEX_CUSTOMER_ID_STRICT As String = "^[A-Za-z0-9]{3,20}$"

'=============================================================================
' テーブル作成用ヘッダー定数
'=============================================================================
Public Const CUSTOMERS_HEADERS As String = "CustomerID,CustomerName,Email,Phone,Zip,Address1,Address2,Category,Status,CreatedAt,UpdatedAt,SourceFile,Notes"
Public Const STAGING_HEADERS As String = "CustomerID,CustomerName,Email,Phone,Zip,Address1,Address2,Category,Status,Email_norm,Phone_norm,Zip_norm,Key_Candidate,IsValid,ErrorMessage,SourceFile"
Public Const CONFIG_HEADERS As String = "ConfigKey,ConfigValue,Description"
Public Const LOGS_HEADERS As String = "Timestamp,Level,Module,Message,Details,RecordCount,ProcessTime,SourceFile"
Public Const CODEBOOK_HEADERS As String = "ExternalColumnName,InternalColumnName,DataType,ValidationRule,NormalizationRule,Required,Description"

'=============================================================================
' メッセージ定数
'=============================================================================
Public Const MSG_IMPORT_STARTED As String = "CSV取り込み処理を開始しています..."
Public Const MSG_IMPORT_COMPLETED As String = "CSV取り込み処理が完了しました"
Public Const MSG_VALIDATION_STARTED As String = "データ検証を実行しています..."
Public Const MSG_VALIDATION_COMPLETED As String = "データ検証が完了しました"
Public Const MSG_UPSERT_STARTED As String = "データ更新処理を開始しています..."
Public Const MSG_UPSERT_COMPLETED As String = "データ更新処理が完了しました"
Public Const MSG_CLEANUP_STARTED As String = "クリーンアップ処理を開始しています..."
Public Const MSG_CLEANUP_COMPLETED As String = "クリーンアップ処理が完了しました"
Public Const MSG_CONFIRM_CLEAR_STAGING As String = "Stagingデータをクリアしますか？"
Public Const MSG_CONFIRM_INACTIVATE As String = "期限切れ顧客データを無効化しますか？"

'=============================================================================
' システム情報定数
'=============================================================================
Public Const SYSTEM_NAME As String = "顧客情報取込＆整形VBAシステム"
Public Const SYSTEM_VERSION As String = "1.0.0"
Public Const SYSTEM_AUTHOR As String = "XVBA Mock Creator"
Public Const SYSTEM_COPYRIGHT As String = "c 2025 XVBA Development Team"