Attribute VB_Name = "modConstants"
'=============================================================================
' modConstants.bas - �V�X�e���萔��`���W���[��
'=============================================================================
' �T�v:
'   �ڋq�f�[�^�Ǘ��V�X�e���Ŏg�p���邷�ׂĂ̒萔���Ǘ�
'   �V�[�g���A�C���f�b�N�X�A�e�[�u�����A�t�H���_�p�X�A�t�@�C�����p�^�[���Ȃ�
'=============================================================================
Option Explicit

'=============================================================================
' �V�[�g�\���萔
'=============================================================================

' �V�[�g���萔�i���l�[���p�j
Public Const SHEET_DASHBOARD As String = "Dashboard"
Public Const SHEET_CUSTOMERS As String = "Customers"
Public Const SHEET_STAGING As String = "Staging"
Public Const SHEET_CONFIG As String = "_Config"
Public Const SHEET_LOGS As String = "Logs"
Public Const SHEET_CODEBOOK As String = "Codebook"

' �V�[�g�C���f�b�N�X�萔�iGetWorksheetByIndex�p�j
Public Const SHEET_INDEX_DASHBOARD As Integer = 1
Public Const SHEET_INDEX_CUSTOMERS As Integer = 2
Public Const SHEET_INDEX_STAGING As Integer = 3
Public Const SHEET_INDEX_CONFIG As Integer = 4
Public Const SHEET_INDEX_LOGS As Integer = 5
Public Const SHEET_INDEX_CODEBOOK As Integer = 6

'=============================================================================
' �e�[�u�����萔
'=============================================================================
Public Const TABLE_CUSTOMERS As String = "tblCustomers"
Public Const TABLE_STAGING As String = "tblStaging"
Public Const TABLE_CONFIG As String = "tblConfig"
Public Const TABLE_LOGS As String = "tblLogs"
Public Const TABLE_CODEBOOK As String = "tblCodebook"

'=============================================================================
' �񖼒萔�iCustomers�e�[�u���j
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
' Staging��p�񖼒萔�i���K���E���ؗp�j
'=============================================================================
Public Const COL_EMAIL_NORM As String = "Email_norm"
Public Const COL_PHONE_NORM As String = "Phone_norm"
Public Const COL_ZIP_NORM As String = "Zip_norm"
Public Const COL_KEY_CANDIDATE As String = "Key_Candidate"
Public Const COL_IS_VALID As String = "IsValid"
Public Const COL_ERROR_MESSAGE As String = "ErrorMessage"

'=============================================================================
' �ݒ�L�[�萔�i_Config�V�[�g�p�j
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
' �f�t�H���g�ݒ�l�萔
'=============================================================================
Public Const DEFAULT_CSV_DIR As String = "C:\Data\Import\"
Public Const DEFAULT_CSV_FILE As String = "customers_*.csv"
Public Const DEFAULT_BACKUP_DIR As String = "C:\Data\Backup\"
Public Const DEFAULT_PRIMARY_KEY As String = "CustomerID"
Public Const DEFAULT_ALT_KEY As String = "Email+CustomerName"
Public Const DEFAULT_REQUIRED As String = "CustomerID,CustomerName,Status"
Public Const DEFAULT_INACTIVATE_DAYS As Integer = 180

'=============================================================================
' �X�e�[�^�X�l�萔
'=============================================================================
Public Const STATUS_ACTIVE As String = "�L��"
Public Const STATUS_INACTIVE As String = "����"
Public Const STATUS_SUSPENDED As String = "��~"

'=============================================================================
' �J�e�S���l�萔
'=============================================================================
Public Const CATEGORY_B2B As String = "B2B"
Public Const CATEGORY_B2C As String = "B2C"
Public Const CATEGORY_PARTNER As String = "�㗝�X"
Public Const CATEGORY_RESELLER As String = "�p�[�g�i�["

'=============================================================================
' ���O���x���萔
'=============================================================================
Public Const LOG_LEVEL_INFO As String = "INFO"
Public Const LOG_LEVEL_WARN As String = "WARN"
Public Const LOG_LEVEL_ERROR As String = "ERROR"
Public Const LOG_LEVEL_DEBUG As String = "DEBUG"

'=============================================================================
' �G���[���b�Z�[�W�萔
'=============================================================================
Public Const ERR_CSV_NOT_FOUND As String = "CSV�t�@�C����������܂���"
Public Const ERR_INVALID_CSV_FORMAT As String = "CSV�t�@�C���̌`�����s���ł�"
Public Const ERR_DUPLICATE_CUSTOMER As String = "�d������ڋq�f�[�^�����݂��܂�"
Public Const ERR_REQUIRED_FIELD_MISSING As String = "�K�{�t�B�[���h���s�����Ă��܂�"
Public Const ERR_INVALID_EMAIL_FORMAT As String = "���[���A�h���X�̌`�����s���ł�"
Public Const ERR_INVALID_PHONE_FORMAT As String = "�d�b�ԍ��̌`�����s���ł�"
Public Const ERR_INVALID_ZIP_FORMAT As String = "�X�֔ԍ��̌`�����s���ł�"
Public Const ERR_CUSTOMER_NOT_FOUND As String = "�ڋq�f�[�^��������܂���"
Public Const ERR_UPDATE_FAILED As String = "�f�[�^�̍X�V�Ɏ��s���܂���"
Public Const ERR_DELETE_FAILED As String = "�f�[�^�̍폜�Ɏ��s���܂���"

'=============================================================================
' UI�֘A�萔
'=============================================================================
Public Const BTN_IMPORT_CSV As String = "btnImportCsv"
Public Const BTN_CLEAR_STAGING As String = "btnClearStaging"
Public Const BTN_EXPORT_REPORT As String = "btnExportReport"
Public Const BTN_OPEN_CONFIG As String = "btnOpenConfig"
Public Const BTN_REFRESH_KPI As String = "btnRefreshKpi"

' KPI�\���ʒu�萔�iDashboard�V�[�g�j
Public Const KPI_TOTAL_CUSTOMERS_CELL As String = "D5"
Public Const KPI_ADDED_COUNT_CELL As String = "D6"
Public Const KPI_UPDATED_COUNT_CELL As String = "D7"
Public Const KPI_DUPLICATE_COUNT_CELL As String = "D8"
Public Const KPI_ERROR_COUNT_CELL As String = "D9"
Public Const KPI_INACTIVE_COUNT_CELL As String = "D10"
Public Const KPI_LAST_IMPORT_CELL As String = "D11"
Public Const KPI_PROCESS_TIME_CELL As String = "D12"

'=============================================================================
' �t�H�[�}�b�g�E�X�^�C���萔
'=============================================================================
Public Const DATE_FORMAT_DISPLAY As String = "yyyy/mm/dd hh:mm"
Public Const DATE_FORMAT_FILE As String = "yyyymmdd_hhnnss"
Public Const NUMBER_FORMAT_COUNT As String = "#,##0"

'=============================================================================
' CSV�����֘A�萔
'=============================================================================
Public Const CSV_ENCODING_UTF8 As String = "UTF-8"
Public Const CSV_DELIMITER As String = ","
Public Const CSV_QUOTE_CHAR As String = """"
Public Const CSV_MAX_FIELD_COUNT As Integer = 20
Public Const CSV_MAX_RECORDS As Long = 100000

'=============================================================================
' �o�b�N�A�b�v�E���O�֘A�萔
'=============================================================================
Public Const BACKUP_FILE_PREFIX As String = "customers_backup_"
Public Const LOG_FILE_PREFIX As String = "customer_system_"
Public Const MAX_LOG_FILE_SIZE As Long = 10485760  ' 10MB
Public Const MAX_BACKUP_FILES As Integer = 30

'=============================================================================
' �p�t�H�[�}���X�֘A�萔
'=============================================================================
Public Const BATCH_SIZE_CSV_IMPORT As Integer = 1000
Public Const BATCH_SIZE_VALIDATION As Integer = 500
Public Const BATCH_SIZE_UPSERT As Integer = 200

'=============================================================================
' ���K�\���p�^�[���萔�i�ڍהŁj
'=============================================================================
Public Const REGEX_EMAIL_STRICT As String = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
Public Const REGEX_PHONE_JAPAN As String = "^0[0-9]{1,4}-[0-9]{1,4}-[0-9]{4}$"
Public Const REGEX_ZIP_JAPAN As String = "^\d{3}-\d{4}$"
Public Const REGEX_CUSTOMER_ID_STRICT As String = "^[A-Za-z0-9]{3,20}$"

'=============================================================================
' �e�[�u���쐬�p�w�b�_�[�萔
'=============================================================================
Public Const CUSTOMERS_HEADERS As String = "CustomerID,CustomerName,Email,Phone,Zip,Address1,Address2,Category,Status,CreatedAt,UpdatedAt,SourceFile,Notes"
Public Const STAGING_HEADERS As String = "CustomerID,CustomerName,Email,Phone,Zip,Address1,Address2,Category,Status,Email_norm,Phone_norm,Zip_norm,Key_Candidate,IsValid,ErrorMessage,SourceFile"
Public Const CONFIG_HEADERS As String = "ConfigKey,ConfigValue,Description"
Public Const LOGS_HEADERS As String = "Timestamp,Level,Module,Message,Details,RecordCount,ProcessTime,SourceFile"
Public Const CODEBOOK_HEADERS As String = "ExternalColumnName,InternalColumnName,DataType,ValidationRule,NormalizationRule,Required,Description"

'=============================================================================
' ���b�Z�[�W�萔
'=============================================================================
Public Const MSG_IMPORT_STARTED As String = "CSV��荞�ݏ������J�n���Ă��܂�..."
Public Const MSG_IMPORT_COMPLETED As String = "CSV��荞�ݏ������������܂���"
Public Const MSG_VALIDATION_STARTED As String = "�f�[�^���؂����s���Ă��܂�..."
Public Const MSG_VALIDATION_COMPLETED As String = "�f�[�^���؂��������܂���"
Public Const MSG_UPSERT_STARTED As String = "�f�[�^�X�V�������J�n���Ă��܂�..."
Public Const MSG_UPSERT_COMPLETED As String = "�f�[�^�X�V�������������܂���"
Public Const MSG_CLEANUP_STARTED As String = "�N���[���A�b�v�������J�n���Ă��܂�..."
Public Const MSG_CLEANUP_COMPLETED As String = "�N���[���A�b�v�������������܂���"
Public Const MSG_CONFIRM_CLEAR_STAGING As String = "Staging�f�[�^���N���A���܂����H"
Public Const MSG_CONFIRM_INACTIVATE As String = "�����؂�ڋq�f�[�^�𖳌������܂����H"

'=============================================================================
' �V�X�e�����萔
'=============================================================================
Public Const SYSTEM_NAME As String = "�ڋq���捞�����`VBA�V�X�e��"
Public Const SYSTEM_VERSION As String = "1.0.0"
Public Const SYSTEM_AUTHOR As String = "XVBA Mock Creator"
Public Const SYSTEM_COPYRIGHT As String = "c 2025 XVBA Development Team"