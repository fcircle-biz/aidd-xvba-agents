Attribute VB_Name = "modConstants"
Option Explicit

' =============================================================================
' ���i�ݏo�Ǘ��V�X�e�� - �V�X�e���萔��`
' =============================================================================

' �V�[�g���萔
Public Const SHEET_DASHBOARD As String = "Dashboard"
Public Const SHEET_ITEMS As String = "Items"
Public Const SHEET_LENDING As String = "Lending" 
Public Const SHEET_INPUT As String = "Input"
Public Const SHEET_CONFIG As String = "_Config"

' �e�[�u�����萔
Public Const TABLE_ITEMS As String = "ItemsTable"
Public Const TABLE_LENDING As String = "LendingTable"

' �񖼒萔 - Items�i���i�}�X�^�j
Public Const COL_ITEM_ID As String = "ItemID"
Public Const COL_ITEM_NAME As String = "ItemName"
Public Const COL_CATEGORY As String = "Category"
Public Const COL_LOCATION As String = "Location"
Public Const COL_QUANTITY As String = "Quantity"

' �񖼒萔 - Lending�i�ݏo�����j
Public Const COL_RECORD_ID As String = "RecordID"
Public Const COL_LENDING_ITEM_ID As String = "ItemID"
Public Const COL_LENDING_ITEM_NAME As String = "ItemName"
Public Const COL_BORROWER As String = "Borrower"
Public Const COL_LEND_DATE As String = "LendDate"
Public Const COL_DUE_DATE As String = "DueDate"
Public Const COL_RETURN_DATE As String = "ReturnDate"
Public Const COL_STATUS As String = "Status"
Public Const COL_REMARKS As String = "Remarks"

' �X�e�[�^�X�l�萔
Public Const STATUS_LENDING As String = "�ݏo��"
Public Const STATUS_RETURNED As String = "�ԋp��"

' �F�萔�iRGB�l�j
Public Const COLOR_HEADER As Long = 12632256        ' �����i�w�b�_�[�p�j
Public Const COLOR_OVERDUE As Long = 16711680       ' �ԁi�������߁j
Public Const COLOR_WARNING As Long = 65535          ' ���F�i�����ԋ߁j
Public Const COLOR_NORMAL As Long = 16777215        ' ���i�ʏ�j
Public Const COLOR_SUCCESS As Long = 65280          ' �΁i�����E�����j
Public Const COLOR_ALTERNATE As Long = 15790320     ' �����O���[�i���ݍs�j

' �ݒ�l�萔
Public Const DEFAULT_LENDING_DAYS As Long = 7       ' �f�t�H���g�ݏo���ԁi���j
Public Const WARNING_DAYS_BEFORE As Long = 1        ' ���������O�Ɍx�����o����
Public Const MAX_LENDING_DAYS As Long = 30          ' �ő�ݏo���ԁi���j

' �Z���͈͒萔�iDashboard�p�j
Public Const RANGE_TOTAL_ITEMS As String = "B2"
Public Const RANGE_LENDING_COUNT As String = "B3"
Public Const RANGE_OVERDUE_COUNT As String = "B4"
Public Const RANGE_AVAILABLE_COUNT As String = "B5"

' ���̓V�[�g �Z���͈͒萔
Public Const INPUT_ITEM_ID As String = "B2"
Public Const INPUT_BORROWER As String = "B3"
Public Const INPUT_LEND_DATE As String = "B4"
Public Const INPUT_LENDING_DAYS As String = "B5"
Public Const INPUT_RETURN_DATE As String = "B6"

' �J�e�S���I����
Public Const CATEGORY_PC As String = "PC�E�m�[�gPC"
Public Const CATEGORY_AV As String = "AV�@��"
Public Const CATEGORY_STATIONERY As String = "����E�����p�i"
Public Const CATEGORY_TOOL As String = "�H��E�v����"
Public Const CATEGORY_OTHER As String = "���̑�"

' �ۊǏꏊ�I����
Public Const LOCATION_OFFICE_1F As String = "������1F"
Public Const LOCATION_OFFICE_2F As String = "������2F"
Public Const LOCATION_WAREHOUSE As String = "�q��"
Public Const LOCATION_MEETING_ROOM As String = "��c��"

' �G���[���b�Z�[�W�萔
Public Const MSG_ITEM_NOT_FOUND As String = "�w�肳�ꂽ���iID��������܂���B"
Public Const MSG_ALREADY_RETURNED As String = "���̔��i�͊��ɕԋp�ς݂ł��B"
Public Const MSG_NO_LENDING_RECORD As String = "�Y������ݏo�L�^��������܂���B"
Public Const MSG_INSUFFICIENT_STOCK As String = "�݌ɂ��s�����Ă��܂��B"
Public Const MSG_REQUIRED_FIELD As String = "�K�{���ڂ����͂���Ă��܂���B"
Public Const MSG_INVALID_DATE As String = "���t�̌`��������������܂���B"
Public Const MSG_INVALID_ITEM_ID As String = "���iID�͐��l�œ��͂��Ă��������B"

' ���O���x���萔
Public Const LOG_LEVEL_INFO As String = "INFO"
Public Const LOG_LEVEL_WARNING As String = "WARNING"
Public Const LOG_LEVEL_ERROR As String = "ERROR"
Public Const LOG_LEVEL_AUDIT As String = "AUDIT"