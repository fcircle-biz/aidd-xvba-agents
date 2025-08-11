Attribute VB_Name = "modTestData"
Option Explicit

' =============================================================================
' 備品貸出管理システム - テストデータ生成モジュール
' =============================================================================

' 全テストデータ作成メイン処理
Public Sub CreateAllTestData()
    On Error GoTo ErrHandler
    
    Dim result As VbMsgBoxResult
    result = MsgBox("テストデータを作成します。既存のデータは削除されます。続行しますか？", vbQuestion + vbYesNo)
    
    If result = vbNo Then Exit Sub
    
    Application.ScreenUpdating = False
    
    ' 既存データクリア
    Call ClearAllData
    
    ' 備品マスタテストデータ作成
    Call CreateItemsTestData
    
    ' 貸出履歴テストデータ作成
    Call CreateLendingTestData
    
    ' ダッシュボード更新
    Call UpdateDashboard
    
    Application.ScreenUpdating = True
    
    ' 監査ログ記録
    Call LogAudit("テストデータ作成", "All test data created successfully")
    
    MsgBox "テストデータの作成が完了しました。", vbInformation
    
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Call LogError("CreateAllTestData", Err.Number, Err.Description)
    MsgBox "テストデータ作成中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' 既存データクリア
Private Sub ClearAllData()
    On Error Resume Next
    
    ' 備品テーブルクリア
    Dim itemsTbl As ListObject
    Set itemsTbl = GetItemsTable()
    If Not itemsTbl Is Nothing Then
        If Not itemsTbl.DataBodyRange Is Nothing Then
            itemsTbl.DataBodyRange.Delete
        End If
    End If
    
    ' 貸出テーブルクリア
    Dim lendingTbl As ListObject
    Set lendingTbl = GetLendingTable()
    If Not lendingTbl Is Nothing Then
        If Not lendingTbl.DataBodyRange Is Nothing Then
            lendingTbl.DataBodyRange.Delete
        End If
    End If
    
    On Error GoTo 0
End Sub

' 備品マスタテストデータ作成
Private Sub CreateItemsTestData()
    On Error GoTo ErrHandler
    
    Dim tbl As ListObject
    Set tbl = GetItemsTable()
    
    If tbl Is Nothing Then
        Call LogError("CreateItemsTestData", 9, "Items table not found")
        Exit Sub
    End If
    
    ' テストデータ配列
    Dim testItems As Variant
    testItems = Array( _
        Array(1001, "ノートPC - ThinkPad", CATEGORY_PC, LOCATION_OFFICE_1F, 5), _
        Array(1002, "ノートPC - MacBook Air", CATEGORY_PC, LOCATION_OFFICE_1F, 3), _
        Array(1003, "デスクトップPC - iMac", CATEGORY_PC, LOCATION_OFFICE_2F, 2), _
        Array(2001, "プロジェクター - EPSON", CATEGORY_AV, LOCATION_MEETING_ROOM, 4), _
        Array(2002, "モニター 24inch", CATEGORY_AV, LOCATION_OFFICE_2F, 6), _
        Array(2003, "Webカメラ - Logitech", CATEGORY_AV, LOCATION_OFFICE_1F, 8), _
        Array(3001, "電卓 - CASIO", CATEGORY_STATIONERY, LOCATION_OFFICE_1F, 10), _
        Array(3002, "USBメモリ 32GB", CATEGORY_STATIONERY, LOCATION_WAREHOUSE, 15), _
        Array(3003, "マウス - ワイヤレス", CATEGORY_STATIONERY, LOCATION_OFFICE_1F, 12), _
        Array(4001, "テスター - デジタル", CATEGORY_TOOL, LOCATION_WAREHOUSE, 3), _
        Array(4002, "ドライバーセット", CATEGORY_TOOL, LOCATION_WAREHOUSE, 5), _
        Array(4003, "LAN ケーブルテスター", CATEGORY_TOOL, LOCATION_OFFICE_2F, 2), _
        Array(5001, "延長コード 10m", CATEGORY_OTHER, LOCATION_WAREHOUSE, 8), _
        Array(5002, "ホワイトボード用マーカー", CATEGORY_OTHER, LOCATION_MEETING_ROOM, 20), _
        Array(5003, "書類バインダー", CATEGORY_OTHER, LOCATION_OFFICE_1F, 25) _
    )
    
    ' データ追加
    Call AddItemsData(tbl, testItems)
    
    Exit Sub
    
ErrHandler:
    Call LogError("CreateItemsTestData", Err.Number, Err.Description)
End Sub

' 備品データ追加（内部関数）
Private Sub AddItemsData(tbl As ListObject, dataArray As Variant)
    On Error GoTo ErrHandler
    
    Dim i As Long
    For i = 0 To UBound(dataArray)
        Dim newRow As ListRow
        Set newRow = tbl.ListRows.Add()
        
        ' 列インデックス取得
        Dim itemIDCol As Long, itemNameCol As Long, categoryCol As Long
        Dim locationCol As Long, quantityCol As Long
        itemIDCol = GetColumnIndex(tbl, COL_ITEM_ID)
        itemNameCol = GetColumnIndex(tbl, COL_ITEM_NAME)
        categoryCol = GetColumnIndex(tbl, COL_CATEGORY)
        locationCol = GetColumnIndex(tbl, COL_LOCATION)
        quantityCol = GetColumnIndex(tbl, COL_QUANTITY)
        
        If itemIDCol > 0 Then newRow.Range.Cells(1, itemIDCol).Value = dataArray(i)(0)
        If itemNameCol > 0 Then newRow.Range.Cells(1, itemNameCol).Value = dataArray(i)(1)
        If categoryCol > 0 Then newRow.Range.Cells(1, categoryCol).Value = dataArray(i)(2)
        If locationCol > 0 Then newRow.Range.Cells(1, locationCol).Value = dataArray(i)(3)
        If quantityCol > 0 Then newRow.Range.Cells(1, quantityCol).Value = dataArray(i)(4)
    Next i
    
    ' テーブルフォーマット適用
    Call ApplyStandardTableFormat(tbl)
    
    Exit Sub
    
ErrHandler:
    Call LogError("AddItemsData", Err.Number, Err.Description)
End Sub

' 貸出履歴テストデータ作成
Private Sub CreateLendingTestData()
    On Error GoTo ErrHandler
    
    Dim tbl As ListObject
    Set tbl = GetLendingTable()
    
    If tbl Is Nothing Then
        Call LogError("CreateLendingTestData", 9, "Lending table not found")
        Exit Sub
    End If
    
    ' 現在の日付を基準にテストデータを作成
    Dim baseDate As Date
    baseDate = Date
    
    ' テストデータ配列（期限超過・期限間近・正常な貸出を含む）
    Dim testLendings As Variant
    testLendings = Array( _
        Array(1001, "田中太郎", baseDate - 10, baseDate - 3, "", STATUS_LENDING), _
        Array(2001, "佐藤花子", baseDate - 8, baseDate - 1, "", STATUS_LENDING), _
        Array(1002, "鈴木一郎", baseDate - 5, baseDate + 2, "", STATUS_LENDING), _
        Array(3001, "高橋美香", baseDate - 3, baseDate + 4, "", STATUS_LENDING), _
        Array(2002, "山田次郎", baseDate - 12, baseDate - 5, baseDate - 2, STATUS_RETURNED), _
        Array(4001, "渡辺三郎", baseDate - 7, baseDate, "", STATUS_LENDING), _
        Array(1003, "伊藤四郎", baseDate - 4, baseDate + 3, "", STATUS_LENDING), _
        Array(3002, "小林五郎", baseDate - 15, baseDate - 8, baseDate - 6, STATUS_RETURNED), _
        Array(2003, "加藤六郎", baseDate - 2, baseDate + 5, "", STATUS_LENDING), _
        Array(5001, "松本七子", baseDate - 9, baseDate - 2, baseDate - 1, STATUS_RETURNED) _
    )
    
    ' データ追加
    Call AddLendingData(tbl, testLendings)
    
    Exit Sub
    
ErrHandler:
    Call LogError("CreateLendingTestData", Err.Number, Err.Description)
End Sub

' 貸出データ追加（内部関数）
Private Sub AddLendingData(tbl As ListObject, dataArray As Variant)
    On Error GoTo ErrHandler
    
    Dim i As Long, recordID As Long
    recordID = 1
    
    For i = 0 To UBound(dataArray)
        Dim newRow As ListRow
        Set newRow = tbl.ListRows.Add()
        
        ' 列インデックス取得
        Dim recordIDCol As Long, itemIDCol As Long, itemNameCol As Long
        Dim borrowerCol As Long, lendDateCol As Long, dueDateCol As Long
        Dim returnDateCol As Long, statusCol As Long, remarksCol As Long
        
        recordIDCol = GetColumnIndex(tbl, COL_RECORD_ID)
        itemIDCol = GetColumnIndex(tbl, COL_LENDING_ITEM_ID)
        itemNameCol = GetColumnIndex(tbl, COL_LENDING_ITEM_NAME)
        borrowerCol = GetColumnIndex(tbl, COL_BORROWER)
        lendDateCol = GetColumnIndex(tbl, COL_LEND_DATE)
        dueDateCol = GetColumnIndex(tbl, COL_DUE_DATE)
        returnDateCol = GetColumnIndex(tbl, COL_RETURN_DATE)
        statusCol = GetColumnIndex(tbl, COL_STATUS)
        remarksCol = GetColumnIndex(tbl, COL_REMARKS)
        
        ' データ設定
        If recordIDCol > 0 Then newRow.Range.Cells(1, recordIDCol).Value = recordID
        ' 一時変数に値をコピーしてByRef問題を回避
        Dim tempItemID As Long
        tempItemID = dataArray(i)(0)
        
        If itemIDCol > 0 Then newRow.Range.Cells(1, itemIDCol).Value = tempItemID
        If itemNameCol > 0 Then
            ' 備品名を自動取得
            Dim itemName As String
            itemName = GetItemName(tempItemID)
            newRow.Range.Cells(1, itemNameCol).Value = itemName
        End If
        If borrowerCol > 0 Then newRow.Range.Cells(1, borrowerCol).Value = dataArray(i)(1)
        If lendDateCol > 0 Then newRow.Range.Cells(1, lendDateCol).Value = dataArray(i)(2)
        If dueDateCol > 0 Then newRow.Range.Cells(1, dueDateCol).Value = dataArray(i)(3)
        If returnDateCol > 0 Then newRow.Range.Cells(1, returnDateCol).Value = dataArray(i)(4)
        If statusCol > 0 Then newRow.Range.Cells(1, statusCol).Value = dataArray(i)(5)
        If remarksCol > 0 Then
            ' ステータスに応じて備考を自動設定（一時変数でByRef問題回避）
            Dim tempStatus As String, tempDueDate As Date
            tempStatus = dataArray(i)(5)
            tempDueDate = dataArray(i)(3)
            
            If tempStatus = STATUS_LENDING Then
                If tempDueDate < Date Then
                    newRow.Range.Cells(1, remarksCol).Value = "期限超過中"
                ElseIf tempDueDate <= Date + WARNING_DAYS_BEFORE Then
                    newRow.Range.Cells(1, remarksCol).Value = "期限間近"
                Else
                    newRow.Range.Cells(1, remarksCol).Value = "正常貸出中"
                End If
            Else
                newRow.Range.Cells(1, remarksCol).Value = "正常返却完了"
            End If
        End If
        
        recordID = recordID + 1
    Next i
    
    ' テーブルフォーマット適用
    Call ApplyStandardTableFormat(tbl)
    
    Exit Sub
    
ErrHandler:
    Call LogError("AddLendingData", Err.Number, Err.Description)
End Sub

' サンプル備品追加
Public Sub AddSampleItem()
    On Error GoTo ErrHandler
    
    Dim tbl As ListObject
    Set tbl = GetItemsTable()
    
    If tbl Is Nothing Then
        MsgBox "備品テーブルが見つかりません。", vbExclamation
        Exit Sub
    End If
    
    ' 次のIDを計算
    Dim nextID As Long
    nextID = 6001 ' サンプル用IDベース
    
    If Not tbl.DataBodyRange Is Nothing Then
        Dim itemIDCol As Long
        itemIDCol = GetColumnIndex(tbl, COL_ITEM_ID)
        If itemIDCol > 0 Then
            Dim i As Long, maxID As Long
            maxID = 0
            For i = 1 To tbl.DataBodyRange.Rows.Count
                If IsNumeric(tbl.DataBodyRange.Cells(i, itemIDCol).Value) Then
                    If tbl.DataBodyRange.Cells(i, itemIDCol).Value > maxID Then
                        maxID = tbl.DataBodyRange.Cells(i, itemIDCol).Value
                    End If
                End If
            Next i
            nextID = maxID + 1
        End If
    End If
    
    ' 新しい備品を追加
    Dim newRow As ListRow
    Set newRow = tbl.ListRows.Add()
    
    newRow.Range.Cells(1, GetColumnIndex(tbl, COL_ITEM_ID)).Value = nextID
    newRow.Range.Cells(1, GetColumnIndex(tbl, COL_ITEM_NAME)).Value = "サンプル備品 - " & Format(Now, "hhmmss")
    newRow.Range.Cells(1, GetColumnIndex(tbl, COL_CATEGORY)).Value = CATEGORY_OTHER
    newRow.Range.Cells(1, GetColumnIndex(tbl, COL_LOCATION)).Value = LOCATION_OFFICE_1F
    newRow.Range.Cells(1, GetColumnIndex(tbl, COL_QUANTITY)).Value = 1
    
    ' フォーマット適用
    Call ApplyStandardTableFormat(tbl)
    
    ' ダッシュボード更新
    Call UpdateDashboard
    
    ' ログ記録
    Call LogAudit("サンプル備品追加", "ItemID: " & nextID)
    
    MsgBox "サンプル備品（ID: " & nextID & "）を追加しました。", vbInformation
    
    Exit Sub
    
ErrHandler:
    Call LogError("AddSampleItem", Err.Number, Err.Description)
    MsgBox "サンプル備品追加中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' テスト用貸出作成
Public Sub CreateTestLending()
    On Error GoTo ErrHandler
    
    ' 利用可能な備品を検索
    Dim tbl As ListObject
    Set tbl = GetItemsTable()
    
    If tbl Is Nothing Then
        MsgBox "備品テーブルが見つかりません。", vbExclamation
        Exit Sub
    End If
    
    Dim availableItemID As Long
    availableItemID = 0
    
    ' 在庫のある備品を検索
    If Not tbl.DataBodyRange Is Nothing Then
        Dim itemIDCol As Long
        itemIDCol = GetColumnIndex(tbl, COL_ITEM_ID)
        If itemIDCol > 0 Then
            Dim i As Long
            For i = 1 To tbl.DataBodyRange.Rows.Count
                Dim itemID As Long
                itemID = tbl.DataBodyRange.Cells(i, itemIDCol).Value
                If GetAvailableQuantity(itemID) > 0 Then
                    availableItemID = itemID
                    Exit For
                End If
            Next i
        End If
    End If
    
    If availableItemID = 0 Then
        MsgBox "貸出可能な備品がありません。", vbExclamation
        Exit Sub
    End If
    
    ' 貸出記録を作成
    Dim lendingTbl As ListObject
    Set lendingTbl = GetLendingTable()
    
    If lendingTbl Is Nothing Then
        MsgBox "貸出テーブルが見つかりません。", vbExclamation
        Exit Sub
    End If
    
    Dim newRow As ListRow
    Set newRow = lendingTbl.ListRows.Add()
    
    newRow.Range.Cells(1, GetColumnIndex(lendingTbl, COL_RECORD_ID)).Value = GetNextRecordID()
    newRow.Range.Cells(1, GetColumnIndex(lendingTbl, COL_LENDING_ITEM_ID)).Value = availableItemID
    newRow.Range.Cells(1, GetColumnIndex(lendingTbl, COL_LENDING_ITEM_NAME)).Value = GetItemName(availableItemID)
    newRow.Range.Cells(1, GetColumnIndex(lendingTbl, COL_BORROWER)).Value = "テストユーザー " & Format(Now, "mmss")
    newRow.Range.Cells(1, GetColumnIndex(lendingTbl, COL_LEND_DATE)).Value = Date
    newRow.Range.Cells(1, GetColumnIndex(lendingTbl, COL_DUE_DATE)).Value = Date + DEFAULT_LENDING_DAYS
    newRow.Range.Cells(1, GetColumnIndex(lendingTbl, COL_RETURN_DATE)).Value = ""
    newRow.Range.Cells(1, GetColumnIndex(lendingTbl, COL_STATUS)).Value = STATUS_LENDING
    newRow.Range.Cells(1, GetColumnIndex(lendingTbl, COL_REMARKS)).Value = "テスト貸出データ"
    
    ' フォーマット適用
    Call ApplyStandardTableFormat(lendingTbl)
    
    ' ダッシュボード更新
    Call UpdateDashboard
    
    ' ログ記録
    Call LogAudit("テスト貸出作成", "ItemID: " & availableItemID)
    
    MsgBox "テスト貸出（備品ID: " & availableItemID & "）を作成しました。", vbInformation
    
    Exit Sub
    
ErrHandler:
    Call LogError("CreateTestLending", Err.Number, Err.Description)
    MsgBox "テスト貸出作成中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' デバッグ情報表示
Public Sub ShowDebugInfo()
    On Error GoTo ErrHandler
    
    Dim info As String
    info = "=== システム状態デバッグ情報 ===" & vbCrLf & vbCrLf
    
    ' テーブル状態
    Dim itemsTbl As ListObject, lendingTbl As ListObject
    Set itemsTbl = GetItemsTable()
    Set lendingTbl = GetLendingTable()
    
    info = info & "備品テーブル: "
    If itemsTbl Is Nothing Then
        info = info & "なし" & vbCrLf
    Else
        info = info & "あり（行数: "
        If itemsTbl.DataBodyRange Is Nothing Then
            info = info & "0"
        Else
            info = info & itemsTbl.DataBodyRange.Rows.Count
        End If
        info = info & "）" & vbCrLf
    End If
    
    info = info & "貸出テーブル: "
    If lendingTbl Is Nothing Then
        info = info & "なし" & vbCrLf
    Else
        info = info & "あり（行数: "
        If lendingTbl.DataBodyRange Is Nothing Then
            info = info & "0"
        Else
            info = info & lendingTbl.DataBodyRange.Rows.Count
        End If
        info = info & "）" & vbCrLf
    End If
    
    ' 統計情報
    info = info & vbCrLf & "統計情報:" & vbCrLf
    info = info & "総備品数: " & GetTotalItemsCount() & vbCrLf
    info = info & "貸出中件数: " & GetTotalLendingCount() & vbCrLf
    info = info & "期限超過件数: " & GetOverdueCount() & vbCrLf
    
    ' システム情報
    info = info & vbCrLf & "システム情報:" & vbCrLf
    info = info & "現在日時: " & Format(Now, "yyyy/mm/dd hh:mm:ss") & vbCrLf
    info = info & "ユーザー: " & Application.UserName & vbCrLf
    info = info & "ワークブック: " & ThisWorkbook.Name & vbCrLf
    
    MsgBox info, vbInformation, "システムデバッグ情報"
    
    Exit Sub
    
ErrHandler:
    Call LogError("ShowDebugInfo", Err.Number, Err.Description)
    MsgBox "デバッグ情報取得中にエラーが発生しました: " & Err.Description, vbCritical
End Sub