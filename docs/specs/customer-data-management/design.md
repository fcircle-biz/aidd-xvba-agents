# 顧客情報 取込＆整形VBAシステム 設計書（フォームなし）

## 1. 概要

* **目的**：外部CSVの顧客データをExcelに取り込み、既存レコードへ安全に**追記・更新（Upsert）**、重複排除、論理削除、整形、検証、レポート出力までを自動化。
* **成果物**：Excelブック（.xlsm）＋VBAモジュール。
* **運用想定**：営業/CS部門が日次・週次でCSVを受領し、ワンクリックで更新。

## 2. 対象環境

* Excel 2016 以降（Windows）
* マクロ有効ブック（.xlsm）

## 3. シート構成

| シート名          | 用途                                     |
| ------------- | -------------------------------------- |
| **Dashboard** | 取込ボタン、KPI（総件数/更新件数/重複検出/無効化件数）、最終取込日時  |
| **Customers** | 顧客マスタ本体（Excelテーブル `tblCustomers`）      |
| **Staging**   | 取込一時領域（`tblStaging`：CSV→ここへ読み込み＆整形）    |
| **\_Config**  | 設定（CSVフォルダ/ファイル名、キー定義、必須列、検証ルール、無効化条件） |
| **Logs**      | 取込ログ・検証エラー・更新差分の記録                     |
| **Codebook**  | 列定義と変換ルール（カラム名マッピング、型・正規化規則）           |

> ポイント：**Staging→Customers** の二段構えで“安全に”取り込む。Customersは常に参照の正本。

## 4. データモデル

### 4.1 顧客マスタ `tblCustomers`（Customers）

| 列名           | 型   | 必須 | 説明                          |
| ------------ | --- | -- | --------------------------- |
| CustomerID   | 文字列 | ◎  | 取引先コード/CRM ID（外部と連携可能な安定キー） |
| CustomerName | 文字列 | ◎  | 顧客名                         |
| Email        | 文字列 | △  | メインメール（重複検出に使用可）            |
| Phone        | 文字列 |    | 半角数字＋ハイフン正規化                |
| Zip          | 文字列 |    | 郵便番号（正規化）                   |
| Address1     | 文字列 |    | 都道府県＋市区郡                    |
| Address2     | 文字列 |    | 丁目・番地等                      |
| Category     | 文字列 |    | 顧客区分（B2B/B2C/代理店…）          |
| Status       | 文字列 | ◎  | 有効/無効                       |
| CreatedAt    | 日時  | ◎  | 取込作成日時                      |
| UpdatedAt    | 日時  | ◎  | 最終更新日時                      |
| SourceFile   | 文字列 |    | 反映元CSV名                     |
| Notes        | 文字列 |    | 備考                          |

**インデックス運用（推奨）**：`CustomerID` を主キー相当、補助で `Email`。
**重複判定キー例**：`CustomerID` を第一キー、なければ `Email + CustomerName` を代替キー。

### 4.2 取込一時 `tblStaging`（Staging）

* CSV列そのまま＋正規化済みの中間列（例：`Email_norm`, `Phone_norm`, `Key_Candidate`）。
* エラーフラグ列：`IsValid`, `ErrorMessage`。

## 5. CSV仕様（例）

* **文字コード**：UTF-8（BOM可）
* **ヘッダ行**：あり
* **列**（例）：`customer_id,customer_name,email,phone,zip,address1,address2,category,status`
* **空値**は空文字で表現、改行なし。

## 6. フロー

1. **CSV取込**（Stagingへ）

   * `_Config` のフォルダ/ファイル名を参照し `tblStaging` に読み込み
   * ヘッダマッピング（別名→正規列名）
2. **整形・正規化**

   * 文字種統一（全角→半角、トリム、連続空白の単一化）
   * 電話/郵便番号のフォーマット化、メール小文字化、住所結合/分解
   * カテゴリ値の正規化（コード表に基づく）
3. **検証**

   * 必須列（ID or 代替キー、顧客名、Status）チェック
   * 型/桁/パターン（Email / Phone / Zip）チェック
   * 代替キー組成（Email+Name など）
4. **重複検出**

   * Staging 内重複、Customers 既存との重複
5. **Upsert**

   * 既存一致：更新（差分のみ更新、UpdatedAt更新）
   * 未存在：追加（CreatedAt/UpdatedAt設定）
6. **論理削除（任意）**

   * Staging に存在しない“長期間未更新のID”を `Status="無効"` に（期間は `_Config`）
7. **ログ出力 & KPI更新**

   * 取込件数、追加件数、更新件数、重複/エラー件数、無効化件数、所要時間
8. **完了**

   * `Dashboard` に最終取込日時と実績表示

## 7. 操作UI（Dashboard）

* ボタン：

  * 「CSV取込→整形→検証→反映（Upsert）」＝ワンクリックで実行
  * 「差分レポート出力（Logsへ）」
  * 「設定を開く（\_Configへ移動）」
  * 「安全クリア（Staging初期化）」
* KPI表示：総件数／追加／更新／重複／エラー／無効化

## 8. 設定 `_Config`（例）

| キー               | 値                                | 説明       |
| ---------------- | -------------------------------- | -------- |
| CSV\_DIR         | `C:\Data\Import\`                | 取込フォルダ   |
| CSV\_FILE        | `customers_*.csv`                | パターン対応   |
| PRIMARY\_KEY     | `CustomerID`                     | 主キー列     |
| ALT\_KEY         | `Email+CustomerName`             | 代替キー式    |
| REQUIRED         | `CustomerID,CustomerName,Status` | 必須列      |
| INACTIVATE\_DAYS | `180`                            | 未更新閾値（日） |
| EMAIL\_REGEX     | `...`                            | 簡易パターン   |
| ZIP\_REGEX       | `^\d{3}-?\d{4}$`                 | 郵便番号     |
| PHONE\_REGEX     | `^[0-9\-]+$`                     | 電話番号     |

## 9. VBAモジュール構成

* `modInit`：名前定義・テーブル存在確認、初期化
* `modImport`：CSV→Staging 取込、ヘッダマッピング、正規化
* `modValidate`：必須・型・正規表現・重複検出
* `modUpsert`：Customers への追加/更新、論理削除
* `modDashboard`：KPI集計、ボタンイベント
* `modUtils`：日付/文字列/正規表現/ログ/安全クリア

## 10. 主要処理（疑似コード＋代表スニペット）

### 10.1 安全クリア（Staging）

```vb
Public Sub SafeClearStaging()
    Dim ws As Worksheet
    Set ws = Sheets("Staging")
    SafeClearSheet ws, keepFormats:=True    ' 値のみクリア
End Sub
```

```vb
' 共通：安全クリア
Public Sub SafeClearSheet(ws As Worksheet, Optional keepFormats As Boolean = False)
    Dim found As Range, isProt As Boolean
    On Error GoTo ExitHandler
    isProt = ws.ProtectContents
    If isProt Then ws.Unprotect Password:="yourpwd"
    Set found = ws.Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If Not found Is Nothing Then
        With ws.Range("A1", ws.Cells(found.Row, found.Column))
            If keepFormats Then .ClearContents Else .Clear
        End With
    End If
ExitHandler:
    If isProt Then ws.Protect Password:="yourpwd", UserInterfaceOnly:=True
End Sub
```

### 10.2 取込（CSV→Staging）

```vb
Public Sub ImportCsvToStaging()
    Dim dirPath$, filePattern$, f$, t As ListObject
    dirPath = GetConfig("CSV_DIR")
    filePattern = GetConfig("CSV_FILE")
    Set t = GetTable("Staging", "tblStaging")

    SafeClearStaging

    f = Dir(dirPath & filePattern)
    Do While f <> ""
        AppendCsvIntoTable dirPath & f, t    ' 開いて読み込み→行追加
        f = Dir()
    Loop

    NormalizeStagingColumns t
    MapHeadersToCanonical t      ' 列名マッピング（Codebookに基づく）
End Sub
```

### 10.3 整形（正規化の例）

```vb
Public Sub NormalizeStagingColumns(t As ListObject)
    Dim r As ListRow, email$, phone$, zip$
    For Each r In t.ListRows
        r.Range(1, t.ListColumns("CustomerName").Index).Value = TrimAll(CStr(GetVal(r, "CustomerName")))
        email = LCase(TrimAll(CStr(GetVal(r, "Email"))))
        SetVal r, "Email_norm", email
        phone = NormalizePhone(CStr(GetVal(r, "Phone")))
        SetVal r, "Phone_norm", phone
        zip = NormalizeZip(CStr(GetVal(r, "Zip")))
        SetVal r, "Zip_norm", zip
        SetVal r, "Key_Candidate", BuildAltKey(r) ' Email+Name など
    Next
End Sub
```

### 10.4 検証

```vb
Public Function ValidateStaging(t As ListObject) As Long
    Dim errors&: errors = 0
    Dim req As Collection: Set req = SplitToCollection(GetConfig("REQUIRED"))
    Dim r As ListRow, msg$

    For Each r In t.ListRows
        msg = ""
        ' 必須
        msg = msg & CheckRequired(r, req)
        ' 形式
        msg = msg & CheckRegex(r, "Email_norm", GetConfig("EMAIL_REGEX"))
        msg = msg & CheckRegex(r, "Zip_norm", GetConfig("ZIP_REGEX"))
        msg = msg & CheckRegex(r, "Phone_norm", GetConfig("PHONE_REGEX"))
        ' 重複（Staging内）
        msg = msg & CheckDupInStaging(r, t)
        If Len(msg) > 0 Then
            SetVal r, "IsValid", False
            SetVal r, "ErrorMessage", msg
            errors = errors + 1
        Else
            SetVal r, "IsValid", True
        End If
    Next
    ValidateStaging = errors
End Function
```

### 10.5 Upsert（Customersへ反映）

```vb
Public Sub UpsertCustomers()
    Dim stg As ListObject, cus As ListObject, r As ListRow
    Set stg = GetTable("Staging", "tblStaging")
    Set cus = GetTable("Customers", "tblCustomers")

    Dim addCnt&, updCnt&
    For Each r In stg.ListRows
        If GetVal(r, "IsValid") <> True Then GoTo ContinueFor

        Dim id$, key$
        id = CStr(GetVal(r, "CustomerID"))
        key = IIf(Len(id) > 0, id, CStr(GetVal(r, "Key_Candidate")))

        Dim target As ListRow
        Set target = FindCustomerRow(cus, id, key)

        If target Is Nothing Then
            AddCustomerRow cus, r
            addCnt = addCnt + 1
        Else
            If ApplyDiffUpdate(target, r) Then updCnt = updCnt + 1
        End If
ContinueFor:
    Next

    WriteLog "Upsert", "Added=" & addCnt & ", Updated=" & updCnt
    UpdateKPI
End Sub
```

**差分更新の考え方**

* 既存値とStaging値を**トリム後比較**し、差分がある列のみ上書き。
* 日時列 `UpdatedAt` を都度更新、`SourceFile` を記録。

### 10.6 論理削除（無効化）

```vb
Public Sub InactivateStaleCustomers()
    Dim cus As ListObject: Set cus = GetTable("Customers", "tblCustomers")
    Dim limitDays&: limitDays = CLng(GetConfig("INACTIVATE_DAYS"))
    Dim cutoff As Date: cutoff = Now - limitDays

    Dim r As ListRow, lastUpd As Date
    For Each r In cus.ListRows
        lastUpd = NzDate(GetCell(r, "UpdatedAt"))
        If lastUpd > 0 And lastUpd < cutoff And GetText(r, "Status") <> "無効" Then
            SetText r, "Status", "無効"
            SetDate r, "UpdatedAt", Now
            AppendLog "Inactivate", GetText(r, "CustomerID")
        End If
    Next
End Sub
```

## 11. エラー処理・ログ

* すべての公開サブルーチンで `On Error GoTo ErrHandler` を徹底。
* `Logs` に以下を追記：日時、処理名、レベル（INFO/WARN/ERROR）、メッセージ、件数、対象ファイル名。
* **ユーザー向けメッセージ**は簡潔に、**詳細はログ**へ。

## 12. パフォーマンス指針

* `Application.ScreenUpdating = False`、`Calculation = xlCalculationManual` 切替。
* **ListObjectの配列バルク操作**（大量行は配列に読み込み→一括書き戻し）。
* `Select/Activate` 禁止。
* Find/Match は**辞書化**で加速（`Scripting.Dictionary`）。

## 13. 品質・テスト

* **リハーサル用CSV**（正常/必須欠落/形式不正/重複/大量）を用意。
* テスト観点：

  * 取込件数＝Staging件数、エラー数の一致
  * 差分更新の正しさ（更新列のみ上書き）
  * 論理削除の閾値動作
  * 正規表現の境界値（メール、郵便、電話）

## 14. 例外・制約

* CSVの列名が大幅に異なる場合は **Codebook** のマッピング更新が必要。
* 代替キーの品質はデータに依存（Email欠損が多い場合は別キー設計を検討）。

## 15. 導入初期手順（ざっくり）

1. シート（Dashboard/Customers/Staging/\_Config/Logs/Codebook）を作成し、`tblCustomers`/`tblStaging` テーブル定義を置く。
2. `_Config` に基本設定と必須列、正規表現を記載。
3. `Codebook` に**外部列名→内部正規列名**のマッピングを一覧化。
4. モジュール（`modImport/modValidate/modUpsert/modUtils`）を追加。
5. Dashboardにボタン配置し、マクロを割当。
6. テストCSVで動作確認→ログとKPIで検証。
