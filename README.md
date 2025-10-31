# AIDD XVBA Agents

VBA版AI駆動開発（AI-Driven Development）プロジェクトです。Claude Codeとカスタムエージェントを活用して、Excel VBAアプリケーションの要件定義から実装、テストまでを自動化します。

**動作環境**: Windows環境専用（PowerShellスクリプトとExcel VBAを使用）

## 主な機能

- **XVBA Mock Creator**: 仕様書に基づく完全なVBA実装の自動生成
- **AI駆動開発**: Claude Codeを活用した自動実装
- **エンコーディング変換システム**: UTF-8開発ファイルからShift-JIS本番ファイルへの自動変換
- **パッケージ管理**: NPMスタイルのXVBAパッケージ依存関係管理
- **VS Code統合**: IntelliSenseとデバッグサポート
- **型定義システム**: TypeScript風のExcel VBAオブジェクト型定義

## プロジェクト構造

```
├── basefile.xlsm                    # ベースExcelワークブックテンプレート
├── config.json                      # メインプロジェクト設定
├── package.json                     # XVBA パッケージ依存関係
├── xvba_pre_export.ps1             # エンコーディング変換スクリプト
├── customize/
│   └── vba-files/                   # UTF-8開発用ソースファイル
│       ├── Class/                   # VBAクラスモジュール (.cls)
│       └── Module/                  # VBA標準モジュール (.bas)
├── vba-files/                       # Shift-JIS本番用ファイル（自動生成）
│   ├── Class/
│   └── Module/
└── xvba_modules/                    # インストールされたXVBAパッケージ
    ├── Xdebug/                      # VBAデバッグユーティリティ
    └── excel-types/                 # Excel VBAオブジェクト型定義
```

## 主要コマンド

### プロジェクトビルド
```powershell
.\xvba_pre_export.ps1
```
このスクリプトは以下を実行します：
- `customize/vba-files/` のUTF-8ファイルを `vba-files/` にShift-JISで変換
- `basefile.xlsm` を設定されたExcelファイル名にコピー
- VBAファイルをExcelインポート用に準備


## 開発ワークフロー

### 1. 初期設定
config.jsonのexcel_fileを修正

### 2. VBA実装の自動生成
```
@xvba-mock-creator <仕様書情報>
```
仕様書に基づいて完全なVBA実装を自動生成。以下が自動で作成されます：
- XVBA開発環境構成
- 設計書に基づいたVBAコード（クラス・モジュール）
- customize/vba-files/配下のソースファイル（UTF-8）

### 3. 開発・カスタマイズ
`customize/vba-files/` でVBAソースコードを確認・編集（UTF-8エンコーディング）

### 4. ビルド
```powershell
.\xvba_pre_export.ps1
```
UTF-8ファイルをShift-JISに変換してExcel用ファイルを生成

### 5. Excelへのエクスポート
`xvba-macro list` でVBAモジュールをExcelに一括エクスポート

### 6. テスト・デバッグ
Excel上でVBAコードをテスト・デバッグ

参考）https://note.com/kiyo_ai_note/n/n9653e7238c49

## 組み込み機能

### Xdebug
VBAデバッグユーティリティ
- `Xdebug.printx` - VS Codeへの変数出力
- `Xdebug.printError` - エラー情報の詳細出力

### excel-types
Excel VBAオブジェクトのTypeScript風型定義
- IntelliSenseサポート
- `.d.vb` 拡張子の型定義ファイル

## 設定ファイル

### config.json
```json
{
  "name": "プロジェクト名",
  "description": "プロジェクトの説明",
  "author": "作成者名",
  "excel_file": "target_workbook.xlsm",
  "vba_folder": "vba-files",
  "xvba_packages": {
    "Xdebug": "^1.0.0",
    "excel-types": "^1.0.0"
  }
}
```

### package.json
NPMスタイルのパッケージ管理設定

## VBA開発のベストプラクティス

### ボタンマクロの参照
```vba
' シートクラス内のプロシージャ参照
btn.OnAction = "Sheet1.ProcedureName"

' 標準モジュール内のプロシージャ参照
btn.OnAction = "ModuleName.ProcedureName"

' 呼び出されるプロシージャはPublic宣言が必要
Public Sub ProcedureName()
    ' 処理内容
End Sub
```

### システム初期化パターン
```vba
Private Sub Workbook_Open()
    Call InitializeSystem
    Call CreateInitialSampleData  ' 初回のみ実行
    Call ShowSplashScreen
End Sub

Private Sub CreateInitialSampleData()
    If IsDataEmpty() Then
        Call CreateSampleData  ' 標準モジュールの関数を呼び出し
    End If
End Sub
```

## エンコーディング処理

XVBAは文字エンコーディングの二重管理システムを採用：

- **開発時**: UTF-8エンコーディング（バージョン管理・編集に最適）
- **本番時**: Shift-JISエンコーディング（Excel VBAインポートに必要）

変換は `xvba_pre_export.ps1` スクリプトが自動で処理します。

## Excel統合

1. **ベースファイル**: `basefile.xlsm` がテンプレートワークブック
2. **ターゲットファイル**: `config.json` の `excel_file` で指定
3. **インポート**: `xvba-macro list` コマンドでVBAモジュールを一括インポート

## 要件

- Windows PowerShell
- Microsoft Excel（.xlsmファイルサポート）
- VS Code（推奨開発環境）

## XVBA Mock Creatorエージェント

このプロジェクトでは、VBA実装を自動生成する専用エージェントを提供しています：

- **xvba-mock-creator**: 仕様書に基づいてXVBA開発環境と完全なVBA実装を自動生成

エージェントは`.claude/subagents/xvba-mock-creator.md`に定義されています。

### 使用例
```
@xvba-mock-creator 在庫管理システムを作成してください。商品マスタ、入出庫履歴、在庫一覧の機能が必要です。
```

エージェントは以下を自動で行います：
1. 要件の分析と設計
2. config.json、package.jsonの作成
3. VBAクラスとモジュールの実装
4. customize/vba-files/配下へのソースコード配置
5. 実装ガイドの提供

## ライセンス

このプロジェクトはAI駆動開発によるモダンなExcel VBA開発環境の構築を目的としています。