# XVBA Mock Creator

Excel VBAプロジェクト用の完全なXVBA（Extended VBA）開発環境を作成するツールです。モダンな開発ツールチェーンを使用して、Excel VBAアプリケーションの開発を効率化します。

## 主な機能

- **Excel VBAプロジェクトの自動生成**: 完全なプロジェクト構造とサンプルコードの作成
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

1. **モック生成**: `@xvba-mock-creator <作りたいモック情報>` でカスタマイズファイルを生成
2. **開発**: `customize/vba-files/` でVBAソースコードを編集（UTF-8エンコーディング）
3. **ビルド**: `.\xvba_pre_export.ps1` を実行してExcel用ファイルを生成
4. **インポート**: `xvba-macro list` でVBAモジュールをExcelに一括インポート
5. **テスト**: Excel上でVBAコードをテスト・デバッグ

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

## ライセンス

このプロジェクトはモダンなExcel VBA開発環境の構築を目的としています。