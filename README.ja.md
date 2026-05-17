# xlsm_devkit

Excel VBA モジュール・フォームの入出力とシートマップのエクスポートを行う開発支援ツールです。  
VBE（Visual Basic Editor）と外部エディタ（VS Code など）を組み合わせた xlsm 開発ワークフローをサポートします。

## コアモジュール

`xlsm_devkit.bas` 単体をインポートすると、以下のマクロが使えます。

| マクロ | 内容 |
| :--- | :--- |
| `ExportAllModulesFormsSheetMaps` | 全モジュール・フォームを `src/` に、全シートマップを `sheet/` に書き出す |
| `ImportAllModulesFormsSheetMaps` | 全モジュール・フォームを `src/` から、全シートマップを `sheet/` から読み込む |
| `CallExportAllComponents` | 全モジュール・フォームを `src/` に書き出す |
| `CallImportAllComponents` | 全モジュール・フォームを `src/` から読み込む |
| `CallExportAllSheetMapsToMD` | 全シートのセル・図形・数式・スタイルを `sheet/*.md` に書き出す |
| `CallImportAllSheetMapsFromMD` | `sheet/*.md` からセル値・数式・スタイル・名前付き範囲・入力規則リスト・結合セルを復元する |

モジュール（`.bas`）とフォーム（`.frm`/`.frx`）は同じ操作でまとめて入出力されます。  
`xlsm_devkit` 自身は `CallImportAllComponents` によってインポートされません（実行中のモジュールは上書きできないため）。

## オプション機能

以下の機能は `xlsm_devkit.bas` に加えて追加ファイルが必要です。

### InsertDelete

行・列の挿入・削除を行い、前後のシートマップを Markdown として保存し、影響を受ける VBA 参照を更新するための AI プロンプトを生成します。

同じ VBA プロジェクトに以下のファイルをすべてインポートしてください。

| ファイル | 役割 |
| :--- | :--- |
| `devkit_InsertDelete.bas` | 機能ロジック |
| `devkit_frmInsertDelete.frm` + `devkit_frmInsertDelete.frx` | セットアップダイアログ |
| `devkit_frmInstruction.frm` + `devkit_frmInstruction.frx` | 結果・インポートダイアログ（Move と共用） |

エントリポイント: `ShowInsertDeleteForm`

### Move

Excel のマクロ記録機能を使ってセル範囲の切り取り・貼り付けを記録し、前後のシートマップを取得して、影響を受ける VBA 参照を更新するための AI プロンプトを生成します。

同じ VBA プロジェクトに以下のファイルをすべてインポートしてください。

| ファイル | 役割 |
| :--- | :--- |
| `devkit_Move.bas` | 機能ロジック |
| `devkit_frmMoveSetup.frm` + `devkit_frmMoveSetup.frx` | セットアップダイアログ |
| `devkit_frmMoveWait.frm` + `devkit_frmMoveWait.frx` | 記録中ダイアログ |
| `devkit_frmInstruction.frm` + `devkit_frmInstruction.frx` | 結果・インポートダイアログ（InsertDelete と共用） |

エントリポイント: `ShowMoveSetupForm`

## 使い方

### 導入手順（新規ブックに組み込む）

1. 開発対象の `.xlsm` を開き、`Alt + F11` で VBE を開く。
2. プロジェクトエクスプローラーで対象プロジェクトを右クリックし、`ファイルのインポート` から `xlsm_devkit.bas` を読み込む。
3. Excel の設定で「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を有効にする。
4. `ExportAllModulesFormsSheetMaps` を実行し、ブックと同じフォルダに `src/` と `sheet/` が作成されることを確認する。
5. VS Code などで `src/` 内のファイルを編集し、`ImportAllModulesFormsSheetMaps` を実行する。

### エクスポート

`ExportAllModulesFormsSheetMaps` を実行すると、全モジュール・フォームとシートマップをまとめてエクスポートします。

- 各モジュール → `src/*.bas`（BOM なし UTF-8）、各フォーム → `src/*.frm`（BOM なし UTF-8）、バイナリリソース → `src/*.frx`。
- 各シート → `sheet/*.md`（BOM なし UTF-8）。

モジュール・フォームのみエクスポートする場合は `CallExportAllComponents` を実行します。  
シートマップのみエクスポートする場合は `CallExportAllSheetMapsToMD` を実行します。

### インポート

`ImportAllModulesFormsSheetMaps` を実行すると、全モジュール・フォームとシートマップをまとめてインポートします。

- `src/` 内のファイルがプロジェクトに読み込まれます。既存のモジュール・フォームはコードが更新され、新規のものは追加されます。同フォルダの `*.frx` は VBE が自動的に参照します。
- `sheet/*.md` が対応するシートに適用されます（セル値・数式・背景色・文字色・フォントサイズ・入力規則リスト・名前付き範囲・結合セル）。

モジュール・フォームのみインポートする場合は `CallImportAllComponents` を実行します。  
シートマップのみインポートする場合は `CallImportAllSheetMapsFromMD` を実行します。

## 前提条件

### VBA プロジェクト オブジェクト モデルへのアクセス

Excel の次の設定を有効にしてください。

```
ファイル → オプション → トラスト センター → トラスト センターの設定
  → マクロの設定 → 「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」にチェック
```

この設定が無効の場合、エクスポート・インポートのいずれも実行できません。

### 文字コードについて

VBA の `VBComponents.Export` / `VBComponents.Import` はシステムの ANSI コードページ（日本語環境では Shift_JIS）でファイルを読み書きします。  
このモジュールでは ADODB.Stream と Win32 API `GetACP()` を使い、ディスク上のファイルを BOM なし UTF-8 として保持しながら VBE との間で相互変換を行っています。

## ファイル構成

```
<ブックと同じフォルダ>/
  src/          # エクスポートされた .bas・.frm ファイル（BOM なし UTF-8）+ バイナリ .frx
  sheet/        # エクスポートされたシートマップ .md ファイル（BOM なし UTF-8）
```

## 制約事項

- `xlsm_devkit` 自身は `ImportAllModulesFormsSheetMaps` や `CallImportAllComponents` によってインポートされません（実行中のモジュールは上書き・削除できないため）。`xlsm_devkit` を更新する場合は、手動で VBE に貼り付けてください。
- Windows + Excel VBA 環境が必要です。

## 動作確認環境

- Windows
- 検証済み: Microsoft Excel 2010 以降（32-bit / 64-bit）

## 対応バージョン

- 32-bit Excel: 2007 以降を想定（VBA6 分岐あり、未検証）
- 64-bit Excel: 2010 以降（`VBA7` / `PtrSafe` が必須）
- Windows API（`GetACP`）と `VBProject` 操作を使用するため、Windows 版 Excel が前提
