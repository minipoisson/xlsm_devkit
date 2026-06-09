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
| `CallInitDevMode` | 現在のブックから `DEV_<name>.xlsm` を作成し、`src/` の `devkit_*` ファイルをすべてインポートする |
| `CallSaveAsRelease` | `DEV_` プレフィックスを除いた本番コピーを保存し、devkit モジュールをすべて削除する（`DEV_` ブックから呼ぶ） |

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

### ランチャー

インポート・エクスポートの全操作を一箇所にまとめたダイアログです。UI 言語の切り替えも行えます。InsertDelete・Move モジュールが読み込まれている場合は、それらのボタンも表示されます。

同じ VBA プロジェクトに以下のファイルをすべてインポートしてください。

| ファイル | 役割 |
| :--- | :--- |
| `devkit_Launch.bas` | 機能ロジック |
| `devkit_frmLauncher.frm` + `devkit_frmLauncher.frx` | ランチャーダイアログ |

エントリポイント: `ShowLauncherForm`

### 国際化（i18n）

オプション機能のダイアログに表示されるテキストは、INI ベースの言語ファイルによってローカライズされます。i18n 関数（`t()`、`Fmt()`、`SetLang()`、`GetLangCode()`）は `xlsm_devkit.bas` に組み込まれており、追加の VBA モジュールは不要です。ブックと同じフォルダに `lang/` フォルダを置くだけで有効になります。

27 言語が同梱されています（アラビア語・ベンガル語・英語・スペイン語・ペルシャ語・フランス語・ドイツ語・ヒンディー語・インドネシア語・日本語・ジャワ語・韓国語・マレー語・マラーティー語・ポルトガル語・パンジャーブ語・ロシア語・スワヒリ語・タミル語・テルグ語・タイ語・トルコ語・ウクライナ語・ウルドゥー語・ベトナム語・中国語（簡体字）・中国語（繁体字））。

使用言語は Windows の設定から自動検出されます。ランチャーの言語セレクターか、以下のように直接指定して上書きすることもできます。

```vba
SetLang "ja"   ' 日本語に切り替え
SetLang ""     ' システム設定に戻す
```

## 使い方

### 導入手順（新規ブックに組み込む）

1. 開発対象の `.xlsm` を開き、`Alt + F11` で VBE を開く。
2. プロジェクトエクスプローラーで対象プロジェクトを右クリックし、`ファイルのインポート` から `xlsm_devkit.bas` を読み込む。
3. Excel の設定で「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を有効にする。
4. `ExportAllModulesFormsSheetMaps` を実行し、ブックと同じフォルダに `src/` と `sheet/` が作成されることを確認する。
5. VS Code などで `src/` 内のファイルを編集し、`ImportAllModulesFormsSheetMaps` を実行する。

### DEV / リリース ワークフロー

ユーザー向けの `.xlsm` を devkit モジュールのない状態で配布したい場合は、`DEV_` 命名規則を使います。

**開発開始 (`CallInitDevMode`)**

1. 本番ブック（例: `MyTool.xlsm`）を開き、上記の導入手順と同様に `xlsm_devkit.bas` を手動でインポートする。
2. xlsm_devkit リリースから必要な `devkit_*.bas`, `devkit_*.frm`, `devkit_*.frx` をブックと同じフォルダの `src/` に配置する。
3. マクロ ダイアログ（`Alt+F8`）から `CallInitDevMode` を実行する。`DEV_MyTool.xlsm` が同フォルダに作成され、`src/` にある `devkit_*` ファイルがすべてインポートされる。
4. `MyTool.xlsm` を**保存せずに**閉じる。こうすることで本番ファイルには devkit モジュールが残らない。
5. `DEV_MyTool.xlsm` を開いて開発を進める。

**リリース (`CallSaveAsRelease`)**

配布する準備ができたら:

1. `DEV_MyTool.xlsm` から `CallSaveAsRelease` を実行する。マクロ ダイアログ（`Alt+F8`）から呼び出すか、ランチャーの **リリースとして保存** ボタンを使う（ブック名が `DEV_` で始まる場合のみ有効）。
2. `DEV_` プレフィックスを除いた名前 `MyTool.xlsm` としてコピーが保存され、`xlsm_devkit` と `devkit_*` モジュールがそのコピーから削除される。
3. `DEV_MyTool.xlsm` はそのまま — 引き続き開発に使用できる。

### エクスポート

`ExportAllModulesFormsSheetMaps` を実行すると、全モジュール・フォームとシートマップをまとめてエクスポートします。

- 各モジュール → `src/*.bas`（BOM なし UTF-8）、各フォーム → `src/*.frm`（BOM なし UTF-8）、バイナリリソース → `src/*.frx`。
- 各シート → `sheet/*.md`（BOM なし UTF-8）。

モジュール・フォームのみエクスポートする場合は `CallExportAllComponents` を実行します。  
シートマップのみエクスポートする場合は `CallExportAllSheetMapsToMD` を実行します。

### インポート

`ImportAllModulesFormsSheetMaps` を実行すると、全モジュール・フォームとシートマップをまとめてインポートします。

- `src/` 内のファイルがプロジェクトに読み込まれます。既存のモジュール・フォームはコードが更新され、新規のものは追加されます。同フォルダの `*.frx` は VBE が自動的に参照します。
- `sheet/*.md` が対応するシートに適用されます（セル値・数式・背景色・文字色・フォントサイズ・表示形式・配置・ロック状態・入力規則リスト・名前付き範囲・結合セル）。

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

## アップグレード

`xlsm_devkit.bas` を新しいバージョンに差し替えた場合は、シートマップを再エクスポートしてください。

1. 新しい `xlsm_devkit.bas` を VBE にインポートします。
2. `CallExportAllSheetMapsToMD` を実行し、`sheet/*.md` を現行の形式で再生成します。
3. 更新された `sheet/*.md` をコミットします。

バージョン間でシートマップのフォーマットが変わることがあるため、古いバージョンで生成したファイルは新しいバージョンで正しくインポートされない場合があります。

## 制約事項

- `xlsm_devkit` 自身は `ImportAllModulesFormsSheetMaps` や `CallImportAllComponents` によってインポートされません（実行中のモジュールは上書き・削除できないため）。`xlsm_devkit` を更新する場合は、手動で VBE に貼り付けてください。
- `SKIP_DEVKIT_MODULES = True`（デフォルト）の場合、`devkit_*` オプションモジュール・フォームはインポート・エクスポートの両方でスキップされます。また `xlsm_devkit` 自身のエクスポートもスキップされます。オプションモジュールを開発する際は `xlsm_devkit.bas` 内の `SKIP_DEVKIT_MODULES` を `False` に設定してください。
- Move キャプチャの実行中は `devkit_Move` を再インポートできません（コールスタック上にあるため、インポートすると VBA ランタイムがリセットされ Excel がクラッシュします）。
- Windows + Excel VBA 環境が必要です。

## 動作確認環境

- Windows
- 検証済み: Microsoft Excel 2010 以降（32-bit / 64-bit）

## 対応バージョン

- 32-bit Excel: 2007 以降を想定（VBA6 分岐あり、未検証）
- 64-bit Excel: 2010 以降（`VBA7` / `PtrSafe` が必須）
- Windows API（`GetACP`）と `VBProject` 操作を使用するため、Windows 版 Excel が前提
