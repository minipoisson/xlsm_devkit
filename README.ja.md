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
| `CallInitDevMode` | 現在のブックから `DEV_<name>.xlsm` を作成し、`src/` の `devkit_*` ファイルをすべてインポートする。ソースが `.xlsx` の場合も DEV_ コピーは自動的に `.xlsm` として作成される。 |
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

### テストハーネス

ブックをコードのようにテストします。オプションの `devkit_Test.bas` モジュールと
PowerShell ランナーが、実務帳票で価値の高い検証に答えます。すなわち
**「入力セルにどんな境界値・空欄・不正値・ランダム値を入れても、結果シートに
Excel エラー（`#DIV/0!`, `#N/A` など）が出ないか」**。

- 標準の Windows PowerShell のみで完結（Python・外部モジュール不要）。
- 仕様は JSON（`tests/workbook.meta.json`, `tests/no_error.test.json`）、結果は
  `test-results/latest/result.{json,md}` に出力。
- シートはタブ名（`sheet`）または VBA の `code_name` で指定可能。テスト前に Markdown の
  `fixtures` でセルへ値・数式を流し込めます。
- 原本ブックは一切変更しません（毎回ブックフォルダ内の作業コピーで実行＝ブックの信頼できる
  場所を継承）。

```powershell
.\tools\New-SampleTestWorkbook.ps1                                   # サンプル生成（開発時のみ）
.\tools\Invoke-XlsmDevkitTest.ps1 -Workbook .\examples\test-harness\sample.xlsm
```

API・JSON スキーマ・実行手順の詳細は **[TESTING.md](TESTING.md)**（英語）を参照してください。

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

> **`.xlsx` ファイルをソースとして使う場合:** `CallInitDevMode` は `.xlsx` ブックにも対応しています。DEV_ コピーは常に `.xlsm`（マクロ有効形式）として作成されるため、ソース形式に関わらず devkit の全機能が使えます。開発後、`CallSaveAsRelease` で `MyTool.xlsm` というリリースファイルが作成されます。`.xlsx` 形式で配布したい場合は、そのファイルから **ファイル → 名前を付けて保存** で `.xlsx` を選んでください。

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
- 前回のエクスポート・インポート以降に変更のない Markdown ファイルは自動的にスキップされます。ファイルサイズと更新日時を同期キャッシュ（`sheet/xlsm_devkit_sync.tsv`）と比較し、変更のないシートを除外するため、未変更のブックへの再インポートはほぼ瞬時に完了します。

モジュール・フォームのみインポートする場合は `CallImportAllComponents` を実行します。  
シートマップのみインポートする場合は `CallImportAllSheetMapsFromMD` を実行します。

### 選択インポート（`xlsm_devkit.ini`）

デフォルトではシートマップ内のすべてのトークンがインポート時に適用されます。特定のカテゴリを無効にするには、ブックと同じフォルダに `xlsm_devkit.ini` ファイルを作成し、`[import]` セクションを記述します。

```ini
[import]
; 1 = インポート時に適用（キーまたはファイルが存在しない場合のデフォルト）、0 = スキップ
; 例外: unlocked のデフォルトは 0（明示的なオプトイン）-- 意図しないセルに
; Locked=False を適用するとシート保護が弱まる。セルのロック状態を復元する
; 場合のみ有効にすること。merge=0 と merge=1 の両方で動作する。
value=1
formula=1
numfmt=1
unlocked=0
list=1
bg=0
fg=0
font_size=0
bold=0
italic=0
strike=0
wrap=0
halign=0
valign=0
merge=0
hidden_rows=0
hidden_cols=0
shapes=0
```

最も効果的な設定は `merge=0` です。`ws.Cells.UnMerge` へのブロッキング COM コールをスキップするため、結合セルが多数あるシート（例：縦横に結合されたヘッダーを持つ大型レポートシート）での 8〜15 分のボトルネックを解消できます。`merge=0` の場合、結合構造はそのままに、セル値・数式のみが更新されます。

## シートマップ形式

各シートは `sheet/<CodeName>.md` としてエクスポートされます。ファイルはメタデータヘッダーで始まり、Markdown テーブル、そして省略可能な Shapes セクションが続きます。

### ファイルヘッダー

```
# Sheet Configuration
- VBA CodeName: Sheet1
- Excel UI Name: Summary
- Hidden Rows: 3, 5-7
- Hidden Columns: B, D-F
```

`Hidden Rows` と `Hidden Columns` は非表示行・列が存在する場合のみ出力されます。

### セルテーブル

```markdown
| Address | Name | Value / Label | Formula | Style |
| :--- | :--- | :--- | :--- | :--- |
| A1 | - | Hello | - | Bold |
| B2 | rate | 0.05 | `=A1*0.05` | FontSize:14; NumFmt:0.00% |
| C3 | - | !merged_left | - | - |
```

| 列 | 内容 |
| :--- | :--- |
| `Address` | セルアドレス（例: `A1`、`B12`） |
| `Name` | このセルに紐づく名前付き範囲の名前。なければ `-`。シートスコープの名前はワークシートプレフィックス付き（`Sheet1!myRange`）。 |
| `Value / Label` | 表示値（数式セルの場合は評価結果）。マージスレーブセルの場合はマージマーカー（後述）。 |
| `Formula` | バッククォートで囲まれた数式（`` `=A1*B1` ``）。数式セルでない場合は `-`。 |
| `Style` | セミコロン区切りのスタイルトークン。スタイルなしの場合は `-`。 |

値も数式も背景色もないセルはエクスポートから除外されます。

### スタイルトークン

| トークン | 値の形式 | Export 条件 | Import フラグ | 不在時の動作 |
| :--- | :--- | :--- | :--- | :--- |
| `BG:#rrggbb` | 6桁 16進 | 背景色が設定されている | `bg=1` | 変更なし |
| `FG:#rrggbb` | 6桁 16進 | 文字色が設定されている | `fg=1` | 変更なし |
| `FontSize:<n>` | 整数 pt | 標準フォントサイズと異なる | `font_size=1` | 変更なし |
| `Bold` | フラグのみ | `Font.Bold = True` | `bold=1` | `False` にリセット |
| `Italic` | フラグのみ | `Font.Italic = True` | `italic=1` | `False` にリセット |
| `Strike` | フラグのみ | `Font.Strikethrough = True` | `strike=1` | `False` にリセット |
| `Wrap` | フラグのみ | `WrapText = True` | `wrap=1` | `False` にリセット |
| `Unlocked` | フラグのみ | `Locked = False` | `unlocked=1` | `True`（Locked）にリセット |
| `NumFmt:<format>` | Excel 書式文字列 | `"General"` 以外 | `numfmt=1` | 変更なし |
| `List:<formula>` | 例: `=$A$1:$A$10` | 入力規則リストが設定されている | `list=1` | 変更なし |
| `HAlign:<value>` | 下記参照 | 出力なし（import-only） | `halign=1` | 変更なし |
| `VAlign:<value>` | 下記参照 | 出力なし（import-only） | `valign=1` | 変更なし |

`HAlign` の値: `General`、`Left`、`Center`、`Right`、`Fill`、`Justify`、`CenterAcrossSelection`、`Distributed`（大文字小文字を区別しない）。  
`VAlign` の値: `Top`、`Center`、`Bottom`、`Justify`、`Distributed`（大文字小文字を区別しない）。

**Boolean トークンのリセット動作:** Boolean import フラグ（`bold=1` 等）が有効なとき、トークンがセルのスタイル文字列に**不在**の場合は、import 時にその属性をデフォルト値にリセットします。これは export の出力方針（非デフォルト値のみ書き出す）と対称です。

### マージスレーブマーカー

結合セルのスレーブセルは `Value / Label` 列にマーカーが出力されます。

| マーカー | 意味 |
| :--- | :--- |
| `!merged_left` | マスターと同行のスレーブセル（右方向に結合） |
| `!merged_up` | マスターと同列のスレーブセル（下方向に結合） |
| `!merged_ul` | マスターの右下にあるスレーブセル |

スレーブ行は import 時にスキップされ、結合構造は別パスで再構築されます（`merge=1` が必要）。

### エスケープシーケンス

セル値とスタイルトークン値で以下の文字がエスケープされます。

| エスケープ | 元の文字 |
| :--- | :--- |
| `\\` | バックスラッシュ |
| `\n` | 改行（CR・LF・CRLF） |
| `\v` | 垂直タブ（Chr 11） |
| `\|` | パイプ（Markdown テーブル区切り） |
| `\;` | セミコロン（スタイルトークン区切り ― スタイル値のみ） |

Shapes フィールドでは、空欄を表すセンチネル値 `-` と区別するため、内容が文字列 `-` そのものの場合は `\-` と書き出されます。

### Shapes セクション

シートに図形（テキストボックス・ボタン等）がある場合、セルテーブルの後に `## Shapes` セクションが追記されます。

```markdown
## Shapes

| Address | Name | Label | Formula | OnAction | Style |
| :--- | :--- | :--- | :--- | :--- | :--- |
| A1 | Button 1 | Click me | - | Module1.DoAction | BG:#4472C4; FG:#FFFFFF |
```

| 列 | 内容 |
| :--- | :--- |
| `Address` | 図形の左上にあるセルのアドレス |
| `Name` | 図形の内部名 |
| `Label` | 表示テキスト（なければ `-`、内容が `-` の場合は `\-`） |
| `Formula` | リンクされたセルの数式（バッククォート囲み）。なければ `-`。 |
| `OnAction` | 割り当てられたマクロ名。なければ `-`。 |
| `Style` | 図形の塗りつぶし色・文字色・フォントサイズ（`BG:`・`FG:`・`FontSize:` トークン） |

Import フラグ: `shapes=1`。

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
  src/                     # エクスポートされた .bas・.frm ファイル（BOM なし UTF-8）+ バイナリ .frx
  sheet/                   # エクスポートされたシートマップ .md ファイル（BOM なし UTF-8）
  xlsm_devkit.ini          # 任意: 選択インポート設定（選択インポート の節を参照）
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
