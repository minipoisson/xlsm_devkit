# xlsm_devkit

Excel VBA モジュールの入出力とシートマップのエクスポートを行う開発支援ツールです。  
VBE（Visual Basic Editor）と外部エディタ（VS Code など）を組み合わせた xlsm 開発ワークフローをサポートします。

## 機能

| マクロ | 内容 |
| :--- | :--- |
| `ExportAllModules` | プロジェクト内の全モジュールを `src/*.bas` に書き出す（ANSI → UTF-8 変換） |
| `ImportAllModules` | `src/*.bas` を VBA プロジェクトに読み込む（UTF-8 → ANSI 変換）。`xlsm_devkit` 自身はスキップ |
| `ExportAllSheetMapsToMD` | 全シートのセル・図形・数式・スタイルを `sheet/*.md` に書き出す |
| `ImportAllSheetMapsFromMD` | `sheet/*.md` からセル値・数式・スタイル・名前付き範囲・入力規則リスト・結合セルをワークブックに復元する |

## 使い方

### 導入手順（新規ブックに組み込む）

1. 開発対象の `.xlsm` を開き、`Alt + F11` で VBE を開く。
2. プロジェクトエクスプローラーで対象プロジェクトを右クリックし、`ファイルのインポート` から `xlsm_devkit.bas` を読み込む。
3. Excel の設定で「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を有効にする。
4. `ExportAllModules` を実行し、ブックと同じフォルダに `src/` が作成されることを確認する。
5. VS Code などで `src/*.bas` を編集し、`ImportAllModules` を実行する。

### モジュールのエクスポート

1. Excel で `ExportAllModules` マクロを実行する。
2. ブックと同じフォルダの `src/` に各モジュールの `.bas` ファイルが UTF-8 で書き出される。

### モジュールのインポート

1. `ImportAllModules` マクロを実行する。
2. `src/*.bas` がプロジェクトに読み込まれる。既存モジュールはコードが上書きされ、新規モジュールは追加される。

### シートマップのエクスポート

1. `ExportAllSheetMapsToMD` マクロを実行する。
2. ブックと同じフォルダの `sheet/` に各シートの Markdown ファイルが書き出される。

### シートマップのインポート

1. `ImportAllSheetMapsFromMD` マクロを実行する。
2. `sheet/*.md` が順に読み込まれ、対応するシートにセル値・数式・背景色・文字色・フォントサイズ・入力規則リスト・名前付き範囲・結合セルが復元される。

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
  src/          # エクスポートされた .bas ファイル（BOM なし UTF-8）
  sheet/        # エクスポートされたシートマップ .md ファイル（BOM なし UTF-8）
```

## 制約事項

- `xlsm_devkit` 自身は `ImportAllModules` によってインポートされません（実行中のモジュールは上書き・削除できないため）。`xlsm_devkit` を更新する場合は、手動で VBE に貼り付けるか、別途インポート処理を行ってください。
- フォームモジュール（`Type = 3`）はエクスポート対象外です。
- Windows + Excel VBA 環境が必要です。

## 動作確認環境

- Windows
- 検証済み: Microsoft Excel 2010 以降（32-bit / 64-bit）

## 対応バージョン

- 32-bit Excel: 2007 以降を想定（VBA6 分岐あり、未検証）
- 64-bit Excel: 2010 以降（`VBA7` / `PtrSafe` が必須）
- Windows API（`GetACP`）と `VBProject` 操作を使用するため、Windows 版 Excel が前提
