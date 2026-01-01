# pub_vba_modules

※ 本マクロは自己責任で使用してください。実行前に文書のバックアップを推奨します。

※ 本リポジトリのテンプレートは一例です。
　請求書以外（案内文・名簿・証明書など）でも利用できます。

Wordの差し込み印刷で、レコードごとにPDFを自動保存するVBAマクロです。
非エンジニアでも使えるように、設定は定数と文書内設定でまとめています。

## 動作環境
- Windows版 Microsoft Word（マクロ有効）

## 特長
- 1レコード=1PDFで保存
- 出力先は「文書と同じ場所/output」を自動作成
- 必要ならフォルダ選択ダイアログに切り替え可能

## 使い方
1. Wordの差し込み印刷テンプレートを開きます。
2. VBAエディタ（Alt+F11）で標準モジュールに `ExportInvoicesToPDFs_Safe.bas` の内容を貼り付けます。
※テンプレートを使用する方は既に貼り付けてあるため貼り付け不要です。
3. Excelなどの差し込みデータを接続します。
4. Alt+F8 → `ExportInvoicesToPDFs` を実行します。
※ 初回実行時に、ファイル名や出力済判定に使うフィールド名を尋ねられます。
　入力した内容は文書内に保存され、次回以降は再入力不要です。

## 出力先のルール
デフォルトでは、文書と同じフォルダに `output` を作成し、そこにPDFを保存します。

例:
```
C:\...\請求書テンプレート.docm
C:\...\output\20250101_氏名.pdf
```

## 設定項目（定数）
`ExportInvoicesToPDFs` の先頭にある定数を変更してください。

```
Const USE_FOLDER_DIALOG As Boolean = False
Const OUTPUT_SUBFOLDER As String = "output"
Const DEFAULT_NAME_FIELD As String = "お名前"
Const DEFAULT_LOG_KEY_FIELD As String = ""
Const FILE_NAME_SUFFIX As String = "様"
Const LOG_FILE_NAME As String = "export_log.csv"
```

- `USE_FOLDER_DIALOG`
  - `False`: 文書と同じ場所に `output` を自動作成
  - `True`: フォルダ選択ダイアログで保存先を指定
- `OUTPUT_SUBFOLDER`
  - 自動作成時のフォルダ名
- `DEFAULT_NAME_FIELD`
  - ファイル名に使う差し込みフィールド名（初期値）
- `DEFAULT_LOG_KEY_FIELD`
  - 出力済判定に使うフィールド名（初期値、空欄ならレコード番号）
- `FILE_NAME_SUFFIX`
  - 出力ファイル名の末尾（例: `様`）
- `LOG_FILE_NAME`
  - ログファイル名（出力先フォルダ配下）

## 文書内設定（保存される）
初回実行時に入力した値は、Word文書内に保存されます。

- `NameField`: ファイル名に使う差し込みフィールド名
- `LogKeyField`: 出力済判定に使うフィールド名（空欄ならレコード番号）

変更したい場合は、以下のマクロを実行してください。
- `SetFileNameField`
- `SetLogKeyField`

## ファイル名のルール
`yyyymmdd_お名前様.pdf` 形式で保存します。
不正な文字は自動で `_` に置換されます。
同名ファイルがある場合は `_001` のように連番を付けて上書きを防ぎます。

## 出力済み判定（ログ方式）
出力先フォルダに `export_log.csv` を作成し、出力ごとに追記します。
ログに記録済のキーは再実行時にスキップされます。
※ LogKeyField が指定されている場合は、そのフィールド値をキーとして判定します。
　未指定の場合は、レコード番号をキーとして使用します。

ログの主な項目:
- 実行日時
- 文書名
- レコード番号
- 出力ファイル名
- NameFieldの値
- LogKeyFieldの値

## 注意点
- 文書が未保存の状態だと自動出力先が作れません。
  - その場合は保存してから実行するか、`USE_FOLDER_DIALOG = True` にしてください。
- 差し込みデータに指定フィールドが無い場合は「未設定」として保存されます。
- 既定ではIRM（権限管理）は無効化しています（`KeepIRM=False`）。

## よくあるエラー
- 「差し込みデータへの接続を確認してください。」
  - Excelが閉じている/接続が外れている可能性があります。
- 「文書を保存するか、フォルダ選択に切り替えてください。」
  - 文書未保存のまま `USE_FOLDER_DIALOG = False` で実行した場合です。
