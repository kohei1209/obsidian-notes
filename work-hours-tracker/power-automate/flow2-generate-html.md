# Flow 2: HTML ダッシュボード自動生成

## フロー名
`[工数管理] HTML ダッシュボード生成`

## 概要
SharePoint の WorkHoursLog と Projects リストから全データを取得し、
Chart.js を使ったセルフコンテインド HTML ファイルを生成して
SharePoint の `WorkHoursDashboard` ドキュメントライブラリに保存する。

---

## トリガー

**コネクタ**: Power Automate
**アクション**: `HTTP 要求を受信したとき`（Flow 1 から呼び出される）

または手動テスト用に:
**アクション**: `手動でフローをトリガーします`

---

## ステップ一覧

### Step 1: WorkHoursLog 全件取得
**コネクタ**: SharePoint
**アクション**: `複数の項目の取得`
- サイトのアドレス: `[YOUR_SHAREPOINT_SITE_URL]`
- リスト名: `WorkHoursLog`
- 並び替え: `StartDateTime desc`
- 上限: `5000`（必要に応じてページネーション対応）

結果を変数 `LogItems` に格納

---

### Step 2: Projects 全件取得
**コネクタ**: SharePoint
**アクション**: `複数の項目の取得`
- サイトのアドレス: `[YOUR_SHAREPOINT_SITE_URL]`
- リスト名: `Projects`
- フィルター クエリ: なし（全件）

結果を変数 `ProjectItems` に格納

---

### Step 3: JSON 文字列化
**Compose** アクション（名前: `LogsJson`）
- 入力: `string(body('WorkHoursLog取得')?['value'])`

**Compose** アクション（名前: `ProjectsJson`）
- 入力: `string(body('Projects取得')?['value'])`

---

### Step 4: HTML ファイルを生成
**Compose** アクション（名前: `GeneratedHTML`）
- 入力: 以下のHTML文字列（`outputs('LogsJson')` と `outputs('ProjectsJson')` を埋め込む）

```
concat(
  '<!DOCTYPE html><html lang="ja">...',
  '<script>const WORK_LOGS = ',
  outputs('LogsJson'),
  '; const PROJECTS = ',
  outputs('ProjectsJson'),
  '; const GENERATED_AT = "',
  utcNow(),
  '";</script>',
  '...(残りのHTMLテンプレート)...',
  '</html>'
)
```

> **注意**: Power Automate の式エディタでは concat() を使い、
> HTMLテンプレートの各部分を文字列として連結します。
> 実際のHTMLテンプレートは `dashboard/index.html` を参照してください。
> Power Automate の式の上限（～100KB）を超える場合は、
> HTMLを複数の Compose に分割して最後に concat してください。

---

### Step 5: SharePoint にファイルを保存
**コネクタ**: SharePoint
**アクション**: `ファイルの作成`
- サイトのアドレス: `[YOUR_SHAREPOINT_SITE_URL]`
- フォルダーのパス: `/WorkHoursDashboard`
- ファイル名: `index.html`
- ファイル コンテンツ: `outputs('GeneratedHTML')`

> 既存ファイルを上書きする場合は「ファイルの作成」の代わりに
> 「ファイルのコンテンツの更新」アクションを使用してください。
> その場合は事前にファイルIDを取得する必要があります。

**代替（上書き対応）**:
1. `ファイルのメタデータの取得` でファイルIDを取得（エラー処理付き）
2. ファイルが存在する場合: `ファイルのコンテンツの更新`
3. ファイルが存在しない場合: `ファイルの作成`

---

## HTMLテンプレートと Power Automate の連携方法

`dashboard/index.html` のプレースホルダーを置換するアプローチ:

| プレースホルダー | Power Automate の置換値 |
|----------------|------------------------|
| `/* LOGS_DATA_JSON */` | `outputs('LogsJson')` |
| `/* PROJECTS_DATA_JSON */` | `outputs('ProjectsJson')` |
| `/* GENERATED_AT */` | `utcNow()` |

**Power Automate での置換式**:
```
replace(
  replace(
    replace(
      variables('HTMLTemplate'),
      '/* LOGS_DATA_JSON */',
      outputs('LogsJson')
    ),
    '/* PROJECTS_DATA_JSON */',
    outputs('ProjectsJson')
  ),
  '/* GENERATED_AT */',
  utcNow()
)
```

> HTMLテンプレートを変数 `HTMLTemplate` に格納するには、
> テンプレートを SharePoint に別ファイルとして保存し、
> 「ファイルのコンテンツの取得」で読み込む方法を推奨します。

---

## SharePoint にテンプレートを保存する方法（推奨）

1. `dashboard/index.html` を `WorkHoursDashboard` ライブラリの `_template` フォルダに保存
2. Flow 2 の先頭で SharePoint からテンプレートを読み込む:
   - `ファイルのコンテンツの取得`: `WorkHoursDashboard/_template/index.html`
3. テンプレートをbase64デコードして文字列に変換:
   - `base64ToString(body('ファイルのコンテンツの取得')?['body'])`
4. プレースホルダーを replace() で置換
5. 置換後の文字列を `index.html` として保存

---

## テスト手順

1. Flow 2 を手動トリガーで実行する
2. SharePoint の `WorkHoursDashboard` フォルダに `index.html` が作成/更新されることを確認
3. `index.html` をダウンロードしてブラウザで開く
4. 円グラフ・棒グラフ・テーブルが表示されることを確認

---

## ダッシュボードの閲覧方法

1. SharePoint サイトを開く
2. `WorkHoursDashboard` ドキュメントライブラリを開く
3. `index.html` を選択して「ダウンロード」
4. ダウンロードしたファイルをダブルクリックしてブラウザで開く
