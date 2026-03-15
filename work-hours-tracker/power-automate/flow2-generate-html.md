# Flow 2: HTML ダッシュボード自動生成

## フロー名
`[工数管理] HTML ダッシュボード生成`

## 概要
SharePoint の WorkHoursLog と Projects リストから全データを取得し、
Chart.js を使ったセルフコンテインド HTML ファイルを生成して
SharePoint の `WorkHoursDashboard` ドキュメントライブラリに保存する。

---

## トリガー

**コネクタ**: SharePoint
**アクション**: `項目が作成または変更されたとき`
- サイトのアドレス: `[YOUR_SHAREPOINT_SITE_URL]`
- リスト名: `WorkHoursLog`

> **ポイント**: Flow 1 が工数を新規登録、Flow 3 が工数を更新/Cancelledに変更 → いずれも
> WorkHoursLog の項目が変更されるため、このトリガーで**自動的に**フロー発火する。
> Flow 1/3 から別途呼び出す必要はない。
>
> **注意事項**:
> - SharePoint トリガーのポーリング間隔は**約1〜5分**。即時反映ではないが実用上問題ない。
> - 短時間に複数の予定が登録/変更されると連続発火するが、毎回最新の全データで HTML を
>   再生成するため、最終結果は常に正しい。
> - Power Automate の「トリガー条件」は設定しない（全変更で再生成する）。

---

## ステップ一覧

### Step 1: WorkHoursLog 全件取得
**コネクタ**: SharePoint
**アクション**: `複数の項目の取得`
- サイトのアドレス: `[YOUR_SHAREPOINT_SITE_URL]`
- リスト名: `WorkHoursLog`
- 取得する列（Select）: `Title, EventId, StartDateTime, EndDateTime, DurationHours, ProjectCode, ProjectName, Status, RecordedAt, LastSyncedAt`
- 並び替え: `StartDateTime desc`
- 上限: `5000`（必要に応じてページネーション対応）

> **⚠️ `Status` を必ず含めること。**
> Status が JSON に含まれないと、HTML 側のフィルタリングが機能せず、
> Cancelled レコードも集計に含まれてしまう。

結果を変数 `LogItems` に格納

---

### Step 2: Projects 全件取得
**コネクタ**: SharePoint
**アクション**: `複数の項目の取得`
- サイトのアドレス: `[YOUR_SHAREPOINT_SITE_URL]`
- リスト名: `Projects`
- 取得する列（Select）: `Title, ProjectCode, Color, PlannedHours, IsActive`
- フィルター クエリ: なし（全件）

結果を変数 `ProjectItems` に格納

---

### Step 3: フィールドマッピング（Select アクション）

SharePoint の REST API は列の型によってレスポンス形式が異なる。
HTML テンプレートが期待する**フラットな JSON キー名**と一致させるため、
**Select アクションで明示的にマッピング**する。

#### 3-1. WorkHoursLog のマッピング

**Select** アクション（名前: `MappedLogs`）
- From: `body('WorkHoursLog取得')?['value']`
- Map:

| キー（出力） | 値（SharePoint フィールド参照） | 備考 |
|-------------|-------------------------------|------|
| `Title` | `item()?['Title']` | |
| `EventId` | `item()?['EventId']` | |
| `StartDateTime` | `item()?['StartDateTime']` | |
| `EndDateTime` | `item()?['EndDateTime']` | |
| `DurationHours` | `item()?['DurationHours']` | |
| `ProjectCode` | `item()?['ProjectCode']` | 空文字の場合あり（未分類） |
| `ProjectName` | `item()?['ProjectName']` | 空文字の場合あり（未分類） |
| `Status` | `item()?['Status']?['Value']` | **⚠️ 下記注意参照** |

> **⚠️ Status フィールドの参照方法:**
> SharePoint の **選択肢列（Choice）** は REST API で
> `{"Value": "Active"}` のようなオブジェクトとして返される。
>
> - **選択肢列として作成した場合**: `item()?['Status']?['Value']` を使用
> - **テキスト列として作成した場合**: `item()?['Status']` を使用
>
> 初回実行時にフローの出力を確認し、正しい参照式を特定すること。
> 間違えると Status が `null` になり、Cancelled レコードが集計に含まれてしまう。

#### 3-2. Projects のマッピング

**Select** アクション（名前: `MappedProjects`）
- From: `body('Projects取得')?['value']`
- Map:

| キー（出力） | 値（SharePoint フィールド参照） |
|-------------|-------------------------------|
| `Title` | `item()?['Title']` |
| `ProjectCode` | `item()?['ProjectCode']` |
| `Color` | `item()?['Color']` |
| `PlannedHours` | `item()?['PlannedHours']` |
| `IsActive` | `item()?['IsActive']` |

---

### Step 4: JSON 文字列化

**Compose** アクション（名前: `LogsJson`）
- 入力: `string(body('MappedLogs'))`

**Compose** アクション（名前: `ProjectsJson`）
- 入力: `string(body('MappedProjects'))`

---

### Step 5: HTML テンプレートの読み込み
**コネクタ**: SharePoint
**アクション**: `ファイルのコンテンツの取得`
- サイトのアドレス: `[YOUR_SHAREPOINT_SITE_URL]`
- ファイル識別子: `/WorkHoursDashboard/_template/index.html`

テンプレートを文字列に変換:
```
base64ToString(body('ファイルのコンテンツの取得')?['$content'])
```

結果を変数 `HTMLTemplate` に格納。

---

### Step 6: プレースホルダーの置換
**Compose** アクション（名前: `GeneratedHTML`）

HTML テンプレート内のプレースホルダーを `replace()` で置換する:

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

| プレースホルダー | 置換値 | HTML テンプレート内の該当箇所 |
|----------------|--------|--------------------------|
| `/* LOGS_DATA_JSON */` | `outputs('LogsJson')` | `const WORK_LOGS = /* LOGS_DATA_JSON */[];` |
| `/* PROJECTS_DATA_JSON */` | `outputs('ProjectsJson')` | `const PROJECTS = /* PROJECTS_DATA_JSON */[];` |
| `/* GENERATED_AT */` | `utcNow()` | `const GENERATED_AT = "/* GENERATED_AT */";` |

> **テンプレートの書き方**: プレースホルダーの後ろにダミーデータ（空配列 `[]` や空文字 `""`）を
> 置いておくと、テンプレート単体でもブラウザで開いてレイアウト確認ができる。
> `replace()` はプレースホルダー部分のみを置換し、ダミーデータはそのまま残るが、
> JavaScript の構文上、`const WORK_LOGS = [実データ][];` は最初の配列が有効になるため問題ない。

> **⚠️ 式の文字数制限**: Power Automate の式は約 8,192 文字まで。
> HTML テンプレートは式に直接埋め込まず、必ず SharePoint からファイルとして読み込むこと。

---

### Step 7: SharePoint にファイルを保存

**方法 A（推奨）: ファイルの作成（上書きモード）**

SharePoint の「ファイルの作成」アクションは、同名ファイルが存在する場合に
上書きするオプションがないため、以下の2段階で対応する:

1. **`パスによるファイルのメタデータの取得`** で既存ファイルの ID を取得（エラー処理付き）
2. **ファイルが存在する場合**: `ファイルのコンテンツの更新` アクションで上書き
   - サイトのアドレス: `[YOUR_SHAREPOINT_SITE_URL]`
   - ファイル識別子: Step 1 で取得した ID
   - ファイル コンテンツ: `outputs('GeneratedHTML')`
3. **ファイルが存在しない場合**: `ファイルの作成` アクション
   - フォルダーのパス: `/WorkHoursDashboard`
   - ファイル名: `index.html`
   - ファイル コンテンツ: `outputs('GeneratedHTML')`

---

## テスト手順

1. SharePoint の `WorkHoursDashboard/_template/` に `index.html`（テンプレート）をアップロードする
2. Flow 2 を手動トリガーで実行する
3. SharePoint の `WorkHoursDashboard` フォルダに `index.html` が作成/更新されることを確認
4. `index.html` をダウンロードしてブラウザで開く
5. 円グラフ・棒グラフ・テーブルが表示されることを確認
6. Cancelled レコードが集計から除外されていることを確認

---

## ダッシュボードの閲覧方法

1. SharePoint サイトを開く
2. `WorkHoursDashboard` ドキュメントライブラリを開く
3. `index.html` を選択して「ダウンロード」
4. ダウンロードしたファイルをダブルクリックしてブラウザで開く

> インターネット接続が必要（Chart.js を CDN から読み込むため）。
