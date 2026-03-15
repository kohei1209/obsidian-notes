# Flow 1: Outlook予定 → カテゴリ自動判定 → 工数記録

## フロー名
`[工数管理] Outlook予定 → 工数自動記録`

## 概要
Outlookカレンダーに予定が追加されたとき、予定の**カテゴリ**から
プロジェクトを自動判定し、SharePoint に工数を記録する。

**従来方式（廃止）**: 予定追加 → Teams Adaptive Card → 手動でプロジェクト選択
**新方式**: 予定追加（カテゴリ設定済み）→ **自動でプロジェクト判定・記録**

> ユーザーは Outlook で予定を作成する際にカテゴリを設定するだけ。
> Teams での操作は不要。

---

## 事前準備: Outlook カテゴリの設定

Outlook で以下のカテゴリを作成する（色は自由）:

| Outlook カテゴリ名 | SharePoint Projects.OutlookCategory |
|-------------------|-------------------------------------|
| プロジェクトA | プロジェクトA |
| プロジェクトB | プロジェクトB |
| プロジェクトC | プロジェクトC |

> **カテゴリ名は Projects リストの `OutlookCategory` 列と完全一致させること。**
> カテゴリの色は Outlook のカレンダー表示用であり、ダッシュボードの色は
> Projects リストの `Color` 列で別途管理する。

---

## トリガー

**コネクタ**: Office 365 Outlook
**アクション**: `予定が追加されたとき (V3)`
**設定**:
- カレンダーID: `予定表`（メインカレンダー）

> 他者から招待された予定を承諾すると、このトリガーも発火する。

> **Outlook V3 トリガーの日付形式について**
>
> V3 トリガーの `start` / `end` は ISO 形式の日付文字列で返される:
> - `triggerBody()?['start']` → `"2026-03-10T09:00:00.0000000"`
> - `triggerBody()?['end']` → `"2026-03-10T10:30:00.0000000"`
>
> そのまま `ticks()` や `formatDateTime()` に渡して使える。

---

## ステップ一覧

### Step 1: 終日イベントのスキップ
**条件分岐** アクション
- 条件: `triggerBody()?['isAllDay']` が `true` に等しい
- Yes の場合: **終了**（終日イベントは工数対象外）
- No の場合: 次のステップへ

---

### Step 2: 重複チェック（同一イベントの再登録防止）
**コネクタ**: SharePoint
**アクション**: `複数の項目の取得`
- サイトのアドレス: `[YOUR_SHAREPOINT_SITE_URL]`
- リスト名: `WorkHoursLog`
- フィルター クエリ: `EventId eq '@{triggerBody()?['id']}'`
- 上限: `1`

**条件分岐**: `length(body('重複チェック')?['value'])` が `0` より大きい場合
- Yes の場合（既に登録済み）: **終了**
- No の場合: 次のステップへ

---

### Step 3: 工数（時間数）の計算

> **⚠️ Power Automate の `div()` は整数除算（小数切り捨て）です。**
> `div(ticks差, 36000000000)` だと 1.5 時間が `1` になってしまう。
> そのため**分単位で計算してから時間に変換**する方式をとる。

#### Step 3-1: 合計分数を計算する

**データ操作 > 作成**（Compose）アクション
- アクション名を `TotalMinutes` に変更する
- 「入力」欄 → **「式」タブ** → 以下を入力して OK:
  ```
  div(sub(ticks(triggerBody()?['end']),ticks(triggerBody()?['start'])),600000000)
  ```

> **式の意味**:
> - `triggerBody()?['end']` / `triggerBody()?['start']` で日時文字列を取得
> - `ticks()` で日時文字列を「tick 数」に変換
> - `sub()` で差を取る（＝所要時間の tick 数）
> - `div(..., 600000000)` で分に変換（1分 = 6億 tick）
> - `div()` は整数除算だが、通常の予定は分単位で終わるため問題ない
>
> 例: 9:00〜10:30 → 90（分）

#### Step 3-2: 分を時間（小数）に変換する

**変数の初期化** アクション
- 名前: `DurationHours`
- 型: `浮動小数点数`
- 値（式）:
  ```
  float(concat(string(div(outputs('TotalMinutes'),60)),'.',if(less(div(mul(mod(outputs('TotalMinutes'),60),100),60),10),concat('0',string(div(mul(mod(outputs('TotalMinutes'),60),100),60))),string(div(mul(mod(outputs('TotalMinutes'),60),100),60)))))
  ```

> **式が長いのでメモ帳にコピーしてから貼り付けてください。**
>
> **式の意味（分解して説明）**:
>
> | 部品 | 式 | 90分の場合 |
> |------|-----|-----------|
> | 整数部分（時間） | `div(outputs('TotalMinutes'), 60)` | `1` |
> | 余り（分） | `mod(outputs('TotalMinutes'), 60)` | `30` |
> | 小数部分 | `div(mul(余り, 100), 60)` | `div(3000, 60)` = `50` |
> | ゼロ埋め | 小数部分が 10 未満なら先頭に `0` を付ける | 不要 |
> | 文字列結合 | `concat(整数部分, '.', 小数部分)` | `"1.50"` |
> | 数値に変換 | `float("1.50")` | `1.5` |
>
> **計算例**:
> | 所要時間 | TotalMinutes | 整数部分 | 余り | 小数部分 | 結果 |
> |---------|-------------|---------|------|---------|------|
> | 30分 | 30 | 0 | 30 | 50 | **0.5** |
> | 45分 | 45 | 0 | 45 | 75 | **0.75** |
> | 1時間 | 60 | 1 | 0 | 00 | **1.0** |
> | 1時間15分 | 75 | 1 | 15 | 25 | **1.25** |
> | 1時間30分 | 90 | 1 | 30 | 50 | **1.5** |
> | 2時間 | 120 | 2 | 0 | 00 | **2.0** |

---

### Step 4: カテゴリの取得
**変数の初期化** アクション
- 名前: `EventCategory`
- 型: `文字列`
- 値（式）:
  ```
  if(
    greater(length(triggerBody()?['categories']), 0),
    first(triggerBody()?['categories']),
    ''
  )
  ```

> Outlook の `categories` は文字列の配列。複数カテゴリがある場合は最初の1つを使用する。
> カテゴリ未設定の場合は空文字になる。

---

### Step 5: プロジェクトのマッチング
**コネクタ**: SharePoint
**アクション**: `複数の項目の取得`
- サイトのアドレス: `[YOUR_SHAREPOINT_SITE_URL]`
- リスト名: `Projects`
- フィルター クエリ:
  ```
  OutlookCategory eq '@{variables('EventCategory')}' and IsActive eq 1
  ```
- 上限: `1`

結果を変数 `MatchedProjects` に格納。

---

### Step 6: プロジェクト情報の変数設定

> **⚠️ Power Automate の制約**: 「変数の初期化」アクションは
> 条件分岐やループの**中には配置できない**。
> そのため、先に空文字で初期化してから、条件分岐の中で「変数の設定」で上書きする。

#### Step 6-1: 変数を空文字で初期化する（条件分岐の**前**に配置）

**アクション①**: `変数を初期化する`
- 名前: `MatchedProjectCode`
- 型: `文字列`
- 値: `''`（空文字）

**アクション②**: `変数を初期化する`
- 名前: `MatchedProjectName`
- 型: `文字列`
- 値: `''`（空文字）

#### Step 6-2: マッチした場合のみ変数を上書きする

**条件分岐** アクション:
- 条件: `length(body('プロジェクトマッチング')?['value'])` が `0` より大きい

**Yes の場合（マッチあり）**:

**アクション③**: `変数の設定`（「変数の**初期化**」ではない）
- 名前: `MatchedProjectCode`
- 値: `first(body('プロジェクトマッチング')?['value'])?['ProjectCode']`

**アクション④**: `変数の設定`
- 名前: `MatchedProjectName`
- 値: `first(body('プロジェクトマッチング')?['value'])?['Title']`

**No の場合（マッチなし）**:
- 何もしない（Step 6-1 で空文字に初期化済み）

> **フロー上の見た目（上から順に）**:
> ```
> [変数を初期化する] MatchedProjectCode = ''
> [変数を初期化する] MatchedProjectName = ''
> [条件]  length(...) > 0 ?
>   ├─ Yes → [変数の設定] MatchedProjectCode = ...ProjectCode
>   │        [変数の設定] MatchedProjectName = ...Title
>   └─ No  → （空）
> ```

---

### Step 7: SharePoint に工数ログを登録
**コネクタ**: SharePoint
**アクション**: `項目の作成`
- サイトのアドレス: `[YOUR_SHAREPOINT_SITE_URL]`
- リスト名: `WorkHoursLog`
- 各フィールドの値:

| フィールド | 式 | 備考 |
|-----------|-----|------|
| Title | `triggerBody()?['subject']` | 予定の件名 |
| EventId | `triggerBody()?['id']` | Outlook イベントの一意 ID |
| StartDateTime | `triggerBody()?['start']` | 予定の開始日時 |
| EndDateTime | `triggerBody()?['end']` | 予定の終了日時 |
| DurationHours | `variables('DurationHours')` | |
| ProjectCode | `variables('MatchedProjectCode')` | マッチなしの場合は空文字 |
| ProjectName | `variables('MatchedProjectName')` | マッチなしの場合は空文字 |
| RecordedAt | `utcNow()` | |
| Status | `Active` | |
| LastSyncedAt | `utcNow()` | |

---

### Step 8: Teams 通知（登録完了）
**コネクタ**: Microsoft Teams
**アクション**: `チャットまたはチャネルにメッセージを投稿する`
- 投稿者: `フロー ボット`
- 投稿先: `チャット`
- 受信者: `[自分のメールアドレス]`

**カテゴリマッチありの場合**:
```
✅ 工数を自動登録しました。

📅 @{triggerBody()?['subject']}
⏰ @{formatDateTime(triggerBody()?['start'], 'yyyy/MM/dd HH:mm')} 〜 @{formatDateTime(triggerBody()?['end'], 'HH:mm')}（@{variables('DurationHours')}h）
📁 @{variables('MatchedProjectName')}
```

**カテゴリ未設定 / マッチなしの場合**:
```
⚠️ 工数を「未分類」で登録しました。

📅 @{triggerBody()?['subject']}
⏰ @{formatDateTime(triggerBody()?['start'], 'yyyy/MM/dd HH:mm')} 〜 @{formatDateTime(triggerBody()?['end'], 'HH:mm')}（@{variables('DurationHours')}h）

Outlookでカテゴリを設定するか、SharePoint WorkHoursLog でプロジェクトを手動編集してください。
```

> **通知の日時表示**: `formatDateTime()` で `2026/03/10 09:00 〜 10:30` のように
> 読みやすい形式で表示する。

> 通知の条件分岐: `equals(variables('MatchedProjectCode'), '')` で判定

---

> **HTML ダッシュボードの更新について**:
> Step 7 で WorkHoursLog に項目を作成すると、Flow 2（SharePoint トリガー）が
> **自動的に発火**するため、このフローから呼び出す必要はない。

---

## エラーハンドリング

| 状況 | 対処 |
|------|------|
| カテゴリ未設定 | 未分類として登録。Teams で警告通知を送信 |
| カテゴリが Projects に存在しない | 未分類として登録。Teams で警告通知を送信 |
| 同一 EventId が既に登録済み | Step 2 で検出し正常終了 |
| DurationHours が 0 以下 | Step 1 の終日チェックで除外。0h 予定は通常発生しない |
| SharePoint 書き込み失敗 | 「失敗した場合に実行」を設定し Teams にエラー通知 |

---

## テスト手順

1. Outlook でカテゴリ「プロジェクトA」を設定した 30 分の予定を作成する
2. Flow 1 が自動発火し、SharePoint の WorkHoursLog に 0.5h のレコードが作成されることを確認
3. Teams に「✅ 工数を自動登録しました」通知が届くことを確認
4. カテゴリなしの予定を作成し、「⚠️ 未分類で登録しました」通知が届くことを確認
5. 同じ予定を再度トリガーしても重複登録されないことを確認
6. 数分後に Flow 2 が自動発火し、WorkHoursDashboard/index.html が更新されることを確認

---

## Outlook カテゴリ運用のヒント

- **カテゴリの一括設定**: 定例会議には予めカテゴリを設定しておくと、
  毎回の手動設定が不要になる
- **招待された予定**: 承諾後に Outlook でカテゴリを設定してから保存すると、
  トリガー発火時にカテゴリが反映される
- **後からカテゴリを変更した場合**: Flow 3a（予定更新トリガー）が発火し、
  プロジェクトが自動的に再マッピングされる
- **プロジェクトの追加**: SharePoint の Projects リストに行を追加し、
  Outlook で同名のカテゴリを作成するだけで完了
