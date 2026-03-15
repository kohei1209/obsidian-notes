# Flow 3: Outlook予定 変更・削除 → 工数ログ同期

## フロー名
`[工数管理] Outlook予定 変更・削除 → 工数ログ同期`

## 概要
Outlookカレンダーの予定が**更新または削除**されたとき、SharePoint の `WorkHoursLog` に
登録済みの工数レコードを自動的に同期し、**HTML ダッシュボードを再生成**する。

- **更新**: 時刻・工数・プロジェクト（カテゴリ変更時）を最新値に上書き。Status は `Active` を維持する。
- **削除**: `Status = "Cancelled"` にマーク。ダッシュボードの集計から自動除外される。

> Flow 1 で登録されていない予定（スキップ済み・重複除外済み）は `WorkHoursLog` に
> レコードが存在しないため、このフローは何もせず正常終了する。

---

## ⚠️ 設計上の重要ポイント

### Status フィルターの設計方針
既存レコード検索時のフィルターは **`Status ne 'Cancelled'`**（Cancelled 以外すべて）とする。

**`Status eq 'Active'` ではダメな理由:**
もし何らかの理由で Status が Active 以外の値になった場合、フィルターにヒットせず
更新・削除が無視される。`ne 'Cancelled'` なら Active 以外の有効レコードも確実に見つかる。

### Status の運用ルール
| 操作 | Status の遷移 |
|------|--------------|
| Flow 1 で新規登録 | → `Active` |
| Flow 3a で更新 | `Active` → `Active`（**変更しない**） |
| Flow 3b で削除 | `Active` → `Cancelled` |

> Status は「このレコードが有効か無効か」を示すフラグ。
> 更新の履歴は `LastSyncedAt` フィールドで追跡する。

---

## このフローは2つのトリガーで構成する

Power Automateでは「更新」と「削除」を別トリガーで検知するため、
**同じフロー定義をコピーしてトリガーだけ変えた2本**を作成する。

| フロー | トリガー |
|--------|---------|
| Flow 3a | `予定が更新されたとき (V3)` |
| Flow 3b | `予定が削除されたとき (V3)` |

以降のステップ説明は **Flow 3a（更新）** を基準とし、
Flow 3b（削除）で異なる箇所は `【削除版】` と注記する。

---

## ステップ一覧

### Step 1: 終日イベントのスキップ
**条件分岐** アクション（Flow 3a のみ）
- 条件: `triggerBody()?['isAllDay']` が `true` に等しい
- Yes の場合: **終了**
- No の場合: 次のステップへ

> Flow 3b（削除）では終日イベントの判定情報が取れない場合があるため、このステップを省略して
> 次のステップに進んでよい。Step 2 で既存レコードが見つからなければ自動的に何もしない。

---

### Step 2: 既存レコードの検索
**コネクタ**: SharePoint
**アクション**: `複数の項目の取得`
- サイトのアドレス: `[YOUR_SHAREPOINT_SITE_URL]`
- リスト名: `WorkHoursLog`
- フィルター クエリ:
  ```
  EventId eq '@{triggerBody()?['id']}' and Status ne 'Cancelled'
  ```
- 取得する列（Select）: `Id, EventId, StartDateTime, EndDateTime, DurationHours, ProjectCode, ProjectName, Status`
- 上限: `1`

結果を変数 `ExistingRecords` に格納。

**条件分岐**: `length(body('既存レコード検索')?['value'])` が `0` に等しい場合
- Yes の場合（レコードなし＝スキップ済みまたは未登録）: **終了**
- No の場合: 次のステップへ

---

### Step 3: 既存レコードの SharePoint ID を取得
**変数の初期化** アクション
- 名前: `ExistingItemId`
- 型: `整数`
- 値（式）:
  ```
  first(body('既存レコード検索')?['value'])?['Id']
  ```

---

### ▼ Flow 3a（更新）のステップ 4〜10

### Step 4: 工数（時間数）の再計算
**変数の初期化** アクション
- 名前: `NewDurationHours`
- 型: `浮動小数点数`
- 値（式）:
  ```
  round(
    div(
      float(
        sub(
          ticks(triggerBody()?['end']),
          ticks(triggerBody()?['start'])
        )
      ),
      36000000000
    ),
    2
  )
  ```

---

### Step 5: カテゴリの取得（プロジェクト再マッピング用）
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

---

### Step 6: プロジェクトのマッチング
**コネクタ**: SharePoint
**アクション**: `複数の項目の取得`
- サイトのアドレス: `[YOUR_SHAREPOINT_SITE_URL]`
- リスト名: `Projects`
- フィルター クエリ:
  ```
  OutlookCategory eq '@{variables('EventCategory')}' and IsActive eq 1
  ```
- 上限: `1`

**変数の設定**:

**マッチありの場合**:
- `NewProjectCode` = `first(body('プロジェクトマッチング')?['value'])?['ProjectCode']`
- `NewProjectName` = `first(body('プロジェクトマッチング')?['value'])?['Title']`

**マッチなしの場合**:
- `NewProjectCode` = `''`
- `NewProjectName` = `''`

---

### Step 7: 変更内容の判定
**変数の初期化** アクション
- 名前: `HasChanges`
- 型: `ブール値`
- 値: `false`

**条件分岐** アクション: 以下のいずれかが `true` の場合、`HasChanges` を `true` に設定する。

| 判定項目 | 条件 |
|---------|------|
| 時刻変更 | `triggerBody()?['start']` ≠ 既存レコードの `StartDateTime` |
| 工数変更 | `variables('NewDurationHours')` ≠ 既存レコードの `DurationHours` |
| プロジェクト変更 | `variables('NewProjectCode')` ≠ 既存レコードの `ProjectCode` |
| 件名変更 | `triggerBody()?['subject']` ≠ 既存レコードの `Title` |

> **ポイント**: 時刻だけでなく**カテゴリ（プロジェクト）の変更**も検出する。
> これにより、Outlook でカテゴリを後から変更するだけでプロジェクトが自動更新される。

---

### Step 8: SharePoint レコードの更新
**条件分岐**: `variables('HasChanges')` が `true` の場合のみ実行

**コネクタ**: SharePoint
**アクション**: `項目の更新`
- サイトのアドレス: `[YOUR_SHAREPOINT_SITE_URL]`
- リスト名: `WorkHoursLog`
- ID: `variables('ExistingItemId')`
- 更新するフィールド:

| フィールド | 式 | 備考 |
|-----------|-----|------|
| Title | `triggerBody()?['subject']` | 件名を最新に同期 |
| StartDateTime | `triggerBody()?['start']` | |
| EndDateTime | `triggerBody()?['end']` | |
| DurationHours | `variables('NewDurationHours')` | round 済み |
| ProjectCode | `variables('NewProjectCode')` | カテゴリ変更時に自動更新 |
| ProjectName | `variables('NewProjectName')` | カテゴリ変更時に自動更新 |
| Status | `Active` | **Active を維持する** |
| LastSyncedAt | `utcNow()` | 更新日時を記録 |

---

### Step 9: Teams 通知（更新）
**コネクタ**: Microsoft Teams
**アクション**: `チャットまたはチャネルにメッセージを投稿する`
- 投稿者: `フロー ボット`
- 投稿先: `チャット`
- 受信者: `[自分のメールアドレス]`

**変更ありの場合**:
```
🔄 工数ログを更新しました。

📅 @{triggerBody()?['subject']}
⏰ @{triggerBody()?['start']} 〜 @{triggerBody()?['end']}（@{variables('NewDurationHours')}h）
📁 @{if(equals(variables('NewProjectCode'), ''), '未分類', variables('NewProjectName'))}
```

**変更なしの場合**: 通知をスキップ（件名や場所のみの変更は通知不要）

---

### Step 10: HTML ダッシュボード再生成
**条件分岐**: `variables('HasChanges')` が `true` の場合のみ実行

**コネクタ**: Power Automate（または HTTP）
**アクション**: Flow 2（`[工数管理] HTML ダッシュボード生成`）を呼び出す

> Step 8 の「項目の更新」アクションの**後ろ（Succeeded ブランチ）** に配置する。
> 変更がない場合は HTML 再生成をスキップしてフロー実行回数を節約する。

---

### ▼ Flow 3b（削除）のステップ 4〜6

### Step 4: SharePoint レコードを Cancelled にマーク
**コネクタ**: SharePoint
**アクション**: `項目の更新`
- サイトのアドレス: `[YOUR_SHAREPOINT_SITE_URL]`
- リスト名: `WorkHoursLog`
- ID: `variables('ExistingItemId')`
- 更新するフィールド:

| フィールド | 値 |
|-----------|-----|
| Status | `Cancelled` |
| LastSyncedAt | `utcNow()` |

> レコードを物理削除せず `Status = Cancelled` にする理由:
> 誤って予定を消した場合の復旧が可能になるため。
> SharePoint リスト上で手動で `Status` を `Active` に戻せば集計に復活する。

---

### Step 5: Teams 通知（削除）
**コネクタ**: Microsoft Teams
**アクション**: `チャットまたはチャネルにメッセージを投稿する`
- メッセージ:
  ```
  🗑️ 予定「@{first(body('既存レコード検索')?['value'])?['Title']}」が削除されたため、
  工数ログをキャンセル済みにしました。

  誤って削除した場合は WorkHoursLog リストで Status を Active に戻してください。
  ```

---

### Step 6（Flow 3b）: HTML ダッシュボード再生成
Flow 2 を呼び出す（Flow 3a の Step 10 と同じ）。

> Step 4 の「項目の更新」が完了してから呼び出す（Succeeded ブランチに配置）。

---

## エラーハンドリング

| 状況 | 対処 |
|------|------|
| 既存レコードが見つからない（スキップ済み等） | Step 2 の条件分岐で検出し正常終了 |
| 予定更新で時刻・カテゴリ・件名いずれも変更なし | Step 7 で `HasChanges=false` → 更新・通知・HTML再生成をスキップ |
| SharePoint 項目更新が失敗 | 「失敗した場合に実行」を設定し、Teams にエラー通知 |
| Flow 2 の呼び出しが失敗 | Teams にエラー通知。手動で Flow 2 を実行して復旧可能 |

---

## セットアップ手順

1. Power Automate で新規フローを作成
2. **Flow 3a**: トリガー `予定が更新されたとき (V3)` → カレンダーID: `予定表` を選択
3. 上記ステップを順に追加
4. **Flow 3b**: Flow 3a を**コピー**し、トリガーのみ `予定が削除されたとき (V3)` に変更
5. Flow 3b の Step 4〜6 を削除版の内容に差し替え（Step 4〜10 は不要）
6. SharePoint `WorkHoursLog` に `Status`（選択肢列）と `LastSyncedAt`（日時列）が存在することを確認

---

## テスト観点チェックリスト

| # | テストケース | 期待結果 |
|---|-------------|---------|
| 1 | 予定の時刻を変更 | SharePoint レコードの時刻・工数が更新される |
| 2 | 同じ予定の時刻をさらに変更（2回目） | 再度更新される（Status=Active のまま） |
| 3 | 予定のカテゴリを「プロジェクトA」→「プロジェクトB」に変更 | ProjectCode・ProjectName が自動更新される |
| 4 | 予定のカテゴリを削除（なしに変更） | ProjectCode・ProjectName が空文字に更新される |
| 5 | 予定を削除 | Status が Cancelled になる |
| 6 | Flow 1 でスキップされた予定を更新 | レコードが見つからず正常終了 |
| 7 | 件名のみ変更（時刻・カテゴリ変更なし） | Title が更新される（HasChanges=true） |
| 8 | 場所のみ変更（件名・時刻・カテゴリ変更なし） | HasChanges=false → 何もしない |
| 9 | 更新後にダッシュボード HTML を確認 | 最新の時刻・プロジェクトで表示される |
| 10 | 削除後にダッシュボード HTML を確認 | 削除した工数が集計から除外されている |
