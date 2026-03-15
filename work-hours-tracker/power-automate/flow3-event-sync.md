# Flow 3: Outlook予定 変更・削除 → 工数ログ同期

## フロー名
`[工数管理] Outlook予定 変更・削除 → 工数ログ同期`

## 概要
Outlookカレンダーの予定が**更新または削除**されたとき、SharePoint の `WorkHoursLog` に
登録済みの工数レコードを自動的に同期し、**HTML ダッシュボードを再生成**する。

- **更新**: 時刻・工数を最新値に上書き。Status は `Active` を維持する。
- **削除**: `Status = "Cancelled"` にマーク。ダッシュボードの集計から自動除外される。

> Flow 1 で登録されていない予定（スキップ済み・タイムアウト済み）は `WorkHoursLog` に
> レコードが存在しないため、このフローは何もせず正常終了する。

---

## ⚠️ 設計上の重要ポイント

### Status フィルターの設計方針
既存レコード検索時のフィルターは **`Status ne 'Cancelled'`**（Cancelled 以外すべて）とする。

**`Status eq 'Active'` ではダメな理由:**
- 1回目の更新で Status が `Active` → `Active`（変更なし）→ OK
- しかし、もし Status を `Updated` に変えてしまうと、2回目以降の更新で
  `Status eq 'Active'` フィルターに**ヒットしなくなり、更新が無視される**
- 削除も同様に、更新済みレコードを削除できなくなる

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
- Yesの場合: **終了**
- Noの場合: 次のステップへ

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

結果を変数 `ExistingRecords` に格納。

> **⚠️ なぜ `Status ne 'Cancelled'` か？**
> `Status eq 'Active'` を使うと、一度更新されたレコード（Status=Updated）は
> 以降のフィルターにヒットせず、**2回目以降の更新・削除が一切反映されなくなる。**
> Cancelled 以外（Active）のレコードをすべて検索対象とすることで、
> 何度でも更新・削除を正しく同期できる。

**条件分岐**: `length(body('既存レコード検索')?['value'])` が `0` に等しい場合
- Yesの場合（レコードなし＝スキップ済みまたは未登録）: **終了**
- Noの場合: 次のステップへ

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

### ▼ Flow 3a（更新）のステップ 4〜8

### Step 4: 工数（時間数）の再計算
**変数の初期化** アクション
- 名前: `NewDurationHours`
- 型: `浮動小数点数`
- 値（式）:
  ```
  div(
    float(
      sub(
        ticks(triggerBody()?['end']),
        ticks(triggerBody()?['start'])
      )
    ),
    36000000000
  )
  ```

---

### Step 5: 時刻変更か否かを判定
**条件分岐** アクション
- 条件: `variables('NewDurationHours')` が
  `first(body('既存レコード検索')?['value'])?['DurationHours']` と等しくない
  **または** `triggerBody()?['start']` が既存レコードの `StartDateTime` と等しくない
- Yesの場合（時刻が変わった）: Step 6 へ
- Noの場合（時刻変更なし）: Step 7（プロジェクト再選択確認）へ

> **補足**: 開始・終了時刻が変わっていない場合（例: 件名だけ変更）は時刻の更新をスキップする。

---

### Step 6: SharePoint レコードの時刻・工数を更新
**コネクタ**: SharePoint
**アクション**: `項目の更新`
- サイトのアドレス: `[YOUR_SHAREPOINT_SITE_URL]`
- リスト名: `WorkHoursLog`
- ID: `variables('ExistingItemId')`
- 更新するフィールド:

| フィールド | 式 | 備考 |
|-----------|-----|------|
| Title | `triggerBody()?['subject']` | 件名も最新に同期 |
| StartDateTime | `triggerBody()?['start']` | |
| EndDateTime | `triggerBody()?['end']` | |
| DurationHours | `variables('NewDurationHours')` | |
| Status | `Active` | **⚠️ Active を維持する（Updated にしない）** |
| LastSyncedAt | `utcNow()` | 更新日時を記録 |

> **⚠️ Status を Active のまま維持する理由:**
> `Updated` に変更すると、次回の更新時に Step 2 のフィルター
> `Status ne 'Cancelled'` では見つかるが、将来的にフィルター変更した場合に
> 問題が再発するリスクがある。Status は「有効/無効」の2値として運用し、
> 更新履歴は `LastSyncedAt` で追跡する設計が最もシンプルで安全。

---

### Step 7: プロジェクト変更通知の送信（オプション）
時刻の変更とは別に、プロジェクトの変更を希望する場合に備えて Teams に通知を送る。

**コネクタ**: Microsoft Teams
**アクション**: `チャットまたはチャネルにメッセージを投稿する`
- 投稿者: `フロー ボット`
- 投稿先: `チャット`
- 受信者: `[自分のメールアドレス]`
- メッセージ（式）:
  ```
  📅 予定「@{triggerBody()?['subject']}」が更新されました。

  ⏰ 新しい時間: @{triggerBody()?['start']} 〜 @{triggerBody()?['end']}（@{variables('NewDurationHours')}h）
  📁 現在のプロジェクト: @{first(body('既存レコード検索')?['value'])?['ProjectName']}

  プロジェクトを変更する場合は WorkHoursLog リストで直接編集してください。
  ```

> **将来の拡張案**: Adaptive Card でプロジェクト再選択ボタンを提示し、
> 応答があれば ProjectCode / ProjectName も上書き更新する実装も可能。

---

### Step 8: HTML ダッシュボード再生成
**コネクタ**: Power Automate（または HTTP）
**アクション**: Flow 2（`[工数管理] HTML ダッシュボード生成`）を呼び出す

> **⚠️ SharePoint 書き込み完了を確認してから呼び出すこと。**
> Step 6 の「項目の更新」アクションの**後ろ（Succeeded ブランチ）** に配置する。
> Power Automate のアクション順序が正しければ、前のアクション完了後に
> 次のアクションが実行されるため、通常は問題ない。
> ただし、並列ブランチを使う場合は注意が必要。

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
Flow 2 を呼び出す（Flow 3a の Step 8 と同じ）。

> Step 4 の「項目の更新」が完了してから呼び出す（Succeeded ブランチに配置）。

---

## エラーハンドリング

| 状況 | 対処 |
|------|------|
| 既存レコードが見つからない（スキップ済み等） | Step 2 の条件分岐で検出し正常終了 |
| 予定更新が件名・場所のみで時刻変更なし | Step 5 の条件分岐でスキップ。Teams通知のみ送信 |
| Step 6（項目更新）が失敗 | 「失敗した場合に実行」を設定し、Teams にエラー通知 |
| Flow 2 の呼び出しが失敗 | Teams にエラー通知。手動で Flow 2 を実行して復旧可能 |

---

## セットアップ手順

1. Power Automate で `[工数管理] Outlook予定 変更・削除 → 工数ログ同期` を新規作成
2. **Flow 3a**: トリガー `予定が更新されたとき (V3)` → カレンダーID: `予定表` を選択
3. 上記ステップを順に追加
4. **Flow 3b**: Flow 3a を**コピー**し、トリガーのみ `予定が削除されたとき (V3)` に変更
5. Flow 3b の Step 4〜6 を削除版の内容に差し替える
6. SharePoint `WorkHoursLog` に `Status`（選択肢列）と `LastSyncedAt`（日時列）を追加済みであることを確認（schema.json 参照）
7. テスト手順:
   - Outlook で予定を作成 → Flow 1 で工数登録
   - 同じ予定の時刻を変更 → Flow 3a が発火し SharePoint レコードが更新されることを確認
   - **同じ予定の時刻をもう一度変更** → Flow 3a が再度発火し**2回目の更新も反映されること**を確認
   - 同じ予定を削除 → Flow 3b が発火し Status が `Cancelled` になることを確認
   - ダッシュボードで削除した工数が集計されないことを確認

---

## テスト観点チェックリスト

| # | テストケース | 期待結果 |
|---|-------------|---------|
| 1 | 予定を新規作成 → Flow 1 で登録 → 時刻を変更 | SharePoint レコードの時刻・工数が更新される |
| 2 | テスト1の予定の時刻をさらに変更（**2回目の更新**） | 再度 SharePoint レコードが更新される（**Status=Active のまま**） |
| 3 | テスト2の予定を削除 | Status が Cancelled になる |
| 4 | Flow 1 でスキップした予定を更新 | レコードが見つからず正常終了（何もしない） |
| 5 | Flow 1 でスキップした予定を削除 | レコードが見つからず正常終了（何もしない） |
| 6 | 予定の件名のみ変更（時刻変更なし） | Step 5 で時刻変更なしと判定。Teams通知のみ |
| 7 | 更新後にダッシュボード HTML を確認 | 更新後の時刻・工数で表示される |
| 8 | 削除後にダッシュボード HTML を確認 | 削除した工数が集計から除外されている |
