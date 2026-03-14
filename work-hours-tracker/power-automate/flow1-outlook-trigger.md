# Flow 1: Outlook予定 → Teams Adaptive Card通知

## フロー名
`[工数管理] Outlook予定 → プロジェクト選択通知`

## 概要
Outlookカレンダーに予定が追加されたとき（自分作成・他者から招待された両方）、
Microsoft Teams に Adaptive Card を送信してプロジェクトを選択させ、
選択されたプロジェクトの工数を SharePoint に記録する。

---

## トリガー

**コネクタ**: Office 365 Outlook
**アクション**: `予定が追加されたとき (V3)`
**設定**:
- カレンダーID: `予定表`（メインカレンダー）

> 他者から招待された予定を承諾すると、このトリガーも発火します。

---

## ステップ一覧

### Step 1: 終日イベントのスキップ
**条件分岐** アクション
- 条件: `triggerBody()?['isAllDay']` が `true` に等しい
- Yesの場合: **終了**（何もしない）
- Noの場合: 次のステップへ

---

### Step 2: 重複チェック（同一イベントの再登録防止）
**コネクタ**: SharePoint
**アクション**: `複数の項目の取得`
- サイトのアドレス: `[YOUR_SHAREPOINT_SITE_URL]`
- リスト名: `WorkHoursLog`
- フィルター クエリ: `EventId eq '@{triggerBody()?['id']}'`

**条件分岐**: 結果の `value` 配列の `length()` が `0` より大きい場合 → **終了**

---

### Step 3: 工数（時間数）の計算
**変数の初期化** アクション
- 名前: `DurationHours`
- 型: `浮動小数点数`
- 値（式）:
```
div(
  dateDifference(triggerBody()?['start'], triggerBody()?['end']),
  3600000
)
```
> `dateDifference` はミリ秒で返すため 3,600,000 で割って時間に変換

**代替式（分→時間）**:
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

### Step 4: プロジェクトリストの取得
**コネクタ**: SharePoint
**アクション**: `複数の項目の取得`
- サイトのアドレス: `[YOUR_SHAREPOINT_SITE_URL]`
- リスト名: `Projects`
- フィルター クエリ: `IsActive eq 1`
- 並び替え: `Title asc`

結果を変数 `ProjectsList` に格納

---

### Step 5: Adaptive Card の actions 配列を構築
**変数の初期化** アクション
- 名前: `actionsArray`
- 型: `アレイ`
- 値: `[]`（空の配列）

**Apply to each** アクション（`ProjectsList` の各行をループ）
内部で **アレイへの追加** アクションを実行:
- アレイ: `actionsArray`
- 値（JSON）:
```json
{
  "type": "Action.Submit",
  "title": "@{items('Apply_to_each')?['Title']}",
  "style": "positive",
  "data": {
    "action": "select_project",
    "projectCode": "@{items('Apply_to_each')?['ProjectCode']}",
    "projectName": "@{items('Apply_to_each')?['Title']}",
    "eventId": "@{triggerBody()?['id']}",
    "eventTitle": "@{triggerBody()?['subject']}",
    "startDateTime": "@{triggerBody()?['start']}",
    "endDateTime": "@{triggerBody()?['end']}",
    "durationHours": "@{variables('DurationHours')}"
  }
}
```

ループ後、**アレイへの追加** でスキップボタンを追加:
```json
{
  "type": "Action.Submit",
  "title": "⏭ スキップ（工数対象外）",
  "style": "destructive",
  "data": {
    "action": "skip",
    "eventId": "@{triggerBody()?['id']}"
  }
}
```

---

### Step 6: Adaptive Card JSON の構築
**Compose** アクション
- 名前: `AdaptiveCardJson`
- 入力（式で以下のJSONを構築）:

```json
{
  "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
  "type": "AdaptiveCard",
  "version": "1.4",
  "body": [
    {
      "type": "Container",
      "style": "emphasis",
      "items": [
        {
          "type": "TextBlock",
          "text": "📅 新しい予定が追加されました",
          "weight": "Bolder",
          "size": "Medium",
          "color": "Accent",
          "wrap": true
        }
      ]
    },
    {
      "type": "FactSet",
      "spacing": "Medium",
      "facts": [
        {
          "title": "予定名",
          "value": "@{triggerBody()?['subject']}"
        },
        {
          "title": "開始日時",
          "value": "@{triggerBody()?['start']}"
        },
        {
          "title": "終了日時",
          "value": "@{triggerBody()?['end']}"
        },
        {
          "title": "時間",
          "value": "@{variables('DurationHours')} 時間"
        }
      ]
    },
    {
      "type": "TextBlock",
      "text": "どのプロジェクトの工数ですか？",
      "weight": "Bolder",
      "spacing": "Medium",
      "wrap": true
    },
    {
      "type": "TextBlock",
      "text": "※ 24時間以内に選択してください。",
      "size": "Small",
      "color": "Warning",
      "isSubtle": true,
      "wrap": true
    }
  ],
  "actions": @{variables('actionsArray')}
}
```

---

### Step 7: Teams に Adaptive Card を投稿して応答を待つ
**コネクタ**: Microsoft Teams
**アクション**: `アダプティブ カードをチャットまたはチャネルに投稿して応答を待機する`
- 投稿者: `フロー ボット`
- 投稿先: `チャット` または `チャネル`（自分のチャットを推奨）
- 受信者: `[自分のメールアドレス]`
- アダプティブ カード: `outputs('AdaptiveCardJson')`
- タイムアウト: `PT24H`（24時間）

応答を変数 `TeamsResponse` に格納

---

### Step 8: スキップ判定
**条件分岐** アクション
- 条件: `body('Teams応答')?['data']?['action']` が `skip` に等しい
- Yesの場合: **終了**
- Noの場合: 次のステップへ

---

### Step 9: SharePoint に工数ログを登録
**コネクタ**: SharePoint
**アクション**: `項目の作成`
- サイトのアドレス: `[YOUR_SHAREPOINT_SITE_URL]`
- リスト名: `WorkHoursLog`
- 各フィールドの値:

| フィールド | 式 |
|-----------|-----|
| Title | `body('Teams応答')?['data']?['eventTitle']` |
| EventId | `body('Teams応答')?['data']?['eventId']` |
| StartDateTime | `body('Teams応答')?['data']?['startDateTime']` |
| EndDateTime | `body('Teams応答')?['data']?['endDateTime']` |
| DurationHours | `float(body('Teams応答')?['data']?['durationHours'])` |
| ProjectCode | `body('Teams応答')?['data']?['projectCode']` |
| ProjectName | `body('Teams応答')?['data']?['projectName']` |
| RecordedAt | `utcNow()` |
| Status | `Active` |
| LastSyncedAt | `utcNow()` |

---

### Step 10: HTML ダッシュボード生成フローを呼び出す
**コネクタ**: Power Automate
**アクション**: `フローを実行する`（または HTTP 要求）
- Flow 2（`[工数管理] HTML ダッシュボード生成`）を呼び出す

---

## エラーハンドリング

| 状況 | 対処 |
|------|------|
| Teams応答タイムアウト（24h） | フローは自動終了。工数は未登録のまま。 |
| DurationHours が 0 以下 | Step 1.5 で条件分岐を追加し、0h以下はスキップ推奨 |
| 同一EventIdが既に登録済み | Step 2 の重複チェックで検出し終了 |

---

## テスト手順

1. Outlookで30分の予定を作成する
2. Teams に Adaptive Card が届くことを確認する
3. 「プロジェクトA」ボタンをクリックする
4. SharePoint の WorkHoursLog リストに 0.5h のレコードが作成されることを確認する
5. WorkHoursDashboard/index.html がSharePointに生成/更新されることを確認する
