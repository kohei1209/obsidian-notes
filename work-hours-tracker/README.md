# 工数管理システム セットアップガイド

## システム概要

Outlookに予定が追加されるたびに、Microsoft TeamsにAdaptive Cardが届き、
プロジェクトを選択するだけで工数をSharePointに自動記録します。
記録されたデータは、HTMLダッシュボードで視覚的に確認できます。

```
Outlook予定追加
    ↓
Power Automate Flow 1
    ↓
Teams Adaptive Card（プロジェクト選択）
    ↓
SharePoint WorkHoursLog に記録
    ↓
Power Automate Flow 2
    ↓
index.html を自動生成 → SharePoint に保存
    ↓
ダウンロードしてブラウザで閲覧
```

---

## 事前準備

### 必要なもの
- Microsoft 365 ライセンス（Power Automate 標準、Teams、SharePoint、Outlook を含む）
- SharePoint サイトへのアクセス権（サイトオーナーまたは編集者）

---

## セットアップ手順

### Step 1: SharePoint の設定

#### 1-1. Projects リストの作成

1. SharePoint サイトを開く
2. **サイトコンテンツ** → **新規** → **リスト**
3. リスト名: `Projects`
4. 以下の列を追加:

| 列名 | 型 | 必須 |
|------|-----|-----|
| Title（既定） | テキスト(1行) | ✓ |
| ProjectCode | テキスト(1行) | ✓ |
| Color | テキスト(1行) | ✓ |
| PlannedHours | 数値 | |
| IsActive | はい/いいえ | ✓ |

5. サンプルデータを登録（`sharepoint/schema.json` の `sampleData` 参照）

#### 1-2. WorkHoursLog リストの作成

1. **サイトコンテンツ** → **新規** → **リスト**
2. リスト名: `WorkHoursLog`
3. 以下の列を追加:

| 列名 | 型 | 必須 |
|------|-----|-----|
| Title（既定） | テキスト(1行) | ✓ |
| EventId | テキスト(1行) | ✓ |
| StartDateTime | 日時 | ✓ |
| EndDateTime | 日時 | ✓ |
| DurationHours | 数値 | ✓ |
| ProjectCode | テキスト(1行) | ✓ |
| ProjectName | テキスト(1行) | ✓ |
| RecordedAt | 日時 | ✓ |

#### 1-3. WorkHoursDashboard ライブラリの作成

1. **サイトコンテンツ** → **新規** → **ドキュメントライブラリ**
2. ライブラリ名: `WorkHoursDashboard`
3. `_template` フォルダを作成して `dashboard/index.html` をアップロード

---

### Step 2: Power Automate Flow 2（HTML生成）の設定

> Flow 1 より先に Flow 2 を設定してください。

1. [power.automate.com](https://make.powerautomate.com) を開く
2. **新しいフロー** → **インスタントクラウドフロー**
3. フロー名: `[工数管理] HTML ダッシュボード生成`
4. トリガー: **HTTP 要求を受信したとき**
5. 詳細な手順は `power-automate/flow2-generate-html.md` を参照
6. フローを保存し、**HTTP POST URL** をコピーしておく

---

### Step 3: Power Automate Flow 1（Outlook通知）の設定

1. **新しいフロー** → **自動化クラウドフロー**
2. フロー名: `[工数管理] Outlook予定 → プロジェクト選択通知`
3. トリガー: **Office 365 Outlook** - **予定が追加されたとき (V3)**
4. 詳細な手順は `power-automate/flow1-outlook-trigger.md` を参照
5. Step 10 で Flow 2 の HTTP URL を指定する

---

### Step 4: Power Apps プロジェクト管理アプリの作成（オプション）

プロジェクトの追加・編集・削除を GUI で行いたい場合:

1. [make.powerapps.com](https://make.powerapps.com) を開く
2. **アプリ** → **新しいアプリ** → **キャンバス**
3. データ接続: SharePoint → `Projects` リスト
4. 以下の画面を作成:
   - プロジェクト一覧（Gallery コントロール）
   - プロジェクト編集（Form コントロール）
5. `Patch(Projects, ...)` 関数で保存処理を実装

> Power Apps なしでも、SharePoint リストを直接編集してプロジェクトを管理できます。

---

## ダッシュボードの閲覧方法

1. SharePoint サイトを開く
2. `WorkHoursDashboard` ドキュメントライブラリを開く
3. `index.html` の右クリックメニューから **ダウンロード**
4. ダウンロードしたファイルをダブルクリックしてブラウザで開く

> インターネット接続が必要です（Chart.js を CDN から読み込むため）

---

## ダッシュボードの機能

| 機能 | 説明 |
|------|------|
| 🥧 円グラフ（今週/今月/今年） | 期間別のプロジェクト別工数と割合を表示 |
| 📈 棒グラフ | 週別または月別の工数推移（プロジェクト別積み上げ） |
| 🎯 進捗率 | 予定工数に対する実績工数の達成率 |
| 📋 ログ一覧 | 全工数記録の詳細テーブル（日付降順） |

---

## ファイル構成

```
work-hours-tracker/
├── README.md                          # このファイル
├── sharepoint/
│   └── schema.json                    # SharePointリスト・ライブラリの定義
├── adaptive-card/
│   └── event-notification.json        # Teams Adaptive Card テンプレート
├── power-automate/
│   ├── flow1-outlook-trigger.md       # Flow1の設定手順
│   └── flow2-generate-html.md         # Flow2の設定手順
└── dashboard/
    └── index.html                      # HTMLダッシュボード
                                        # （このファイルをSharePointの
                                        #   _templateフォルダに配置する）
```

---

## よくある質問

**Q: 招待された予定にも反応しますか？**
A: はい。Outlookで会議招待を承諾するとカレンダーに追加されるため、Flow 1 が発火します。

**Q: 終日予定はどうなりますか？**
A: Flow 1 で自動的にスキップされます（工数0時間のため）。

**Q: プロジェクトを追加したい場合は？**
A: SharePoint の `Projects` リストに行を追加し、`IsActive` を `はい` に設定してください。次回の Adaptive Card から新しいプロジェクトボタンが表示されます。

**Q: 同じ予定が更新されたときに重複登録されませんか？**
A: Flow 1 の Step 2 で EventId を確認し、既に登録済みの場合はスキップします。

**Q: 過去の工数を手動で追加できますか？**
A: SharePoint の `WorkHoursLog` リストに直接行を追加することで可能です。その後 Flow 2 を手動実行してダッシュボードを更新してください。

---

## トラブルシューティング

| 症状 | 確認事項 |
|------|---------|
| Teams通知が届かない | Flow 1 の実行履歴を確認。Outlookトリガーが発火しているか確認。 |
| プロジェクトボタンが表示されない | Projects リストに IsActive=true の項目があるか確認。 |
| HTMLが更新されない | Flow 2 の実行履歴を確認。SharePointへの書き込み権限を確認。 |
| グラフが表示されない | インターネット接続を確認（Chart.js CDN への接続が必要）。 |
