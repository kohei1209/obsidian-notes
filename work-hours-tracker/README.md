# 工数管理システム セットアップガイド

## システム概要

Outlookカレンダーの予定に**カテゴリを設定するだけ**で、工数がSharePointに自動記録されます。
予定の変更・削除も自動で同期されます。
記録されたデータは、HTMLダッシュボードで視覚的に確認できます。

```
Outlook予定追加（カテゴリ設定済み）
    ↓
Power Automate Flow 1（カテゴリ → プロジェクト自動判定）
    ↓
SharePoint WorkHoursLog に記録
    ↓
Power Automate Flow 2（HTML 自動生成）
    ↓
index.html → SharePoint に保存
    ↓
ダウンロードしてブラウザで閲覧
```

### 予定の変更・削除も自動同期
```
Outlook予定を変更（時刻変更、カテゴリ変更等）
    ↓
Power Automate Flow 3a（変更検知 → SharePoint 更新 → HTML 再生成）

Outlook予定を削除
    ↓
Power Automate Flow 3b（削除検知 → Status=Cancelled → HTML 再生成）
```

---

## 事前準備

### 必要なもの
- Microsoft 365 ライセンス（Power Automate 標準、Teams、SharePoint、Outlook を含む）
- SharePoint サイトへのアクセス権（サイトオーナーまたは編集者）

---

## セットアップ手順

### Step 1: Outlook カテゴリの作成

Outlook（デスクトップまたは Web）でプロジェクトに対応するカテゴリを作成する:

1. Outlook を開く → **分類** → **すべてのカテゴリ**
2. 以下のカテゴリを作成（色は自由）:

| カテゴリ名 | 用途 |
|-----------|------|
| プロジェクトA | プロジェクトA の工数 |
| プロジェクトB | プロジェクトB の工数 |
| プロジェクトC | プロジェクトC の工数 |

> カテゴリ名は後で SharePoint の Projects リストに登録する
> `OutlookCategory` 列と**完全一致**させること。

---

### Step 2: SharePoint の設定

#### 2-1. Projects リストの作成

1. SharePoint サイトを開く
2. **サイトコンテンツ** → **新規** → **リスト**
3. リスト名: `Projects`
4. 以下の列を追加:

| 列名 | 型 | 必須 | 備考 |
|------|-----|-----|------|
| Title（既定） | テキスト(1行) | Yes | プロジェクト名 |
| ProjectCode | テキスト(1行) | Yes | 英数字・アンダースコアのみ（例: proj_a） |
| OutlookCategory | テキスト(1行) | Yes | Outlook カテゴリ名と完全一致させる |
| Color | テキスト(1行) | Yes | HEX形式（例: #FF6384） |
| PlannedHours | 数値 | No | 年度ごとの予定工数 |
| IsActive | はい/いいえ | Yes | 初期値: はい |

5. サンプルデータを登録（`sharepoint/schema.json` の `sampleData` 参照）

#### 2-2. WorkHoursLog リストの作成

1. **サイトコンテンツ** → **新規** → **リスト**
2. リスト名: `WorkHoursLog`
3. 以下の列を追加:

| 列名 | 型 | 必須 | 備考 |
|------|-----|-----|------|
| Title（既定） | テキスト(1行) | Yes | 予定名（自動入力） |
| EventId | テキスト(1行) | Yes | Outlook イベント ID（自動入力） |
| StartDateTime | 日時 | Yes | |
| EndDateTime | 日時 | Yes | |
| DurationHours | 数値 | Yes | |
| ProjectCode | テキスト(1行) | No | 未分類の場合は空 |
| ProjectName | テキスト(1行) | No | 未分類の場合は空 |
| RecordedAt | 日時 | Yes | |
| Status | 選択肢 | Yes | 選択肢: Active, Cancelled（初期値: Active） |
| LastSyncedAt | 日時 | No | |

#### 2-3. WorkHoursDashboard ライブラリの作成

1. **サイトコンテンツ** → **新規** → **ドキュメントライブラリ**
2. ライブラリ名: `WorkHoursDashboard`
3. `_template` フォルダを作成して `index.html` をアップロード

---

### Step 3: Power Automate の設定

> Flow 2 → Flow 1 → Flow 3 の順に設定してください。

#### Flow 2: HTML ダッシュボード生成
1. **新しいフロー** → **インスタントクラウドフロー**
2. フロー名: `[工数管理] HTML ダッシュボード生成`
3. トリガー: **HTTP 要求を受信したとき**
4. 詳細: `power-automate/flow2-generate-html.md` を参照
5. 保存後、**HTTP POST URL** をコピーしておく

#### Flow 1: Outlook予定 → 工数自動記録
1. **新しいフロー** → **自動化クラウドフロー**
2. フロー名: `[工数管理] Outlook予定 → 工数自動記録`
3. トリガー: **Office 365 Outlook** - **予定が追加されたとき (V3)**
4. 詳細: `power-automate/flow1-outlook-trigger.md` を参照

#### Flow 3a: 予定更新 → 工数ログ同期
1. **新しいフロー** → **自動化クラウドフロー**
2. フロー名: `[工数管理] Outlook予定更新 → 工数ログ同期`
3. トリガー: **Office 365 Outlook** - **予定が更新されたとき (V3)**
4. 詳細: `power-automate/flow3-event-sync.md` を参照

#### Flow 3b: 予定削除 → 工数ログ同期
1. Flow 3a をコピーし、トリガーのみ **予定が削除されたとき (V3)** に変更
2. 詳細: `power-automate/flow3-event-sync.md`（削除版のステップを参照）

---

## 使い方

### 工数の記録
1. Outlook で予定を作成する
2. 予定にカテゴリを設定する（例: 「プロジェクトA」）
3. 完了！工数が自動的に SharePoint に記録され、ダッシュボードが更新される

### カテゴリを設定し忘れた場合
- 「未分類」として記録され、Teams に警告通知が届く
- 後から Outlook でカテゴリを設定すると、Flow 3a が発火してプロジェクトが自動更新される
- または SharePoint の WorkHoursLog リストで直接編集可能

### ダッシュボードの閲覧
1. SharePoint サイトを開く
2. `WorkHoursDashboard` ドキュメントライブラリを開く
3. `index.html` の右クリックメニューから **ダウンロード**
4. ダウンロードしたファイルをダブルクリックしてブラウザで開く

> インターネット接続が必要（Chart.js を CDN から読み込むため）

---

## ダッシュボードの機能

| 機能 | 説明 |
|------|------|
| 円グラフ（今週/今月/今年度） | 期間別のプロジェクト別工数と割合を表示 |
| 棒グラフ | 週別/月別/年度別の工数推移（プロジェクト別積み上げ） |
| 進捗率 | 予定工数に対する実績工数の達成率 |
| ログ一覧 | 全工数記録の詳細テーブル（日付降順、削除済みは薄く表示） |

---

## ファイル構成

```
work-hours-tracker/
├── README.md                          # このファイル
├── sharepoint/
│   └── schema.json                    # SharePoint リスト・ライブラリの定義
├── power-automate/
│   ├── flow1-outlook-trigger.md       # Flow 1: Outlook → 工数自動記録
│   ├── flow2-generate-html.md         # Flow 2: HTML ダッシュボード生成
│   └── flow3-event-sync.md            # Flow 3: 予定変更・削除 → 工数同期
└── index.html                         # HTML ダッシュボードテンプレート
                                       # （SharePoint の _template フォルダに配置する）
```

---

## よくある質問

**Q: 招待された予定にも反応しますか？**
A: はい。Outlook で会議招待を承諾するとカレンダーに追加されるため、Flow 1 が発火します。

**Q: 終日予定はどうなりますか？**
A: Flow 1 で自動的にスキップされます（工数対象外）。

**Q: プロジェクトを追加したい場合は？**
A: (1) SharePoint の Projects リストに行を追加し、OutlookCategory を設定する。(2) Outlook で同名のカテゴリを作成する。

**Q: カテゴリを後から変更したらどうなりますか？**
A: Flow 3a が発火し、新しいカテゴリに対応するプロジェクトに自動更新されます。

**Q: 同じ予定が重複登録されませんか？**
A: Flow 1 の Step 2 で EventId を確認し、既に登録済みの場合はスキップします。

**Q: 予定を削除したらどうなりますか？**
A: Flow 3b が発火し、Status が Cancelled に設定されます（物理削除ではないため復旧可能）。

**Q: 過去の工数を手動で追加できますか？**
A: SharePoint の WorkHoursLog リストに直接行を追加し、Flow 2 を手動実行してダッシュボードを更新してください。

---

## トラブルシューティング

| 症状 | 確認事項 |
|------|---------|
| 工数が記録されない | Flow 1 の実行履歴を確認。Outlook トリガーが発火しているか確認。 |
| プロジェクトが「未分類」になる | Outlook のカテゴリ名と Projects リストの OutlookCategory が完全一致しているか確認。 |
| 予定を更新しても反映されない | Flow 3a の実行履歴を確認。Status フィルターが正しく設定されているか確認。 |
| HTML が更新されない | Flow 2 の実行履歴を確認。SharePoint への書き込み権限を確認。 |
| グラフが表示されない | インターネット接続を確認（Chart.js CDN への接続が必要）。 |
| Cancelled レコードが集計に含まれる | Flow 2 の Select アクションで Status フィールドが正しく参照されているか確認（Choice 列の場合は `?['Value']` が必要）。 |
