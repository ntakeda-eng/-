# デスクトップ整理スクリプト (Windows)

デスクトップ上のファイルを整理する PowerShell スクリプトです。用途に応じて2種類あります。

- **`Organize-Desktop.ps1`** … ファイルを **種別ごと**（画像・PDF・文書…）に分ける
- **`Sort-DesktopByUsage.ps1`** … ファイルを **使用中／未使用** に分ける（下記参照）

---

## 使用中／未使用で分ける: `Sort-DesktopByUsage.ps1`

各ファイルの「最終使用日時」（最終アクセス日時と最終更新日時のうち新しい方）を見て、
一定期間より古いものを `デスクトップ\未使用\` へ、新しいものを `デスクトップ\使用中\` へ振り分けます。

```powershell
# まずプレビュー（何も移動しません／既定は90日基準）
powershell -ExecutionPolicy Bypass -File .\Sort-DesktopByUsage.ps1 -WhatIf

# 本番実行
powershell -ExecutionPolicy Bypass -File .\Sort-DesktopByUsage.ps1

# 基準日数を変える（例: 30日）
powershell -ExecutionPolicy Bypass -File .\Sort-DesktopByUsage.ps1 -Days 30
```

**注意:** Windows は既定で「最終アクセス日時」の更新を無効にしている場合があります。
その場合は「最終更新日時」で判定されるため、「開いただけ」のファイルは未使用側に分類されることがあります。

---

## 種別ごとに分ける: `Organize-Desktop.ps1`

## 何をするか

デスクトップ直下のファイルを拡張子で判別し、`デスクトップ\整理済み\<カテゴリ>\` に振り分けます。

| カテゴリ | 対象の例 |
|---|---|
| 画像 | jpg / png / gif / heic / svg など |
| PDF | pdf |
| 文書 | doc / docx / txt / md など |
| 表計算 | xls / xlsx / csv |
| プレゼン | ppt / pptx |
| 圧縮ファイル | zip / rar / 7z |
| 音声 | mp3 / wav / m4a など |
| 動画 | mp4 / mov / mkv など |
| インストーラ | exe / msi |
| コード | py / js / html / ps1 など |
| その他 | 上記以外すべて |

- **フォルダは動かしません**（中身も触りません）
- **ショートカット (.lnk / .url) は既定で移動しません**（デスクトップに残します）
- **同名ファイルは自動で連番リネーム**して安全に移動します
- 実行内容は `整理ログ_<日時>.txt` に保存されます

## 使い方

1. `Organize-Desktop.ps1` を自分のPCにダウンロードする
2. PowerShell を開く（スタートメニューで「PowerShell」と検索）
3. スクリプトのある場所へ移動（例）:
   ```powershell
   cd $env:USERPROFILE\Downloads
   ```
4. **まずプレビュー**で動作確認（何も移動しません）:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\Organize-Desktop.ps1 -WhatIf
   ```
5. 問題なければ **本番実行**:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\Organize-Desktop.ps1
   ```

## オプション

- ショートカットも整理対象に含めたい場合:
  ```powershell
  powershell -ExecutionPolicy Bypass -File .\Organize-Desktop.ps1 -IncludeShortcuts
  ```

## カスタマイズ

カテゴリや対象拡張子を変えたい場合は、`Organize-Desktop.ps1` の `$CategoryMap` を編集してください。

---

# ACES Meet → Slack 要約スクリプト: `Aces-MeetingSummary-ToSlack.ps1`

ACES Meet API から会議データ（AI要約・書き起こし）を取得し、要約を Slack に投稿します。
上のデスクトップ整理系スクリプトとは別用途の単独ユーティリティです。

> **実行場所について:** ACES / Slack に到達できる環境（あなたの手元の Windows マシンや社内サーバー）で
> 実行してください。ネットワークが制限されたクラウド実行環境からは `acesinc.co.jp` に接続できない場合があります。

## 事前準備

1. ACES Meet API キーを取得（管理者 / API 設定より）
2. ACES Meet API ドキュメントで「ベースURL」を確認
3. Slack の Incoming Webhook URL を用意（または Bot トークン）
4. 環境変数をセット:
   ```powershell
   $env:ACES_MEET_API_KEY  = "xxxxxxxx"
   $env:ACES_MEET_API_BASE = "https://<APIベースURL>"
   $env:SLACK_WEBHOOK_URL  = "https://hooks.slack.com/services/XXX/YYY/ZZZ"
   ```

## 使い方

```powershell
# 1) まず API レスポンスの構造を確認（Slack へは投稿しない）
powershell -ExecutionPolicy Bypass -File .\Aces-MeetingSummary-ToSlack.ps1 -Discover

# 2) 直近3件の内容だけ確認（投稿しない）
powershell -ExecutionPolicy Bypass -File .\Aces-MeetingSummary-ToSlack.ps1 -Limit 3 -DryRun

# 3) 全会議を要約して Slack に投稿
powershell -ExecutionPolicy Bypass -File .\Aces-MeetingSummary-ToSlack.ps1

# 4) 期間を絞る（例: 7/1 以降）
powershell -ExecutionPolicy Bypass -File .\Aces-MeetingSummary-ToSlack.ps1 -Since 2026-07-01
```

## 重要: フィールドの対応付け

API ドキュメントに直接アクセスできなかったため、スクリプト上部の `$Config` ブロックには
**一般的な REST 慣習に基づく既定値**（エンドポイントのパス・認証ヘッダ形式・フィールド名）を
置いています。**必ず最初に `-Discover` を実行**して実際のレスポンス構造を確認し、
`$Config` の以下を実データに合わせて調整してください。

| 設定 | 意味 |
|---|---|
| `AuthHeaderName` / `AuthValueFormat` | 認証ヘッダ名と値の形式（`Bearer {0}` など） |
| `ListPath` / `DetailPath` | 会議一覧・詳細のエンドポイントパス |
| `Field_*` | タイトル・日時・要約・書き起こし等の実フィールド名 |

## オプション一覧

| パラメータ | 説明 |
|---|---|
| `-Discover` | 一覧/詳細の生JSONを表示（マッピング確認用） |
| `-DryRun` | Slack投稿せずメッセージ内容を表示 |
| `-Since` / `-Until` | 対象会議の期間フィルタ |
| `-Limit` | 処理する最大件数（0=全件） |
| `-SlackChannel` | Bot トークン利用時の投稿先チャンネル |
