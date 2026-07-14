# 業務効率化スクリプト集 (Windows / Microsoft 365)

PowerShell スクリプト集です。

- **`New-OutlookReplyDrafts.ps1`** … Outlook を仕分けし、**返信すべきメールだけに返信の下書きを自動作成**（下記参照）
- **`Organize-Desktop.ps1`** … デスクトップのファイルを **種別ごと**（画像・PDF・文書…）に分ける
- **`Sort-DesktopByUsage.ps1`** … デスクトップのファイルを **使用中／未使用** に分ける（下記参照）

---

## Outlook 返信下書きの自動作成: `New-OutlookReplyDrafts.ps1`

受信トレイを仕分けし、**返信が必要なメールにだけ返信の下書きを Outlook 上に作成**します。
**送信は一切自動化しません**。作られるのは「下書き」だけなので、内容を確認・手直しして、
送信ボタンはあなたが押します（誤送信ゼロ）。

### 仕分けの考え方（返信不要メールの除外）

次のルールで自動的にふるい分け、残ったものだけ下書きを作ります。

| 除外するもの | 判定方法 |
|---|---|
| 自動送信・配信専用 | 差出人が `noreply@` 等 / 本文に「配信停止・セミナー・アンケート」等 |
| 自分が CC のみ | 主宛（To）が別の人 → その人が主担当なので下書きを作らない |
| **知らない業者・新規営業** | **送信済みメールから作った「既知の取引先リスト」に載っていない差出人** |
| （フリーメール） | gmail 等はアドレス完全一致のときだけ対象 |
| 処理済み | 一度下書きを作ったメールにはカテゴリを付け、二重作成を防止 |

「既知の取引先」は、あなたが過去に返信したことのある相手（送信済みトレイ）から
自動生成されるため、リストを手で作る必要はありません。

### 返信本文

内容の種別（見積受領／資料受領／依頼・質問／汎用）に応じたテンプレート定型文を、
相手の名前を差し込んで挿入します。依頼・質問型には「回答を書き足す箇所」を明示します。

### 事前準備（初回のみ・管理者権限不要）

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

初回実行時に、ブラウザで会社アカウント（`n.takeda@willows-group.com`）にログインし、
`Mail.ReadWrite` の権限に同意します（以降はログインが維持されます）。

### 使い方

```powershell
# まずプレビュー（下書きは作らず、仕分け結果だけ表示）
powershell -ExecutionPolicy Bypass -File .\New-OutlookReplyDrafts.ps1 -WhatIf

# 本番実行（直近3日・社外の既知取引先のみに下書き作成）
powershell -ExecutionPolicy Bypass -File .\New-OutlookReplyDrafts.ps1

# 未読のみ・直近7日を対象
powershell -ExecutionPolicy Bypass -File .\New-OutlookReplyDrafts.ps1 -LookbackDays 7 -UnreadOnly

# 新しい取引先を反映（取引先リストを作り直す）
powershell -ExecutionPolicy Bypass -File .\New-OutlookReplyDrafts.ps1 -RebuildAllowlist
```

### 主なオプション

| オプション | 意味 |
|---|---|
| `-LookbackDays <日数>` | 受信トレイの何日前までを対象にするか（既定: 3） |
| `-UnreadOnly` | 未読メールだけを対象にする |
| `-IncludeInternal` | 社内（willows-group.com 等）宛も対象に含める（既定は社外のみ） |
| `-MaxDrafts <件数>` | 1回に作る下書きの上限（既定: 20） |
| `-RebuildAllowlist` | 既知の取引先リストを作り直す |
| `-WhatIf` | 実際には作らず、対象だけ表示（プレビュー） |

### カスタマイズ

`New-OutlookReplyDrafts.ps1` の冒頭「設定」ブロックや `Get-ReplyTemplate` 関数を編集すると、
署名・テンプレート文面・除外キーワード・社内ドメインなどを調整できます。

> **メモ:** 本スクリプトはメールの**読み取りと下書き作成**のみを行い、送信は行いません。
> 下書きは Outlook の［下書き］フォルダに入ります。

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
