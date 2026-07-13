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
