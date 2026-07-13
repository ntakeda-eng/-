# デスクトップ整理スクリプト (Windows)

デスクトップ上のファイルを **種別ごと** にフォルダ分けする PowerShell スクリプトです。

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
