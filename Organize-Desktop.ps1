<#
.SYNOPSIS
    デスクトップのファイルを種別ごとにフォルダ分けして整理します。

.DESCRIPTION
    デスクトップ直下のファイルを拡張子で判別し、
    "整理済み\<カテゴリ>" フォルダへ振り分けます。
      - フォルダ自体は移動しません（中身は触りません）
      - ショートカット (.lnk / .url) は既定で移動しません
      - 同名ファイルがある場合は自動で連番を付けてリネーム
      - 実行内容はログファイルに記録されます

.PARAMETER WhatIf
    実際には移動せず、何が起きるかだけを表示します（プレビュー）。
    まずはこれを付けて実行し、問題なければ外して本番実行してください。

.PARAMETER IncludeShortcuts
    指定するとショートカット (.lnk / .url) も整理対象に含めます。

.EXAMPLE
    # プレビュー（何も移動しない）
    powershell -ExecutionPolicy Bypass -File .\Organize-Desktop.ps1 -WhatIf

.EXAMPLE
    # 本番実行
    powershell -ExecutionPolicy Bypass -File .\Organize-Desktop.ps1
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [switch]$IncludeShortcuts
)

# --- 整理先の親フォルダ名 --------------------------------------------------
$OrganizedFolderName = '整理済み'

# --- 拡張子 → カテゴリの対応表 ---------------------------------------------
$CategoryMap = @{
    '画像'           = @('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tif', '.tiff', '.webp', '.heic', '.svg', '.ico')
    'PDF'            = @('.pdf')
    '文書'           = @('.doc', '.docx', '.txt', '.rtf', '.odt', '.md', '.pages')
    '表計算'         = @('.xls', '.xlsx', '.csv', '.ods')
    'プレゼン'       = @('.ppt', '.pptx', '.odp', '.key')
    '圧縮ファイル'   = @('.zip', '.rar', '.7z', '.tar', '.gz', '.bz2', '.xz')
    '音声'           = @('.mp3', '.wav', '.flac', '.aac', '.m4a', '.ogg', '.wma')
    '動画'           = @('.mp4', '.mov', '.avi', '.mkv', '.wmv', '.flv', '.webm')
    'インストーラ'   = @('.exe', '.msi')
    'コード'         = @('.py', '.js', '.ts', '.html', '.css', '.json', '.xml', '.java', '.c', '.cpp', '.cs', '.go', '.rb', '.php', '.sh', '.ps1', '.bat')
    'ショートカット' = @('.lnk', '.url')
}

# --- デスクトップの実パスを取得（OneDrive リダイレクトにも対応）------------
$DesktopPath = [Environment]::GetFolderPath('Desktop')
if (-not (Test-Path -LiteralPath $DesktopPath)) {
    Write-Error "デスクトップのパスが見つかりませんでした: $DesktopPath"
    return
}

$OrganizedRoot = Join-Path $DesktopPath $OrganizedFolderName

# --- ログの準備 ------------------------------------------------------------
$Stamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
$LogPath = Join-Path $DesktopPath ("整理ログ_{0}.txt" -f $Stamp)
$Log     = [System.Collections.Generic.List[string]]::new()

function Write-Log {
    param([string]$Message)
    $line = "{0}  {1}" -f (Get-Date -Format 'HH:mm:ss'), $Message
    $Log.Add($line)
    Write-Host $line
}

Write-Log "対象デスクトップ: $DesktopPath"
if ($WhatIfPreference) { Write-Log "※ プレビュー(-WhatIf)モード: 実際には移動しません" }

# 拡張子からカテゴリ名を引く
function Get-Category {
    param([string]$Extension)
    $ext = $Extension.ToLower()
    foreach ($cat in $CategoryMap.Keys) {
        if ($CategoryMap[$cat] -contains $ext) { return $cat }
    }
    return 'その他'
}

# 移動先で重複しない一意なパスを作る
function Get-UniquePath {
    param([string]$Directory, [string]$FileName)
    $target = Join-Path $Directory $FileName
    if (-not (Test-Path -LiteralPath $target)) { return $target }

    $base = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $ext  = [System.IO.Path]::GetExtension($FileName)
    $n    = 1
    do {
        $candidate = Join-Path $Directory ("{0} ({1}){2}" -f $base, $n, $ext)
        $n++
    } while (Test-Path -LiteralPath $candidate)
    return $candidate
}

# --- 整理対象ファイルの収集 -----------------------------------------------
# デスクトップ直下のファイルのみ。フォルダ・スクリプト自身・ログは除外。
$scriptName = Split-Path -Leaf $PSCommandPath

$files = Get-ChildItem -LiteralPath $DesktopPath -File | Where-Object {
    $_.Name -ne $scriptName -and
    $_.Name -notlike '整理ログ_*.txt'
}

if (-not $IncludeShortcuts) {
    $files = $files | Where-Object { $_.Extension.ToLower() -notin @('.lnk', '.url') }
}

if (-not $files -or $files.Count -eq 0) {
    Write-Log "整理対象のファイルはありませんでした。"
    $Log | Set-Content -LiteralPath $LogPath -Encoding UTF8
    return
}

Write-Log "整理対象: $($files.Count) 件"

# --- 振り分け実行 ----------------------------------------------------------
$moved = 0
foreach ($file in $files) {
    $category = Get-Category $file.Extension
    $destDir  = Join-Path $OrganizedRoot $category

    if (-not (Test-Path -LiteralPath $destDir)) {
        if ($PSCmdlet.ShouldProcess($destDir, 'フォルダ作成')) {
            New-Item -ItemType Directory -Path $destDir -Force | Out-Null
        }
    }

    $destPath = Get-UniquePath -Directory $destDir -FileName $file.Name

    if ($PSCmdlet.ShouldProcess($file.FullName, "移動 → $category")) {
        try {
            Move-Item -LiteralPath $file.FullName -Destination $destPath -ErrorAction Stop
            Write-Log ("[{0}] {1}" -f $category, $file.Name)
            $moved++
        }
        catch {
            Write-Log ("[失敗] {0} : {1}" -f $file.Name, $_.Exception.Message)
        }
    }
    else {
        Write-Log ("[予定:{0}] {1}" -f $category, $file.Name)
    }
}

Write-Log "完了: $moved 件を移動しました。"

# ログ保存（プレビュー時も記録を残す）
try {
    $Log | Set-Content -LiteralPath $LogPath -Encoding UTF8
    Write-Host "`nログを保存しました: $LogPath"
}
catch {
    Write-Host "ログの保存に失敗しました: $($_.Exception.Message)"
}
