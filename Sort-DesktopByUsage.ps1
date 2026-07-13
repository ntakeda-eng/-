<#
.SYNOPSIS
    デスクトップのファイルを「使用中」と「未使用」に分けて整理します。

.DESCRIPTION
    各ファイルの「最終使用日時」（最終アクセス日時と最終更新日時のうち
    新しい方）を見て、しきい値の日数より古いものを「未使用」、
    新しいものを「使用中」フォルダへ振り分けます。

      - フォルダ自体は移動しません（中身は触りません）
      - ショートカット (.lnk / .url) は既定で移動しません
      - 同名ファイルがある場合は自動で連番リネーム
      - 実行内容はログファイルに記録されます

    ※注意: Windows は既定で「最終アクセス日時」の更新を無効にしている
      場合があります。その場合は「最終更新日時」で判定されるため、
      「開いただけ」のファイルは未使用と判定されることがあります。

.PARAMETER Days
    この日数より長く使われていないファイルを「未使用」とします（既定: 90）。

.PARAMETER WhatIf
    実際には移動せず、何が起きるかだけを表示します（プレビュー）。
    まずはこれを付けて実行し、問題なければ外して本番実行してください。

.PARAMETER IncludeShortcuts
    指定するとショートカット (.lnk / .url) も対象に含めます。

.EXAMPLE
    # プレビュー（何も移動しない）
    powershell -ExecutionPolicy Bypass -File .\Sort-DesktopByUsage.ps1 -WhatIf

.EXAMPLE
    # 本番実行（90日基準）
    powershell -ExecutionPolicy Bypass -File .\Sort-DesktopByUsage.ps1

.EXAMPLE
    # 30日基準にする
    powershell -ExecutionPolicy Bypass -File .\Sort-DesktopByUsage.ps1 -Days 30
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [int]$Days = 90,
    [switch]$IncludeShortcuts
)

# --- 振り分け先フォルダ名 --------------------------------------------------
$UsedFolderName   = '使用中'
$UnusedFolderName = '未使用'

# --- デスクトップの実パスを取得（OneDrive リダイレクトにも対応）------------
$DesktopPath = [Environment]::GetFolderPath('Desktop')
if (-not (Test-Path -LiteralPath $DesktopPath)) {
    Write-Error "デスクトップのパスが見つかりませんでした: $DesktopPath"
    return
}

$UsedRoot   = Join-Path $DesktopPath $UsedFolderName
$UnusedRoot = Join-Path $DesktopPath $UnusedFolderName

# --- しきい値の日時 --------------------------------------------------------
$Threshold = (Get-Date).AddDays(-$Days)

# --- ログの準備 ------------------------------------------------------------
$Stamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
$LogPath = Join-Path $DesktopPath ("使用状況ログ_{0}.txt" -f $Stamp)
$Log     = [System.Collections.Generic.List[string]]::new()

function Write-Log {
    param([string]$Message)
    $line = "{0}  {1}" -f (Get-Date -Format 'HH:mm:ss'), $Message
    $Log.Add($line)
    Write-Host $line
}

Write-Log "対象デスクトップ: $DesktopPath"
Write-Log ("基準: 直近 {0} 日以内に使用 = 使用中 / それより古い = 未使用" -f $Days)
Write-Log ("しきい日時: {0} 以前を未使用と判定" -f $Threshold.ToString('yyyy-MM-dd HH:mm'))
if ($WhatIfPreference) { Write-Log "※ プレビュー(-WhatIf)モード: 実際には移動しません" }

# 最終使用日時（アクセスと更新の新しい方）
function Get-LastUsed {
    param([System.IO.FileInfo]$File)
    if ($File.LastAccessTime -gt $File.LastWriteTime) { return $File.LastAccessTime }
    return $File.LastWriteTime
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

# --- 対象ファイルの収集 ----------------------------------------------------
# デスクトップ直下のファイルのみ。フォルダ・スクリプト自身・ログは除外。
$scriptName = Split-Path -Leaf $PSCommandPath

$files = Get-ChildItem -LiteralPath $DesktopPath -File | Where-Object {
    $_.Name -ne $scriptName -and
    $_.Name -notlike '使用状況ログ_*.txt' -and
    $_.Name -notlike '整理ログ_*.txt'
}

if (-not $IncludeShortcuts) {
    $files = $files | Where-Object { $_.Extension.ToLower() -notin @('.lnk', '.url') }
}

if (-not $files -or $files.Count -eq 0) {
    Write-Log "対象のファイルはありませんでした。"
    $Log | Set-Content -LiteralPath $LogPath -Encoding UTF8
    return
}

Write-Log "対象: $($files.Count) 件"

# --- 振り分け実行 ----------------------------------------------------------
$usedCount = 0
$unusedCount = 0
foreach ($file in $files) {
    $lastUsed = Get-LastUsed $file
    if ($lastUsed -lt $Threshold) {
        $category = $UnusedFolderName
        $destDir  = $UnusedRoot
    }
    else {
        $category = $UsedFolderName
        $destDir  = $UsedRoot
    }

    if (-not (Test-Path -LiteralPath $destDir)) {
        if ($PSCmdlet.ShouldProcess($destDir, 'フォルダ作成')) {
            New-Item -ItemType Directory -Path $destDir -Force | Out-Null
        }
    }

    $destPath = Get-UniquePath -Directory $destDir -FileName $file.Name
    $usedStr  = $lastUsed.ToString('yyyy-MM-dd')

    if ($PSCmdlet.ShouldProcess($file.FullName, "移動 → $category")) {
        try {
            Move-Item -LiteralPath $file.FullName -Destination $destPath -ErrorAction Stop
            Write-Log ("[{0}] {1}  (最終使用: {2})" -f $category, $file.Name, $usedStr)
            if ($category -eq $UnusedFolderName) { $unusedCount++ } else { $usedCount++ }
        }
        catch {
            Write-Log ("[失敗] {0} : {1}" -f $file.Name, $_.Exception.Message)
        }
    }
    else {
        Write-Log ("[予定:{0}] {1}  (最終使用: {2})" -f $category, $file.Name, $usedStr)
    }
}

Write-Log ("完了: 使用中 {0} 件 / 未使用 {1} 件" -f $usedCount, $unusedCount)

# ログ保存（プレビュー時も記録を残す）
try {
    $Log | Set-Content -LiteralPath $LogPath -Encoding UTF8
    Write-Host "`nログを保存しました: $LogPath"
}
catch {
    Write-Host "ログの保存に失敗しました: $($_.Exception.Message)"
}
