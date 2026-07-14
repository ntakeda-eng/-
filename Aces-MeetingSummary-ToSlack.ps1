<#
.SYNOPSIS
    ACES Meet の会議データ（書き起こし / AI要約）を API 経由で取得し、要約を Slack に投稿します。

.DESCRIPTION
    ACES Meet API から「会議」を一覧取得し、各会議の詳細（AI要約・書き起こし）を取り出して
    Slack（Incoming Webhook もしくは Bot トークン）へ投稿する再利用スクリプトです。

    ネットワーク制約のあるクラウド実行環境ではなく、ACES / Slack に到達できる
    あなたの手元の Windows マシンやサーバーで実行してください。

    ── 事前準備 ──────────────────────────────────────────────
    1. ACES Meet API キーを取得（管理者 / API 設定画面より）
    2. 環境変数をセット（PowerShell 例）:
         $env:ACES_MEET_API_KEY = "xxxxxxxx"          # ACES の API キー
         $env:ACES_MEET_API_BASE = "https://<APIベースURL>"  # ドキュメント記載のベースURL
         $env:SLACK_WEBHOOK_URL = "https://hooks.slack.com/services/XXX/YYY/ZZZ"
       （Bot トークンを使う場合は SLACK_WEBHOOK_URL の代わりに）
         $env:SLACK_BOT_TOKEN = "xoxb-..."
         # 投稿先チャンネルは -SlackChannel で指定
    3. まず -Discover で実レスポンスの構造を確認し、下の $Config のフィールド名を合わせる

.PARAMETER Discover
    API のリスト/詳細レスポンスの生 JSON を表示して終了します（Slack へは投稿しません）。
    実際のフィールド名を確認し、$Config のマッピングを調整するために使います。

.PARAMETER DryRun
    Slack へ投稿せず、生成されるメッセージをコンソールに出力します。

.PARAMETER Since
    この日時以降に作成/開催された会議のみ対象（例: "2026-07-01" や "2026-07-01T00:00:00"）。

.PARAMETER Until
    この日時以前の会議のみ対象。

.PARAMETER Limit
    処理する会議の最大件数（0 = 無制限）。既定は 0（全件）。

.PARAMETER SlackChannel
    Bot トークン利用時の投稿先チャンネル（例: "#general" または チャンネルID）。

.EXAMPLE
    # まず構造を確認
    .\Aces-MeetingSummary-ToSlack.ps1 -Discover

.EXAMPLE
    # 直近3件を投稿せず内容だけ確認
    .\Aces-MeetingSummary-ToSlack.ps1 -Limit 3 -DryRun

.EXAMPLE
    # 7/1 以降の会議を全件、Slack に投稿
    .\Aces-MeetingSummary-ToSlack.ps1 -Since 2026-07-01
#>

[CmdletBinding()]
param(
    [string] $ApiKey       = $env:ACES_MEET_API_KEY,
    [string] $ApiBaseUrl   = $env:ACES_MEET_API_BASE,
    [string] $SlackWebhook = $env:SLACK_WEBHOOK_URL,
    [string] $SlackBotToken = $env:SLACK_BOT_TOKEN,
    [string] $SlackChannel  = $env:SLACK_CHANNEL,

    [Nullable[datetime]] $Since,
    [Nullable[datetime]] $Until,
    [int] $Limit = 0,

    [switch] $Discover,
    [switch] $DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ════════════════════════════════════════════════════════════════════════
#  設定ブロック  ─ ★ここを ACES Meet API ドキュメントに合わせて調整してください★
#
#  私（作成者）は API ドキュメントに直接アクセスできなかったため、一般的な REST の
#  慣習に基づく既定値を置いています。-Discover で実際のレスポンスを確認し、
#  必要に応じてパス名 / フィールド名を修正してください。
# ════════════════════════════════════════════════════════════════════════
$Config = @{
    # --- 認証（要確認）-------------------------------------------------
    # ヘッダ名と値の形式。ACES が "Authorization: Bearer <key>" 方式なら既定のままでOK。
    # 独自ヘッダ（例: "X-API-Key: <key>"）の場合は AuthHeaderName/AuthValueFormat を変更。
    AuthHeaderName   = 'Authorization'
    AuthValueFormat  = 'Bearer {0}'      # {0} に $ApiKey が入る。例: 'Bearer {0}' / '{0}'

    # --- エンドポイント（要確認）--------------------------------------
    # 会議一覧を取得するパス（ベースURLからの相対）
    ListPath         = '/v1/meetings'
    # 個別会議の詳細を取得するパス。{id} が会議IDに置換される。
    DetailPath       = '/v1/meetings/{id}'

    # --- ページング（要確認）------------------------------------------
    # 一覧レスポンス内で、会議配列が入っているプロパティ名の候補（先に見つかったものを使用）
    ListItemsKeys    = @('data', 'items', 'meetings', 'results', 'records')
    # 次ページ用トークン/URL が入っているプロパティ名の候補
    NextPageKeys     = @('next', 'next_cursor', 'nextPageToken', 'next_page')
    # ページサイズ用クエリパラメータ名（不要なら空文字）
    PageSizeParam    = 'limit'
    PageSizeValue    = 100

    # --- 日付フィルタ用クエリパラメータ（任意・要確認）----------------
    # サーバー側フィルタに対応していれば設定。未対応ならクライアント側で $Since/$Until を適用。
    SinceParam       = ''                # 例: 'created_from'（空なら送らない）
    UntilParam       = ''                # 例: 'created_to'

    # --- フィールドマッピング（要確認）--------------------------------
    # 各値を取り出す際に試すプロパティ名の候補（先に見つかった非空の値を使用）
    Field_Id         = @('id', 'meeting_id', 'uuid', 'meetingId')
    Field_Title      = @('title', 'name', 'subject', 'meeting_title')
    Field_Date       = @('started_at', 'start_time', 'date', 'created_at', 'held_at')
    Field_Summary    = @('summary', 'ai_summary', 'summary_text', 'minutes', 'abstract')
    Field_Transcript = @('transcript', 'transcription', 'text', 'body')
    Field_Url        = @('url', 'app_url', 'permalink', 'link')

    # --- 出力 ---------------------------------------------------------
    TranscriptMaxChars = 1200   # 要約が無い場合に書き起こしから抜粋する最大文字数
}

# ────────────────────────────────────────────────────────────────────────
#  検証
# ────────────────────────────────────────────────────────────────────────
if ([string]::IsNullOrWhiteSpace($ApiKey)) {
    throw "ACES Meet API キーが未設定です。環境変数 ACES_MEET_API_KEY か -ApiKey を指定してください。"
}
if ([string]::IsNullOrWhiteSpace($ApiBaseUrl)) {
    throw "API ベースURL が未設定です。環境変数 ACES_MEET_API_BASE か -ApiBaseUrl を指定してください（ドキュメント記載のベースURL）。"
}
$ApiBaseUrl = $ApiBaseUrl.TrimEnd('/')

if (-not $Discover -and -not $DryRun) {
    if ([string]::IsNullOrWhiteSpace($SlackWebhook) -and [string]::IsNullOrWhiteSpace($SlackBotToken)) {
        throw "Slack の投稿先が未設定です。SLACK_WEBHOOK_URL か SLACK_BOT_TOKEN(+ -SlackChannel) を指定してください。"
    }
    if (-not [string]::IsNullOrWhiteSpace($SlackBotToken) -and [string]::IsNullOrWhiteSpace($SlackChannel)) {
        throw "SLACK_BOT_TOKEN 利用時は -SlackChannel（または SLACK_CHANNEL）で投稿先を指定してください。"
    }
}

# ────────────────────────────────────────────────────────────────────────
#  ヘルパー
# ────────────────────────────────────────────────────────────────────────

function Get-FirstProp {
    # オブジェクトから候補プロパティ名のうち最初に見つかった非空の値を返す
    param([object] $Object, [string[]] $Names)
    if ($null -eq $Object) { return $null }
    foreach ($n in $Names) {
        if ($Object.PSObject.Properties.Name -contains $n) {
            $v = $Object.$n
            if ($null -ne $v -and -not ([string]::IsNullOrWhiteSpace([string]$v))) { return $v }
        }
    }
    return $null
}

function Invoke-AcesApi {
    param(
        [string] $Path,          # 相対パス or 絶対URL（次ページURL用）
        [hashtable] $Query
    )
    if ($Path -match '^https?://') {
        $uri = $Path
    } else {
        $uri = "$ApiBaseUrl$Path"
    }
    if ($Query -and $Query.Count -gt 0) {
        $pairs = foreach ($k in $Query.Keys) {
            if ($null -ne $Query[$k] -and "$($Query[$k])" -ne '') {
                "{0}={1}" -f [uri]::EscapeDataString($k), [uri]::EscapeDataString("$($Query[$k])")
            }
        }
        $qs = ($pairs -join '&')
        if ($qs) { $uri += ($(if ($uri.Contains('?')) {'&'} else {'?'})) + $qs }
    }

    $headers = @{
        ($Config.AuthHeaderName) = ($Config.AuthValueFormat -f $ApiKey)
        'Accept'                 = 'application/json'
    }

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            return Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -TimeoutSec 60
        } catch {
            $status = $null
            try { $status = [int]$_.Exception.Response.StatusCode } catch {}
            # 429 / 5xx は指数バックオフでリトライ（最大4回）
            if (($status -eq 429 -or ($status -ge 500 -and $status -lt 600)) -and $attempt -le 4) {
                $wait = [math]::Pow(2, $attempt)
                Write-Warning "API $status。$wait 秒後にリトライ ($attempt/4): $uri"
                Start-Sleep -Seconds $wait
                continue
            }
            throw "ACES API 呼び出し失敗 ($status): $uri`n$($_.Exception.Message)"
        }
    }
}

function Get-AcesMeetings {
    # 一覧をページングしながら全件取得（$Limit で打ち切り）
    $all = New-Object System.Collections.Generic.List[object]
    $query = @{}
    if ($Config.PageSizeParam) { $query[$Config.PageSizeParam] = $Config.PageSizeValue }
    if ($Config.SinceParam -and $Since) { $query[$Config.SinceParam] = $Since.Value.ToString('s') }
    if ($Config.UntilParam -and $Until) { $query[$Config.UntilParam] = $Until.Value.ToString('s') }

    $path = $Config.ListPath
    do {
        $resp = Invoke-AcesApi -Path $path -Query $query
        $query = @{}   # 2ページ目以降は next の URL/トークンを使うのでクエリはクリア

        # 配列本体を取り出す
        $items = $null
        if ($resp -is [System.Array]) {
            $items = $resp
        } else {
            $items = Get-FirstProp -Object $resp -Names $Config.ListItemsKeys
            if ($null -eq $items) {
                # ラップされていない単一オブジェクト or 想定外構造
                $items = @($resp)
            }
        }
        foreach ($it in $items) {
            $all.Add($it) | Out-Null
            if ($Limit -gt 0 -and $all.Count -ge $Limit) { return $all }
        }

        # 次ページ
        $next = Get-FirstProp -Object $resp -Names $Config.NextPageKeys
        $path = if ($next) { [string]$next } else { $null }
    } while ($path)

    return $all
}

function ConvertTo-DateSafe {
    param($Value)
    if ($null -eq $Value) { return $null }
    try { return [datetime]$Value } catch { return $null }
}

function Format-SlackMessage {
    param([object] $Meeting, [object] $Detail)

    $src   = if ($Detail) { $Detail } else { $Meeting }
    $title = Get-FirstProp -Object $src -Names $Config.Field_Title
    if (-not $title) { $title = '(無題の会議)' }
    $dateRaw = Get-FirstProp -Object $src -Names $Config.Field_Date
    $url     = Get-FirstProp -Object $src -Names $Config.Field_Url
    $summary = Get-FirstProp -Object $src -Names $Config.Field_Summary

    if (-not $summary) {
        # 要約が無ければ書き起こしから抜粋
        $tr = Get-FirstProp -Object $src -Names $Config.Field_Transcript
        if ($tr) {
            $tr = [string]$tr
            if ($tr.Length -gt $Config.TranscriptMaxChars) {
                $summary = $tr.Substring(0, $Config.TranscriptMaxChars) + " …（以下省略）"
            } else {
                $summary = $tr
            }
            $summary = "（AI要約が無いため書き起こしより抜粋）`n" + $summary
        } else {
            $summary = "（要約・書き起こしデータが見つかりませんでした）"
        }
    }

    $dateStr = ''
    $d = ConvertTo-DateSafe $dateRaw
    if ($d) { $dateStr = $d.ToString('yyyy-MM-dd HH:mm') } elseif ($dateRaw) { $dateStr = [string]$dateRaw }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("*📝 $title*")
    if ($dateStr) { [void]$sb.AppendLine("🗓 $dateStr") }
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine([string]$summary)
    if ($url) { [void]$sb.AppendLine(""); [void]$sb.AppendLine("🔗 $url") }
    return $sb.ToString()
}

function Send-Slack {
    param([string] $Text)

    if (-not [string]::IsNullOrWhiteSpace($SlackWebhook)) {
        $body = @{ text = $Text } | ConvertTo-Json -Depth 5 -Compress
        Invoke-RestMethod -Uri $SlackWebhook -Method Post -ContentType 'application/json; charset=utf-8' -Body $body | Out-Null
        return
    }

    # Bot トークン方式: chat.postMessage
    $headers = @{
        Authorization  = "Bearer $SlackBotToken"
        'Content-Type' = 'application/json; charset=utf-8'
    }
    $body = @{ channel = $SlackChannel; text = $Text } | ConvertTo-Json -Depth 5 -Compress
    $resp = Invoke-RestMethod -Uri 'https://slack.com/api/chat.postMessage' -Method Post -Headers $headers -Body $body
    if (-not $resp.ok) { throw "Slack 投稿失敗: $($resp.error)" }
}

# ────────────────────────────────────────────────────────────────────────
#  メイン
# ────────────────────────────────────────────────────────────────────────

Write-Host "ACES Meet → Slack 要約スクリプト" -ForegroundColor Cyan
Write-Host "ベースURL: $ApiBaseUrl"

if ($Discover) {
    Write-Host "`n=== -Discover: 一覧レスポンス生JSON ===" -ForegroundColor Yellow
    $q = @{}
    if ($Config.PageSizeParam) { $q[$Config.PageSizeParam] = 3 }
    $listResp = Invoke-AcesApi -Path $Config.ListPath -Query $q
    $listResp | ConvertTo-Json -Depth 8 | Write-Host

    # 先頭1件の詳細も表示
    $items = if ($listResp -is [System.Array]) { $listResp } else { Get-FirstProp -Object $listResp -Names $Config.ListItemsKeys }
    if ($items -and @($items).Count -gt 0) {
        $firstId = Get-FirstProp -Object (@($items)[0]) -Names $Config.Field_Id
        if ($firstId) {
            Write-Host "`n=== -Discover: 詳細レスポンス生JSON (id=$firstId) ===" -ForegroundColor Yellow
            $detailPath = $Config.DetailPath -replace '\{id\}', [uri]::EscapeDataString([string]$firstId)
            (Invoke-AcesApi -Path $detailPath) | ConvertTo-Json -Depth 8 | Write-Host
        }
    }
    Write-Host "`n※ 上記の実フィールド名に合わせて、スクリプト上部の `$Config を調整してください。" -ForegroundColor Yellow
    return
}

Write-Host "会議一覧を取得中..." -ForegroundColor Cyan
$meetings = Get-AcesMeetings
Write-Host ("取得: {0} 件" -f $meetings.Count)

# クライアント側の日付フィルタ（サーバー側フィルタ未対応の場合の保険）
if ($Since -or $Until) {
    $meetings = $meetings | Where-Object {
        $d = ConvertTo-DateSafe (Get-FirstProp -Object $_ -Names $Config.Field_Date)
        (-not $d) -or ((-not $Since -or $d -ge $Since.Value) -and (-not $Until -or $d -le $Until.Value))
    }
    Write-Host ("日付フィルタ後: {0} 件" -f @($meetings).Count)
}

$posted = 0; $failed = 0
foreach ($m in $meetings) {
    $id = Get-FirstProp -Object $m -Names $Config.Field_Id
    try {
        $detail = $null
        if ($id) {
            $detailPath = $Config.DetailPath -replace '\{id\}', [uri]::EscapeDataString([string]$id)
            try { $detail = Invoke-AcesApi -Path $detailPath } catch { Write-Warning "詳細取得失敗 (id=$id): $($_.Exception.Message)" }
        }
        $msg = Format-SlackMessage -Meeting $m -Detail $detail

        if ($DryRun) {
            Write-Host "`n---------------- (DryRun) ----------------" -ForegroundColor DarkGray
            Write-Host $msg
        } else {
            Send-Slack -Text $msg
            $posted++
            $srcForTitle = if ($detail) { $detail } else { $m }
            Write-Host ("投稿: {0}" -f (Get-FirstProp -Object $srcForTitle -Names $Config.Field_Title)) -ForegroundColor Green
            Start-Sleep -Milliseconds 400   # Slack レート制限に配慮
        }
    } catch {
        $failed++
        Write-Warning "会議処理失敗 (id=$id): $($_.Exception.Message)"
    }
}

if (-not $DryRun) {
    Write-Host ("`n完了: {0} 件投稿 / {1} 件失敗" -f $posted, $failed) -ForegroundColor Cyan
} else {
    Write-Host "`n(DryRun 完了 — Slack へは投稿していません)" -ForegroundColor Cyan
}
