<#
.SYNOPSIS
    Outlook の受信トレイを仕分けし、「返信すべきメール」だけに返信の下書きを
    自動作成します（送信はしません）。

.DESCRIPTION
    Microsoft Graph（委任ログイン）を使って、次の流れで動きます。

      1. 送信済みトレイを読み、あなたが過去にやり取りした相手から
         「既知の取引先リスト（アドレス／ドメイン）」を自動生成します。
         （知らない業者・新規営業は、このリストに載らないため自動的に除外）
      2. 受信トレイの直近メールを取得します。
      3. 次のルールで仕分けし、残ったものだけを下書き対象にします。
           - noreply / メルマガ / 自動通知など「自動送信」を除外
           - 自分が CC のみ（主宛が別の人）のメールを除外
           - 既知の取引先リストに載っていない差出人を除外
           - フリーメール（gmail 等）はアドレス完全一致のときだけ対象
      4. 対象メールごとに、内容の種別に合わせたテンプレート定型文で
         返信の下書きを Outlook 上に作成します（元メールの引用付き）。
      5. 処理済みのメールには分類（カテゴリ）を付け、次回は二重作成しません。

    ★ 送信は絶対に自動化しません。作成されるのは「下書き」だけです。
      あなたが Outlook の［下書き］フォルダで内容を確認し、手直しして
      送信ボタンを押してください。

.PARAMETER LookbackDays
    受信トレイの何日前までを対象にするか（既定: 3）。

.PARAMETER UnreadOnly
    指定すると未読メールだけを対象にします。

.PARAMETER SentLookbackMonths
    既知の取引先リストを作る際、送信済みトレイを何か月分さかのぼるか（既定: 12）。

.PARAMETER RebuildAllowlist
    既知の取引先リストのキャッシュを作り直します（新しい取引先を反映したいとき）。

.PARAMETER IncludeInternal
    指定すると社内（willows-group.com 等）宛のメールも対象に含めます。
    既定では社外の取引先のみを対象にします。

.PARAMETER MaxDrafts
    1回の実行で作成する下書きの上限（既定: 20／安全のためのブレーキ）。

.PARAMETER WhatIf
    実際には下書きを作らず、何が対象になるかだけを表示します（プレビュー）。
    まずはこれを付けて実行し、仕分け結果を確認してください。

.EXAMPLE
    # 事前準備（初回のみ／管理者権限不要）
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser

.EXAMPLE
    # まずプレビュー（下書きは作らず、仕分け結果だけ表示）
    powershell -ExecutionPolicy Bypass -File .\New-OutlookReplyDrafts.ps1 -WhatIf

.EXAMPLE
    # 本番実行（直近3日・社外の既知取引先のみに下書き作成）
    powershell -ExecutionPolicy Bypass -File .\New-OutlookReplyDrafts.ps1

.EXAMPLE
    # 未読のみ・直近7日を対象
    powershell -ExecutionPolicy Bypass -File .\New-OutlookReplyDrafts.ps1 -LookbackDays 7 -UnreadOnly

.NOTES
    必要モジュール : Microsoft.Graph.Authentication
    必要権限(スコープ): Mail.ReadWrite, Mail.Read
    初回実行時にブラウザで会社アカウントにログイン・同意が必要です。
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [int]$LookbackDays = 3,
    [switch]$UnreadOnly,
    [int]$SentLookbackMonths = 12,
    [switch]$RebuildAllowlist,
    [switch]$IncludeInternal,
    [int]$MaxDrafts = 20
)

# ===========================================================================
#  設定
# ===========================================================================
$GraphBase        = 'https://graph.microsoft.com/v1.0'
$ProcessedCategory = 'AI返信下書き作成済み'   # 処理済みメールに付ける印
$AllowlistPath    = Join-Path $PSScriptRoot 'known-contacts.json'
$AllowlistMaxAgeDays = 7                       # これより古いキャッシュは作り直す

# あなた自身の署名の差出人名（下書きの結び）
$SenderSignName   = '武田'

# 社内ドメイン（既定では対象外。-IncludeInternal で含める）
$InternalDomains  = @('willows-group.com', 'willows.co.jp')

# フリーメール（ドメイン一致では通さず、アドレス完全一致のときだけ対象にする）
$FreeMailDomains  = @(
    'gmail.com', 'googlemail.com', 'yahoo.co.jp', 'yahoo.com', 'ymail.com',
    'hotmail.com', 'hotmail.co.jp', 'outlook.com', 'outlook.jp', 'live.jp',
    'icloud.com', 'me.com', 'nifty.com', 'so-net.ne.jp', 'ocn.ne.jp',
    'biglobe.ne.jp', 'docomo.ne.jp', 'ezweb.ne.jp', 'au.com', 'softbank.ne.jp'
)

# 「自動送信・一斉配信」とみなす差出人アドレスの部分文字列（→ 除外）
$AutomatedSenderPatterns = @(
    'noreply', 'no-reply', 'no_reply', 'donotreply', 'do-not-reply',
    'notification', 'notifications', 'mailer-daemon', 'postmaster', 'bounce',
    'mailmagazine', 'mailmag', 'newsletter', 'mail-magazine'
)

# 件名・本文プレビューに含まれると「返信不要」に寄せるキーワード（→ 除外）
$BulkKeywords = @(
    '配信停止', 'メルマガ', 'メールマガジン', 'セミナー', 'ウェビナー',
    'アンケート', 'キャンペーン', '登録解除', '購読', '【PR】', '広告'
)

# ===========================================================================
#  ログ
# ===========================================================================
$Stamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
$LogPath = Join-Path $PSScriptRoot ("返信下書きログ_{0}.txt" -f $Stamp)
$Log     = [System.Collections.Generic.List[string]]::new()

function Write-Log {
    param([string]$Message)
    $line = "{0}  {1}" -f (Get-Date -Format 'HH:mm:ss'), $Message
    $Log.Add($line)
    Write-Host $line
}

function Save-Log {
    try { $Log | Set-Content -LiteralPath $LogPath -Encoding UTF8 }
    catch { Write-Host "ログ保存に失敗: $($_.Exception.Message)" }
}

# ===========================================================================
#  補助関数
# ===========================================================================

# メールアドレスからドメインを取り出す（小文字化）
function Get-Domain {
    param([string]$Address)
    if ([string]::IsNullOrWhiteSpace($Address) -or ($Address -notmatch '@')) { return '' }
    return ($Address.Split('@')[-1]).Trim().ToLower()
}

# プレーンテキストを HTML 用にエスケープし、改行を <br/> に変換
function ConvertTo-HtmlText {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    $t = $Text -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;'
    $t = $t -replace "`r`n", "`n" -replace "`r", "`n"
    return ($t -replace "`n", '<br/>')
}

# Graph 呼び出し（薄いラッパー。失敗時は $null）
function Invoke-Graph {
    param(
        [string]$Method,
        [string]$Uri,
        $Body
    )
    try {
        if ($null -ne $Body) {
            $json = $Body | ConvertTo-Json -Depth 12
            return Invoke-MgGraphRequest -Method $Method -Uri $Uri -Body $json -ContentType 'application/json' -OutputType PSObject
        }
        return Invoke-MgGraphRequest -Method $Method -Uri $Uri -OutputType PSObject
    }
    catch {
        Write-Log ("[Graphエラー] {0} {1} : {2}" -f $Method, $Uri, $_.Exception.Message)
        return $null
    }
}

# ===========================================================================
#  テンプレート定型文
# ===========================================================================
# {NAME} = 差出人名（○○様）／ 本文は種別で選択。
# 「依頼・質問」型には、回答を書き足す箇所を明示するプレースホルダを入れる。

function Get-ReplyTemplate {
    param(
        [string]$SenderName,
        [string]$Subject,
        [string]$Preview
    )

    $name = if ([string]::IsNullOrWhiteSpace($SenderName)) { 'ご担当者' } else { $SenderName }
    $text = "$Subject`n$Preview"

    # 種別判定（上から優先）
    $isEstimate = ($text -match '見積|御見積|お見積|見積書|請求書')
    $isDocument = ($text -match '資料|書類|ご案内|送付|共有|添付|データ')
    $isRequest  = ($text -match 'ください|下さい|いただけ|頂け|頂戴|お願い|ご確認|ご教示|ご回答|でしょうか|ますか|\?|？')

    $head = "$name 様`n`nお世話になっております。$SenderSignName でございます。`n`n"
    $foot = "`n`n引き続きどうぞよろしくお願いいたします。`n`n$SenderSignName"

    if ($isEstimate) {
        $bodyType = '見積受領'
        $core = "お見積りをお送りいただき、ありがとうございます。`n内容を確認のうえ、改めてご連絡いたします。"
    }
    elseif ($isRequest) {
        $bodyType = '依頼・質問への返信'
        $core = "ご連絡いただき、ありがとうございます。`n確認のうえ、追ってご対応・ご返信いたします。`n`n【ご確認ください：必要な回答・資料をこの下書きに追記してください】"
    }
    elseif ($isDocument) {
        $bodyType = '資料受領'
        $core = "資料をお送りいただき、ありがとうございます。`n確認のうえ、改めてご連絡いたします。"
    }
    else {
        $bodyType = '汎用受領'
        $core = "ご連絡いただき、ありがとうございます。`n内容を確認のうえ、改めてご連絡いたします。"
    }

    return [pscustomobject]@{
        Type = $bodyType
        Text = "$head$core$foot"
    }
}

# ===========================================================================
#  0. 前提チェック（Graph モジュール）
# ===========================================================================
if (-not (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue)) {
    Write-Host ""
    Write-Host "Microsoft Graph モジュールが見つかりません。" -ForegroundColor Yellow
    Write-Host "PowerShell で次を1回だけ実行してください（管理者権限は不要）:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser" -ForegroundColor Cyan
    Write-Host ""
    return
}

# ===========================================================================
#  1. ログイン
# ===========================================================================
Write-Log "Microsoft Graph へログインします（初回はブラウザで会社アカウントに同意）..."
try {
    Connect-MgGraph -Scopes 'Mail.ReadWrite', 'Mail.Read' -NoWelcome -ErrorAction Stop
}
catch {
    # 古いモジュールは -NoWelcome 非対応のことがあるため再試行
    try { Connect-MgGraph -Scopes 'Mail.ReadWrite', 'Mail.Read' -ErrorAction Stop }
    catch {
        Write-Log "ログインに失敗しました: $($_.Exception.Message)"
        Save-Log
        return
    }
}

$me = Invoke-Graph -Method GET -Uri "$GraphBase/me?`$select=mail,userPrincipalName,displayName"
if ($null -eq $me) { Write-Log "ユーザー情報の取得に失敗しました。"; Save-Log; return }
$MyAddress = ([string]$me.mail).ToLower()
if ([string]::IsNullOrWhiteSpace($MyAddress)) { $MyAddress = ([string]$me.userPrincipalName).ToLower() }
Write-Log "ログイン: $($me.displayName) <$MyAddress>"

# ===========================================================================
#  2. 既知の取引先リスト（送信済みトレイから自動生成）
# ===========================================================================
$needBuild = $RebuildAllowlist -or (-not (Test-Path -LiteralPath $AllowlistPath))
if (-not $needBuild) {
    $age = (Get-Date) - (Get-Item -LiteralPath $AllowlistPath).LastWriteTime
    if ($age.TotalDays -gt $AllowlistMaxAgeDays) { $needBuild = $true }
}

if ($needBuild) {
    Write-Log ("既知の取引先リストを作成します（送信済み 直近 {0} か月分を集計）..." -f $SentLookbackMonths)
    $addressSet = [System.Collections.Generic.HashSet[string]]::new()
    $domainSet  = [System.Collections.Generic.HashSet[string]]::new()

    $sinceSent = (Get-Date).ToUniversalTime().AddMonths(-$SentLookbackMonths).ToString('yyyy-MM-ddTHH:mm:ssZ')
    $filter = [uri]::EscapeDataString("sentDateTime ge $sinceSent")
    $sel    = [uri]::EscapeDataString('toRecipients,ccRecipients')
    $uri = "$GraphBase/me/mailFolders/SentItems/messages?`$filter=$filter&`$select=$sel&`$top=100"

    $page = 0
    while ($uri -and $page -lt 20) {
        $resp = Invoke-Graph -Method GET -Uri $uri
        if ($null -eq $resp) { break }
        foreach ($msg in $resp.value) {
            $rcpts = @()
            if ($msg.toRecipients) { $rcpts += $msg.toRecipients }
            if ($msg.ccRecipients) { $rcpts += $msg.ccRecipients }
            foreach ($r in $rcpts) {
                $addr = ([string]$r.emailAddress.address).Trim().ToLower()
                if ([string]::IsNullOrWhiteSpace($addr) -or ($addr -notmatch '@')) { continue }
                [void]$addressSet.Add($addr)
                $dom = Get-Domain $addr
                if ($dom -and ($InternalDomains -notcontains $dom) -and ($FreeMailDomains -notcontains $dom)) {
                    [void]$domainSet.Add($dom)
                }
            }
        }
        $uri = $resp.'@odata.nextLink'
        $page++
    }

    $allow = [pscustomobject]@{
        generatedAt = (Get-Date).ToString('s')
        addresses   = @($addressSet)
        domains     = @($domainSet)
    }
    try {
        $allow | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $AllowlistPath -Encoding UTF8
        Write-Log ("取引先リスト作成: アドレス {0} 件 / ドメイン {1} 件" -f $addressSet.Count, $domainSet.Count)
    }
    catch { Write-Log "取引先リストの保存に失敗: $($_.Exception.Message)" }
}
else {
    Write-Log "既存の取引先リストを使用します（作り直すには -RebuildAllowlist）。"
}

$allowData   = Get-Content -LiteralPath $AllowlistPath -Raw | ConvertFrom-Json
$KnownAddrs  = [System.Collections.Generic.HashSet[string]]::new()
$KnownDomains = [System.Collections.Generic.HashSet[string]]::new()
foreach ($a in $allowData.addresses) { [void]$KnownAddrs.Add(([string]$a).ToLower()) }
foreach ($d in $allowData.domains)   { [void]$KnownDomains.Add(([string]$d).ToLower()) }

# ===========================================================================
#  3. 受信トレイの取得
# ===========================================================================
$sinceInbox = (Get-Date).ToUniversalTime().AddDays(-$LookbackDays).ToString('yyyy-MM-ddTHH:mm:ssZ')
$filterExpr = "receivedDateTime ge $sinceInbox"
if ($UnreadOnly) { $filterExpr += ' and isRead eq false' }

$filter = [uri]::EscapeDataString($filterExpr)
$order  = [uri]::EscapeDataString('receivedDateTime desc')
$sel    = [uri]::EscapeDataString('id,subject,from,toRecipients,ccRecipients,receivedDateTime,categories,bodyPreview,conversationId')
$uri = "$GraphBase/me/mailFolders/Inbox/messages?`$filter=$filter&`$orderby=$order&`$select=$sel&`$top=50"

$messages = [System.Collections.Generic.List[object]]::new()
$page = 0
while ($uri -and $page -lt 10) {
    $resp = Invoke-Graph -Method GET -Uri $uri
    if ($null -eq $resp) { break }
    foreach ($m in $resp.value) { $messages.Add($m) }
    $uri = $resp.'@odata.nextLink'
    $page++
}
Write-Log ("受信トレイ取得: {0} 件（直近 {1} 日{2}）" -f $messages.Count, $LookbackDays, ($(if ($UnreadOnly) { '・未読のみ' } else { '' })))

if ($messages.Count -eq 0) { Write-Log "対象メールはありませんでした。"; Save-Log; return }

# ===========================================================================
#  4. 仕分け ＆ 下書き作成
# ===========================================================================
$made = 0; $skipAuto = 0; $skipCc = 0; $skipUnknown = 0; $skipDone = 0; $skipBulk = 0

foreach ($m in $messages) {
    $fromAddr = ([string]$m.from.emailAddress.address).Trim().ToLower()
    $fromName = [string]$m.from.emailAddress.name
    $subject  = [string]$m.subject
    $preview  = [string]$m.bodyPreview
    $dom      = Get-Domain $fromAddr

    # (a) 既に処理済み
    if ($m.categories -and ($m.categories -contains $ProcessedCategory)) { $skipDone++; continue }

    # (b) 自動送信・配信専用の除外
    $isAuto = $false
    foreach ($p in $AutomatedSenderPatterns) { if ($fromAddr -like "*$p*") { $isAuto = $true; break } }
    if ($isAuto) { $skipAuto++; Write-Log ("[除外:自動送信] {0} <{1}>" -f $subject, $fromAddr); continue }

    # (c) メルマガ・広告キーワードの除外
    $isBulk = $false
    foreach ($k in $BulkKeywords) { if (("$subject`n$preview") -match [regex]::Escape($k)) { $isBulk = $true; break } }
    if ($isBulk) { $skipBulk++; Write-Log ("[除外:一斉配信] {0} <{1}>" -f $subject, $fromAddr); continue }

    # (d) 社内除外（既定）
    if ((-not $IncludeInternal) -and ($InternalDomains -contains $dom)) {
        Write-Log ("[除外:社内] {0} <{1}>" -f $subject, $fromAddr); continue
    }

    # (e) 主宛でない（自分は CC のみ）→ 除外
    $toAddrs = @()
    if ($m.toRecipients) { $toAddrs = $m.toRecipients | ForEach-Object { ([string]$_.emailAddress.address).ToLower() } }
    if ($toAddrs -notcontains $MyAddress) { $skipCc++; Write-Log ("[除外:CCのみ] {0} <{1}>" -f $subject, $fromAddr); continue }

    # (f) 既知の取引先か？（フリーメールはアドレス完全一致のみ）
    $known = $false
    if ($KnownAddrs.Contains($fromAddr)) { $known = $true }
    elseif ($dom -and ($FreeMailDomains -notcontains $dom) -and $KnownDomains.Contains($dom)) { $known = $true }
    if (-not $known) { $skipUnknown++; Write-Log ("[除外:未知の相手] {0} <{1}>" -f $subject, $fromAddr); continue }

    # --- ここまで残ったものが「返信すべきメール」---
    $tpl = Get-ReplyTemplate -SenderName $fromName -Subject $subject -Preview $preview

    if ($made -ge $MaxDrafts) {
        Write-Log ("[上限到達] 下書き作成を {0} 件で打ち切り（-MaxDrafts で調整可）。以降はスキップ。" -f $MaxDrafts)
        break
    }

    if ($PSCmdlet.ShouldProcess("$subject <$fromAddr>", "返信下書き作成（$($tpl.Type)）")) {
        # 1) 返信下書きを作成（元メールの引用付きの draft が返る）
        $draft = Invoke-Graph -Method POST -Uri "$GraphBase/me/messages/$($m.id)/createReply"
        if ($null -eq $draft) { Write-Log ("[失敗] 下書き作成不可: {0}" -f $subject); continue }

        # 2) 本文の先頭に定型文を差し込む（引用はそのまま残す）
        $topHtml = "<div style='font-family:游ゴシック,Yu Gothic,Meiryo,sans-serif'>" +
                   (ConvertTo-HtmlText $tpl.Text) + "</div><br/>"
        $existing = [string]$draft.body.content
        $newBody = @{ body = @{ contentType = 'HTML'; content = ($topHtml + $existing) } }
        $upd = Invoke-Graph -Method PATCH -Uri "$GraphBase/me/messages/$($draft.id)" -Body $newBody
        if ($null -eq $upd) { Write-Log ("[警告] 本文差し込み失敗（空の下書きは作成済み）: {0}" -f $subject) }

        # 3) 元メールに「処理済み」カテゴリを付ける（次回の二重作成を防止）
        $cats = @(); if ($m.categories) { $cats = @($m.categories) }
        if ($cats -notcontains $ProcessedCategory) { $cats += $ProcessedCategory }
        [void](Invoke-Graph -Method PATCH -Uri "$GraphBase/me/messages/$($m.id)" -Body @{ categories = $cats })

        $made++
        Write-Log ("[下書き作成:{0}] {1} <{2}>" -f $tpl.Type, $subject, $fromAddr)
    }
    else {
        # プレビュー（-WhatIf）
        Write-Log ("[予定:{0}] {1} <{2}>" -f $tpl.Type, $subject, $fromAddr)
        $made++
    }
}

# ===========================================================================
#  5. まとめ
# ===========================================================================
Write-Log ""
Write-Log "==== 実行結果 ===="
Write-Log ("下書き{0}: {1} 件" -f ($(if ($WhatIfPreference) { '(予定)' } else { '作成' })), $made)
Write-Log ("除外  : 自動送信 {0} / 一斉配信 {1} / CCのみ {2} / 未知の相手 {3} / 処理済み {4}" -f `
    $skipAuto, $skipBulk, $skipCc, $skipUnknown, $skipDone)
if (-not $WhatIfPreference) {
    Write-Log "Outlook の［下書き］フォルダをご確認ください。内容を手直しのうえ、送信は手動でお願いします。"
}

Save-Log
Write-Host "`nログ: $LogPath"
