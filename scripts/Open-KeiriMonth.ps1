<#
.SYNOPSIS
    横浜経理部「法人間処理データ（ｺﾋﾟｰ）」フォルダ／入力用ブックをワンクリックで開く。

.DESCRIPTION
    Box の深い階層を毎月手で潜らなくて済むように、処理年月から
    フォルダパスを組み立てて Explorer / Excel で開きます。

    処理年月（例: 2026年09月処理）に対して、データ月はその 1 か月前
    （例: ☆入力用輸入商品2026年8月各院経費.xlsm）として扱います。

.PARAMETER Month
    処理年月の指定。省略すると今日の年月。
      202610 / 2026-10 / 2026/10 / 2026年10月 … 絶対指定
      -1 / +2                                  … 今月からの相対指定

.PARAMETER OpenWorkbook
    フォルダではなく、入力用ブック本体を開く。

.PARAMETER ShowPath
    開かずにパスだけ表示する（確認用）。

.EXAMPLE
    .\Open-KeiriMonth.ps1
    今月（＝2026年09月処理）のフォルダを開く

.EXAMPLE
    .\Open-KeiriMonth.ps1 -OpenWorkbook
    ☆入力用輸入商品2026年8月各院経費.xlsm を開く

.EXAMPLE
    .\Open-KeiriMonth.ps1 -Month -1
    先月（＝2026年08月処理）のフォルダを開く
#>
[CmdletBinding()]
param(
    [string]$Month,
    [switch]$OpenWorkbook,
    [switch]$ShowPath
)

$ErrorActionPreference = 'Stop'

# ============================================================
#  設定 ここから（フォルダ名が変わったらここだけ直す）
# ============================================================

# Box のルート。通常は %USERPROFILE%\Box
$BoxRootCandidates = @(
    (Join-Path $env:USERPROFILE 'Box'),
    (Join-Path $env:USERPROFILE 'Box Sync'),
    'C:\Box'
)

# Box ルート直下から「YYYY年MM月処理」フォルダの 1 つ上まで
$SegmentsBeforeMonth = @(
    '横浜経理部',
    '内部関係',
    '1_日)■■奉行 流し込み、各種作業□□',
    '★☆せるふばんく各種振込作業【新宿･横浜】、法人間振替請求書'
)

# 「YYYY年MM月処理」フォルダから目的地まで
$SegmentsAfterMonth = @(
    '請求書',
    '根拠ﾃﾞｰﾀ',
    '医療法人',
    '法人間処理データ（ｺﾋﾟｰ）'
)

# 月フォルダ名の形（{0}=年, {1}=月 ゼロ埋め2桁）
$MonthFolderFormat = '{0}年{1:00}月処理'

# 入力用ブック名の形（{0}=データ年, {1}=データ月 ゼロ埋めなし）
$WorkbookNameFormat = '☆入力用輸入商品{0}年{1}月各院経費.xlsm'

# 入力用ブックが見つからないときの検索パターン
$WorkbookSearchPattern = '☆入力用輸入商品*各院経費.xls*'

# 処理年月とデータ月の差（か月）。2026年09月処理 → 2026年8月データ なので -1
$DataMonthOffset = -1

# ============================================================
#  設定 ここまで
# ============================================================


function Resolve-MonthArgument {
    <#
        引数を処理年月の DateTime（各月 1 日）に変換する。
    #>
    param([string]$Value)

    $now   = Get-Date
    $today = Get-Date -Year $now.Year -Month $now.Month -Day 1 -Hour 0 -Minute 0 -Second 0

    if ([string]::IsNullOrWhiteSpace($Value)) { return $today }

    $v = $Value.Trim()

    # 相対指定: -1, +2, 3
    if ($v -match '^[+-]\d+$') {
        return $today.AddMonths([int]$v)
    }

    # 絶対指定: 202610 / 2026-10 / 2026/10 / 2026年10月 / 2026.10
    if ($v -match '^(?<y>\d{4})\D?(?<m>\d{1,2})月?$') {
        $y = [int]$Matches['y']
        $m = [int]$Matches['m']
        if ($m -lt 1 -or $m -gt 12) {
            throw "月の指定が不正です: $Value"
        }
        return Get-Date -Year $y -Month $m -Day 1
    }

    throw @"
年月の指定を解釈できませんでした: $Value
  例) 202610 / 2026-10 / 2026年10月 / -1（先月） / +1（来月）
"@
}

function ConvertTo-LooseKey {
    <#
        フォルダ名／ファイル名を「ゆれを吸収した比較用キー」に変換する。
        NFKC 正規化で 半角ｶﾀｶﾅ→全角カタカナ、全角（）→半角() などを揃え、
        さらに空白を落とす。
          例) 根拠ﾃﾞｰﾀ              → 根拠データ
              法人間処理データ（ｺﾋﾟｰ） → 法人間処理データ(コピー)
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)

    $normalized = $Text.Normalize([Text.NormalizationForm]::FormKC)
    return ($normalized -replace '[\s　]', '')
}

function Resolve-ChildDirectory {
    <#
        親フォルダの下から、名前が Name のフォルダを探す。
        完全一致 → ゆれ吸収の一致 → ゆれ吸収の先頭一致 の順で探す。
        （半角ｶﾀｶﾅ・全角カッコ・スペースのゆれで見つからない事故を防ぐため）
    #>
    param(
        [Parameter(Mandatory)][string]$Parent,
        [Parameter(Mandatory)][string]$Name
    )

    $exact = Join-Path $Parent $Name
    if (Test-Path -LiteralPath $exact -PathType Container) { return $exact }
    if (-not (Test-Path -LiteralPath $Parent -PathType Container)) { return $null }

    $children = @(Get-ChildItem -LiteralPath $Parent -Directory -ErrorAction SilentlyContinue)
    if ($children.Count -eq 0) { return $null }

    $target = ConvertTo-LooseKey -Text $Name

    $hit = $children |
        Where-Object { (ConvertTo-LooseKey -Text $_.Name) -eq $target } |
        Select-Object -First 1
    if ($hit) { return $hit.FullName }

    # \u8a18\u53f7\u3082\u843d\u3068\u3057\u305f\u3046\u3048\u3067\u306e\u5148\u982d\u4e00\u81f4\uff08\u2605\u2606 \u3084 \u25a0\u25a1 \u306e\u5897\u6e1b\u306b\u8010\u3048\u308b\uff09
    $stripSymbol = { param($s) ((ConvertTo-LooseKey -Text $s) -replace '[^\p{L}\p{N}]', '') }
    $head = & $stripSymbol $Name
    if ($head.Length -ge 4) {
        $prefix = $head.Substring(0, [Math]::Min(6, $head.Length))
        $hit = $children |
            Where-Object { (& $stripSymbol $_.Name).StartsWith($prefix) } |
            Select-Object -First 1
        if ($hit) { return $hit.FullName }
    }

    return $null
}

function Find-BoxRoot {
    foreach ($candidate in $BoxRootCandidates) {
        if ($candidate -and (Test-Path -LiteralPath $candidate -PathType Container)) {
            return $candidate
        }
    }
    throw @"
Box フォルダが見つかりませんでした。
探した場所:
$($BoxRootCandidates | ForEach-Object { "  $_" } | Out-String)
Box Drive が起動しているか確認するか、スクリプト冒頭の `$BoxRootCandidates` を直してください。
"@
}


# ---- 年月を決める -------------------------------------------------
try {
    $processMonth = Resolve-MonthArgument -Value $Month
}
catch {
    Write-Host ''
    Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ''
    exit 1
}
$dataMonth = $processMonth.AddMonths($DataMonthOffset)

$monthFolderName = $MonthFolderFormat -f $processMonth.Year, $processMonth.Month
$workbookName    = $WorkbookNameFormat -f $dataMonth.Year, $dataMonth.Month

Write-Host ''
Write-Host ("  処理年月 : {0}" -f $monthFolderName) -ForegroundColor Cyan
Write-Host ("  データ月 : {0}年{1}月" -f $dataMonth.Year, $dataMonth.Month) -ForegroundColor Cyan
Write-Host ''

# ---- パスを 1 段ずつ解決する ---------------------------------------
$segments = @($SegmentsBeforeMonth) + @($monthFolderName) + @($SegmentsAfterMonth)

$boxRoot  = Find-BoxRoot
$current  = $boxRoot
$deepest  = $boxRoot
$missing  = $null

foreach ($segment in $segments) {
    $next = Resolve-ChildDirectory -Parent $current -Name $segment
    if (-not $next) {
        $missing = $segment
        break
    }
    $current = $next
    $deepest = $next
}

$fullPath = $boxRoot
foreach ($segment in $segments) { $fullPath = Join-Path $fullPath $segment }

if ($ShowPath) {
    Write-Host '  想定パス:' -ForegroundColor DarkGray
    Write-Host "    $fullPath"
    Write-Host '  実在する一番深い階層:' -ForegroundColor DarkGray
    Write-Host "    $deepest"
    if ($missing) {
        Write-Host "  見つからなかった階層: $missing" -ForegroundColor Yellow
    }
    Write-Host ''
    return
}

if ($missing) {
    Write-Warning @"
「$missing」が見つかりませんでした。
$(if ($missing -eq $monthFolderName) { "  → $monthFolderName フォルダがまだ作られていない可能性があります。" } else { "  → フォルダ名が変わったかもしれません。スクリプト冒頭の設定を確認してください。" })
  代わりに、実在する一番深い階層を開きます:
    $deepest
"@
    Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$deepest`""
    exit 1
}

# ---- 開く ----------------------------------------------------------
if (-not $OpenWorkbook) {
    Write-Host "  フォルダを開きます:" -ForegroundColor Green
    Write-Host "    $current"
    Write-Host ''
    Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$current`""
    exit 0
}

$workbookPath = Join-Path $current $workbookName

if (-not (Test-Path -LiteralPath $workbookPath -PathType Leaf)) {
    $candidates = @(
        Get-ChildItem -LiteralPath $current -File -Filter $WorkbookSearchPattern -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -notlike '~$*' } |
            Sort-Object LastWriteTime -Descending
    )

    if ($candidates.Count -eq 0) {
        Write-Warning @"
「$workbookName」が見つかりませんでした。
  フォルダだけ開きます:
    $current
"@
        Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$current`""
        exit 1
    }

    $workbookPath = $candidates[0].FullName
    Write-Warning "「$workbookName」が無かったので、最新の『$($candidates[0].Name)』を開きます。"
}

Write-Host "  ブックを開きます:" -ForegroundColor Green
Write-Host "    $workbookPath"
Write-Host ''
Invoke-Item -LiteralPath $workbookPath
exit 0
