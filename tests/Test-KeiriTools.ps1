<#
.SYNOPSIS
    モックの Box ツリーを作って、パス解決と CSV 生成を検証する。

.DESCRIPTION
    実際の Box には一切触れない。テンポラリに偽の Box を組み立て、
    USERPROFILE をそこへ向けてスクリプトを動かす。

    設定（フォルダ名・CSV の固定値）を直したあとに走らせて、
    壊していないことを確認するために使う。

.PARAMETER Workbook
    ① にあたる .xlsm。省略時は tests\fixtures 配下を探す。
    実物が無い場合、CSV 生成のテストはスキップする。

.EXAMPLE
    .\tests\Test-KeiriTools.ps1
#>
[CmdletBinding()]
param(
    [string]$Workbook
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot   = Split-Path -Parent $PSScriptRoot
$scriptsDir = Join-Path $repoRoot 'scripts'

$script:Passed = 0
$script:Failed = 0

function Assert-That {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][bool]$Condition,
        [string]$Detail
    )
    if ($Condition) {
        $script:Passed++
        Write-Host "  [PASS] $Name" -ForegroundColor Green
    }
    else {
        $script:Failed++
        Write-Host "  [FAIL] $Name" -ForegroundColor Red
        if ($Detail) { Write-Host "         $Detail" -ForegroundColor DarkRed }
    }
}

# ---- モックの Box を組み立てる --------------------------------------

$sandbox = Join-Path ([IO.Path]::GetTempPath()) ("keiri_test_" + [Guid]::NewGuid().ToString('N'))
$boxRoot = Join-Path $sandbox 'Box'

$baseSegments = @(
    '横浜経理部'
    '内部関係'
    '1_日)■■奉行 流し込み、各種作業□□'
    '★☆せるふばんく各種振込作業【新宿･横浜】、法人間振替請求書'
)

function Join-Segments {
    param([string]$Root, [string[]]$Segments)
    $p = $Root
    foreach ($s in $Segments) { $p = Join-Path $p $s }
    return $p
}

# 2026年09月処理 … 正規の表記
$sourceDir = Join-Segments -Root $boxRoot -Segments ($baseSegments + @(
    '2026年09月処理', '請求書', '根拠ﾃﾞｰﾀ', '医療法人', '法人間処理データ（ｺﾋﾟｰ）'))
New-Item -ItemType Directory -Path $sourceDir -Force | Out-Null

# 2026年10月処理 … 全角「根拠データ」「(ｺﾋﾟｰ)」の表記ゆれ版
$looseDir = Join-Segments -Root $boxRoot -Segments ($baseSegments + @(
    '2026年10月処理', '請求書', '根拠データ', '医療法人', '法人間処理データ (ｺﾋﾟｰ)'))
New-Item -ItemType Directory -Path $looseDir -Force | Out-Null

# ② 出力先
$destDir = Join-Segments -Root $boxRoot -Segments @(
    '横浜経理部', '内部関係', '14_バクラク', '請求書発行', '法人間', '法人間_個別流し込みCSV', '輸入振替')
New-Item -ItemType Directory -Path $destDir -Force | Out-Null

$savedProfile = $env:USERPROFILE
$env:USERPROFILE = $sandbox

try {
    Write-Host ''
    Write-Host "モック Box: $boxRoot" -ForegroundColor DarkGray
    Write-Host ''
    Write-Host 'パス解決' -ForegroundColor Cyan

    . (Join-Path $scriptsDir 'Common.ps1')

    # 正規の表記
    $p = Resolve-KeiriPath -Segments ($baseSegments + @(
        '2026年09月処理', '請求書', '根拠ﾃﾞｰﾀ', '医療法人', '法人間処理データ（ｺﾋﾟｰ）'))
    Assert-That '正規パスを解決できる' ($p.Resolved -eq $sourceDir) "Resolved=$($p.Resolved)"

    # 表記ゆれ（半角ｶﾀｶﾅ／全角括弧）を NFKC で吸収する
    $p = Resolve-KeiriPath -Segments ($baseSegments + @(
        '2026年10月処理', '請求書', '根拠ﾃﾞｰﾀ', '医療法人', '法人間処理データ（ｺﾋﾟｰ）'))
    Assert-That '半角ｶﾀｶﾅ／全角括弧のゆれを吸収する' ($p.Resolved -eq $looseDir) "Resolved=$($p.Resolved)"

    # 存在しない月を、似た月に取り違えない
    $p = Resolve-KeiriPath -Segments ($baseSegments + @('2026年12月処理'))
    Assert-That '存在しない月を似た月に誤解決しない' `
        ($null -eq $p.Resolved -and $p.Missing -eq '2026年12月処理') `
        "Resolved=$($p.Resolved) Missing=$($p.Missing)"

    # ② の固定パス
    $p = Resolve-KeiriPath -Segments $KeiriConfig.Destination.Segments
    Assert-That '② の出力先を解決できる' ($p.Resolved -eq $destDir) "Resolved=$($p.Resolved)"

    Write-Host ''
    Write-Host '年月の解釈' -ForegroundColor Cyan

    $cases = @(
        @{ In = '202610';     Y = 2026; M = 10 }
        @{ In = '2026-10';    Y = 2026; M = 10 }
        @{ In = '2026/10';    Y = 2026; M = 10 }
        @{ In = '2026年10月'; Y = 2026; M = 10 }
    )
    foreach ($c in $cases) {
        $d = Resolve-MonthArgument -Value $c.In
        Assert-That "「$($c.In)」を解釈できる" (($d.Year -eq $c.Y) -and ($d.Month -eq $c.M)) "→ $d"
    }

    $now  = Get-Date
    $prev = Resolve-MonthArgument -Value '-1'
    $expected = (Get-Date -Year $now.Year -Month $now.Month -Day 1).AddMonths(-1)
    Assert-That '「-1」を先月として解釈できる' `
        (($prev.Year -eq $expected.Year) -and ($prev.Month -eq $expected.Month)) "→ $prev"

    $threw = $false
    try { [void](Resolve-MonthArgument -Value 'abc') } catch { $threw = $true }
    Assert-That '不正な年月を弾く' $threw

    Write-Host ''
    Write-Host '月末日' -ForegroundColor Cyan
    $end = Get-MonthEndDate -Month (Get-Date -Year 2026 -Month 2 -Day 1)
    Assert-That '2026年2月の末日は28日' ($end.Day -eq 28) "→ $end"
    $end = Get-MonthEndDate -Month (Get-Date -Year 2026 -Month 8 -Day 1)
    Assert-That '2026年8月の末日は31日' ($end.Day -eq 31) "→ $end"

    # ---- CSV 生成（① の実物がある場合のみ）-------------------------

    if ([string]::IsNullOrWhiteSpace($Workbook)) {
        $fixture = @(
            Get-ChildItem -Path (Join-Path $PSScriptRoot 'fixtures') -Filter '*.xlsm' -ErrorAction SilentlyContinue |
                Select-Object -First 1
        )
        if ($fixture.Count -gt 0) { $Workbook = $fixture[0].FullName }
    }

    Write-Host ''
    Write-Host 'CSV 生成' -ForegroundColor Cyan

    if ([string]::IsNullOrWhiteSpace($Workbook) -or -not (Test-Path -LiteralPath $Workbook -PathType Leaf)) {
        Write-Host '  [SKIP] ① のブックが無いため省略（-Workbook <path> で指定できます）' -ForegroundColor Yellow
    }
    else {
        Copy-Item -LiteralPath $Workbook -Destination (Join-Path $sourceDir '☆入力用輸入商品2026年8月各院経費.xlsm')

        $outFile = Join-Path $destDir '輸入振替_樹慶会流し込み用CSV.csv'
        $log = & (Join-Path $scriptsDir 'New-YunyuFurikaeCsv.ps1') -Month 202609 -Force *>&1 |
               ForEach-Object { "$_" } | Out-String

        Assert-That '② が作られる' (Test-Path -LiteralPath $outFile -PathType Leaf)
        Assert-That '明細シートから再集計して照合している' `
            ($log -match '明細の再集計と一致') "log=$($log.Length) 文字"

        $bytes = [IO.File]::ReadAllBytes($outFile)
        $text  = (Get-Cp932Encoding).GetString($bytes)
        $lines = @($text -split "`r`n" | Where-Object { $_ -ne '' })

        Assert-That 'ヘッダーが仕様どおり' ($lines[0] -eq $KeiriConfig.Destination.Header) "→ $($lines[0])"
        Assert-That 'CRLF で終わる' ($text.EndsWith("`r`n"))
        Assert-That 'データが 23 行（樹慶会）' ($lines.Count - 1 -eq 23) "→ $($lines.Count - 1) 行"

        $data = $lines[1..($lines.Count - 1)]
        Assert-That '先頭行だけ OBCD001 が *' `
            (($data[0].Split(',')[0] -eq '*') -and (@($data[1..($data.Count-1)] | Where-Object { $_.Split(',')[0] -ne '' }).Count -eq 0))
        Assert-That '全行 13 列' (@($data | Where-Object { $_.Split(',').Count -ne 13 }).Count -eq 0)
        Assert-That '計上日が 2026/8/31' (@($data | Where-Object { $_.Split(',')[1] -ne '2026/8/31' }).Count -eq 0)
        Assert-That 'G列(CSJS213) と L列(CSJS313) が一致' `
            (@($data | Where-Object { $_.Split(',')[6] -ne $_.Split(',')[11] }).Count -eq 0)
        Assert-That '摘要が仕様どおり' `
            (@($data | Where-Object { $_.Split(',')[12] -ne $KeiriConfig.Destination.Summary }).Count -eq 0)
        Assert-That '金額が整数（小数点を含まない）' `
            (@($data | Where-Object { $_.Split(',')[6] -match '\.' }).Count -eq 0)

        # ブックの突合表と一致するか
        $rows = Import-XlsxSheet -Path $Workbook -SheetName $KeiriConfig.Source.SheetName
        foreach ($acct in 5501, 5502, 5503) {
            $mine = (@($data | Where-Object { [int]$_.Split(',')[3] -eq $acct }) |
                     ForEach-Object { [long]$_.Split(',')[6] } | Measure-Object -Sum).Sum
            $book = @($rows | Where-Object {
                (Get-CellValue -Row $_ -Column 'I') -eq '樹慶会' -and
                (Get-CellValue -Row $_ -Column 'J') -eq $acct
            }) | Select-Object -First 1
            if ($book) {
                $expectedSum = [double](Get-CellValue -Row $book -Column 'M')   # SUMIFS(F列)
                Assert-That "科目 $acct がブックの SUMIFS と一致" ($mine -eq [long]$expectedSum) `
                    "CSV=$mine ブック=$expectedSum"
            }
        }

        # 二重実行しても結果が変わらない
        $before = [IO.File]::ReadAllBytes($outFile)
        & (Join-Path $scriptsDir 'New-YunyuFurikaeCsv.ps1') -Month 202609 -Force *> $null
        $after = [IO.File]::ReadAllBytes($outFile)
        Assert-That '二重実行しても内容が変わらない' `
            ([Convert]::ToBase64String($before) -eq [Convert]::ToBase64String($after))

        Assert-That '上書き前のファイルが _backup に残る' `
            (@(Get-ChildItem -Path (Join-Path $destDir '_backup') -Filter '*.csv' -ErrorAction SilentlyContinue).Count -ge 1)

        # -DryRun は ② を触らない
        $stamp = (Get-Item -LiteralPath $outFile).LastWriteTimeUtc
        Start-Sleep -Milliseconds 1100
        & (Join-Path $scriptsDir 'New-YunyuFurikaeCsv.ps1') -Month 202609 -DryRun *> $null
        Assert-That '-DryRun は ② を書き換えない' `
            ((Get-Item -LiteralPath $outFile).LastWriteTimeUtc -eq $stamp)
    }
}
finally {
    $env:USERPROFILE = $savedProfile
    Remove-Item -LiteralPath $sandbox -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host ''
Write-Host ("結果: {0} 件成功 / {1} 件失敗" -f $script:Passed, $script:Failed) `
    -ForegroundColor $(if ($script:Failed -eq 0) { 'Green' } else { 'Red' })
Write-Host ''

exit $(if ($script:Failed -eq 0) { 0 } else { 1 })
