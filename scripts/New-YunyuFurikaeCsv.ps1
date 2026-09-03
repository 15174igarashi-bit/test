<#
.SYNOPSIS
    ①「☆入力用輸入商品YYYY年M月各院経費.xlsm」から
    ②「輸入振替_<法人>流し込み用CSV.csv」を作る。

.DESCRIPTION
    手作業でやっていた
      1. ① の B 列を法人（既定 樹慶会）で絞る
      2. C列→C列 / D列→D列 / F列→G列およびL列 へ貼り付け
      3. 日付は処理日の前月末日
    をそのまま自動化する。

    ・Excel は起動しない（ブックを zip として直接読む）。①を開いたままでも動く。
    ・上書き前に ② を _backup フォルダへ退避する。
    ・書き込む前に必ず検算結果を表示し、確認を求める（-Force で省略）。

    奉行への流し込みは、できあがった ② を人が取り込む（ここは自動化しない）。

.PARAMETER Month
    処理年月。省略時は今日の年月。202610 / 2026-10 / 2026年10月 / -1 / +1

.PARAMETER Corporation
    対象法人。省略時は設定の既定値（樹慶会）。

.PARAMETER SourceFile
    ① を直接指定する（動作確認用）。省略時は Box から自動で探す。

.PARAMETER DestinationFile
    ② を直接指定する（動作確認用）。省略時は Box から自動で探す。

.PARAMETER DryRun
    ② を書き換えず、検算結果と生成内容だけ見せる。

.PARAMETER Force
    確認プロンプトを出さずに上書きする。

.PARAMETER SkipDetailCheck
    明細シート「商品ごと」からの再集計による照合を省く（数秒速くなる）。

.EXAMPLE
    .\New-YunyuFurikaeCsv.ps1 -DryRun
    今月分を試算して、書き込まずに結果だけ見る

.EXAMPLE
    .\New-YunyuFurikaeCsv.ps1
    今月分を作成し、確認のうえ ② を上書きする
#>
[CmdletBinding()]
param(
    [string]$Month,
    [string]$Corporation,
    [string]$SourceFile,
    [string]$DestinationFile,
    [switch]$DryRun,
    [switch]$Force,
    [switch]$SkipDetailCheck
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'Common.ps1')

$src  = $KeiriConfig.Source
$dest = $KeiriConfig.Destination
if ([string]::IsNullOrWhiteSpace($Corporation)) { $Corporation = $KeiriConfig.Corporation }


function Write-Heading {
    param([string]$Text)
    Write-Host ''
    Write-Host "── $Text " -ForegroundColor Cyan -NoNewline
    Write-Host ('─' * [Math]::Max(0, 56 - $Text.Length)) -ForegroundColor DarkCyan
}

function Read-ExistingCsv {
    <#
        既存の ② を読んで (医院ID, 科目) → 金額 の一覧にする。
        壊れていても落とさず、読めた分だけ返す。
    #>
    param([Parameter(Mandatory)][string]$Path)

    $result = [pscustomobject]@{ Rows = @(); Dates = @(); Raw = $null }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $result }

    $bytes = [IO.File]::ReadAllBytes($Path)
    $text  = (Get-Cp932Encoding).GetString($bytes)
    $result.Raw = $text

    $rows  = New-Object Collections.Generic.List[object]
    $dates = New-Object Collections.Generic.List[string]

    foreach ($line in ($text -split "`r?`n")) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $f = $line -split ','
        if ($f.Count -lt 12) { continue }
        if ($f[0] -eq 'OBCD001') { continue }          # ヘッダー

        $clinicId = 0; $accountId = 0; $amount = [long]0
        if (-not [int]::TryParse($f[2], [ref]$clinicId))  { continue }
        if (-not [int]::TryParse($f[3], [ref]$accountId)) { continue }
        [void][long]::TryParse($f[6], [ref]$amount)

        $rows.Add([pscustomobject]@{ ClinicId = $clinicId; AccountId = $accountId; Amount = $amount })
        $dates.Add($f[1])
    }

    $result.Rows  = $rows.ToArray()
    $result.Dates = ($dates | Select-Object -Unique)
    return $result
}


# ============================================================
#  1. 年月と計上日を決める
# ============================================================

try { $processMonth = Resolve-MonthArgument -Value $Month }
catch {
    Write-Host ''
    Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ''
    exit 1
}

$dataMonth   = $processMonth.AddMonths($KeiriConfig.DataMonthOffset)
$postingDate = Get-MonthEndDate -Month $dataMonth      # 処理日の前月末日
$dateText    = $postingDate.ToString($dest.DateFormat, [Globalization.CultureInfo]::InvariantCulture)

Write-Heading '対象'
Write-Host ("  処理年月 : {0}年{1:00}月処理" -f $processMonth.Year, $processMonth.Month)
Write-Host ("  データ月 : {0}年{1}月" -f $dataMonth.Year, $dataMonth.Month)
Write-Host ("  計上日   : {0}  （処理日の前月末日）" -f $dateText)
Write-Host ("  対象法人 : {0}" -f $Corporation)


# ============================================================
#  2. ① を読む
# ============================================================

if ([string]::IsNullOrWhiteSpace($SourceFile)) {
    $SourceFile = Get-SourceWorkbookPath -ProcessMonth $processMonth
}
if (-not (Test-Path -LiteralPath $SourceFile -PathType Leaf)) {
    throw "① が見つかりません: $SourceFile"
}

Write-Heading '① 入力元'
Write-Host "  $SourceFile"

$sheetRows = Import-XlsxSheet -Path $SourceFile -SheetName $src.SheetName

$records  = New-Object Collections.Generic.List[object]
$mismatch = New-Object Collections.Generic.List[object]

foreach ($row in $sheetRows) {
    if ($row['_row'] -lt $src.FirstDataRow) { continue }

    $corp = Get-CellValue -Row $row -Column $src.ColumnCorp
    if ($null -eq $corp -or "$corp" -ne $Corporation) { continue }

    $clinicId  = Get-CellValue -Row $row -Column $src.ColumnClinicId
    $accountId = Get-CellValue -Row $row -Column $src.ColumnAccountId
    $amountRaw = Get-CellValue -Row $row -Column $src.ColumnAmount
    if ($null -eq $clinicId -or $null -eq $accountId -or $null -eq $amountRaw) { continue }

    # F列 小計 = ROUNDDOWN(E,0) を E から作り直す（0 方向への切り捨て）
    $subtotal = [long][Math]::Truncate([double]$amountRaw)

    # ブックに残っている F の計算結果と食い違わないか確認する
    $cachedF = Get-CellValue -Row $row -Column $src.ColumnSubtotal
    if ($null -ne $cachedF -and $cachedF -is [double] -and [long]$cachedF -ne $subtotal) {
        $mismatch.Add([pscustomobject]@{
            Row = $row['_row']; ClinicId = [int]$clinicId; AccountId = [int]$accountId
            Calculated = $subtotal; Cached = [long]$cachedF
        })
    }

    $records.Add([pscustomobject]@{
        Row       = $row['_row']
        Clinic    = "$(Get-CellValue -Row $row -Column $src.ColumnClinic)"
        ClinicId  = [int]$clinicId
        AccountId = [int]$accountId
        Amount    = $subtotal
        AmountRaw = [double]$amountRaw
    })
}

if ($records.Count -eq 0) {
    throw @"
「$Corporation」の行が ① に 1 件もありませんでした。
  シート「$($src.SheetName)」の $($src.ColumnCorp) 列に「$Corporation」があるか確認してください。
"@
}


# ============================================================
#  3. 検算
# ============================================================

Write-Heading '検算'

$byAccount = $records | Group-Object AccountId | Sort-Object Name
$total     = ($records | Measure-Object -Property Amount -Sum).Sum

Write-Host ("  行数 : {0} 行 / 医院 {1} 件" -f $records.Count, (($records.ClinicId | Select-Object -Unique).Count))
Write-Host ''
Write-Host ("    {0,-8} {1,5} {2,18}" -f '勘定科目', '件数', '金額')
foreach ($g in $byAccount) {
    $sum = ($g.Group | Measure-Object -Property Amount -Sum).Sum
    Write-Host ("    {0,-8} {1,5} {2,18:N0}" -f $g.Name, $g.Count, $sum)
}
Write-Host ("    {0,-8} {1,5} {2,18:N0}" -f '合計', $records.Count, $total)

# ブック上の突合表（I:N 列の法人別合計）と照合する
$checkRows = @(
    $sheetRows | Where-Object {
        $v = Get-CellValue -Row $_ -Column 'I'
        $null -ne $v -and "$v" -eq $Corporation
    }
)
if ($checkRows.Count -gt 0) {
    Write-Host ''
    Write-Host '    ブックの突合表との照合（差は切り捨ての端数。件数[円]未満なら正常）'
    $ng = $false
    foreach ($cr in $checkRows) {
        $acct = Get-CellValue -Row $cr -Column 'J'
        $book = Get-CellValue -Row $cr -Column 'K'
        if ($null -eq $acct -or $null -eq $book) { continue }

        $mine  = ($records | Where-Object { $_.AccountId -eq [int]$acct } | Measure-Object -Property Amount -Sum).Sum
        $count = ($records | Where-Object { $_.AccountId -eq [int]$acct }).Count
        $diff  = [double]$book - [double]$mine
        $ok    = ($diff -ge 0) -and ($diff -lt [Math]::Max(1, $count))

        if (-not $ok) { $ng = $true }
        $mark = if ($ok) { 'OK' } else { '要確認' }
        Write-Host ("      科目 {0}  ブック {1,16:N2}  今回 {2,16:N0}  差 {3,8:N2}  {4}" -f `
            $acct, [double]$book, $mine, $diff, $mark) -ForegroundColor $(if ($ok) { 'Gray' } else { 'Yellow' })
    }
    if ($ng) {
        Write-Warning 'ブックの合計と食い違う科目があります。① 側の集計を確認してください。'
    }
}
else {
    Write-Warning "ブック内に「$Corporation」の突合表（I列）が見つからず、合計照合はできませんでした。"
}

# 金額が 0 以下の行
$zeroRows = @($records | Where-Object { $_.Amount -le 0 })
if ($zeroRows.Count -gt 0) {
    Write-Host ''
    Write-Warning "金額が 0 以下の行が $($zeroRows.Count) 件あります。"
    $zeroRows | ForEach-Object {
        Write-Host ("      行{0} 医院{1}({2}) 科目{3} → {4}" -f $_.Row, $_.ClinicId, $_.Clinic, $_.AccountId, $_.Amount)
    }
}

# 明細シート「商品ごと」から集計し直して、新TWE の貼り直し漏れを検知する
if (-not $SkipDetailCheck) {
    Write-Host ''
    Write-Host "    明細シート「$($src.DetailSheetName)」との照合" -NoNewline
    try {
        $detailRows = Import-XlsxSheet -Path $SourceFile -SheetName $src.DetailSheetName
        Write-Host ("  （{0} 行を再集計）" -f $detailRows.Count) -ForegroundColor DarkGray

        $detailSum = @{}
        foreach ($row in $detailRows) {
            if ($row['_row'] -lt $src.DetailFirstDataRow) { continue }

            $corp = Get-CellValue -Row $row -Column $src.DetailColumnCorp
            if ($null -eq $corp -or "$corp" -ne $Corporation) { continue }   # 総計行もここで落ちる

            $cid  = Get-CellValue -Row $row -Column $src.DetailColumnClinicId
            $acct = Get-CellValue -Row $row -Column $src.DetailColumnAccountId
            $amt  = Get-CellValue -Row $row -Column $src.DetailColumnSubtotal
            if ($null -eq $cid -or $null -eq $acct -or $null -eq $amt) { continue }

            $key = '{0}/{1}' -f [int]$cid, [int]$acct
            if (-not $detailSum.ContainsKey($key)) { $detailSum[$key] = [double]0 }
            $detailSum[$key] += [double]$amt
        }

        $detailNg = New-Object Collections.Generic.List[string]

        foreach ($r in $records) {
            $key = '{0}/{1}' -f $r.ClinicId, $r.AccountId
            if (-not $detailSum.ContainsKey($key)) {
                $detailNg.Add(("医院{0} 科目{1} : 新TWE にあるが明細に無い" -f $r.ClinicId, $r.AccountId))
                continue
            }
            $diff = [Math]::Abs($detailSum[$key] - $r.AmountRaw)
            if ($diff -ge 0.005) {
                $detailNg.Add(("医院{0} 科目{1} : 新TWE {2:N2} / 明細 {3:N2}（差 {4:N2}）" -f `
                    $r.ClinicId, $r.AccountId, $r.AmountRaw, $detailSum[$key], ($r.AmountRaw - $detailSum[$key])))
            }
            $detailSum.Remove($key)
        }

        foreach ($key in $detailSum.Keys) {
            $detailNg.Add(("{0} : 明細にあるが新TWE に無い（{1:N2}）" -f $key, $detailSum[$key]))
        }

        if ($detailNg.Count -eq 0) {
            Write-Host "      全 $($records.Count) 行が明細の再集計と一致  OK" -ForegroundColor Gray
        }
        else {
            Write-Warning @"
新TWE の値が、明細シートを集計し直した結果と一致しません。
  新TWE への貼り直しが済んでいない可能性があります。① を確認してください。
"@
            $detailNg | Select-Object -First 15 | ForEach-Object { Write-Host "      $_" -ForegroundColor Yellow }
            if ($detailNg.Count -gt 15) { Write-Host "      … 他 $($detailNg.Count - 15) 件" -ForegroundColor Yellow }
        }
    }
    catch {
        Write-Host ''
        Write-Warning "明細シートとの照合はできませんでした: $($_.Exception.Message)"
    }
}

# F列のキャッシュ値との食い違い
if ($mismatch.Count -gt 0) {
    Write-Host ''
    Write-Warning "① の F 列（小計）の計算結果と食い違う行が $($mismatch.Count) 件あります。ブックを開いて再計算されていない可能性があります。"
    $mismatch | Select-Object -First 10 | ForEach-Object {
        Write-Host ("      行{0} 医院{1} 科目{2} : 算出 {3} / ブック {4}" -f $_.Row, $_.ClinicId, $_.AccountId, $_.Calculated, $_.Cached)
    }
}


# ============================================================
#  4. CSV を組み立てる
# ============================================================

$lines = New-Object Collections.Generic.List[string]
$lines.Add($dest.Header)

$isFirst = $true
foreach ($r in $records) {
    $mark = if ($isFirst) { $dest.FirstMark } else { '' }
    $isFirst = $false

    $lines.Add(@(
        $mark              # OBCD001
        $dateText          # CSJS005  計上日
        $r.ClinicId        # CSJS200  ← ① C列
        $r.AccountId       # CSJS201  ← ① D列
        $dest.CSJS202
        $dest.CSJS206
        $r.Amount          # CSJS213  ← ① F列
        $dest.CSJS300
        $dest.CSJS301
        $dest.CSJS302
        $dest.CSJS306
        $r.Amount          # CSJS313  ← ① F列
        $dest.Summary      # CSJS100
    ) -join ',')
}

$csvText  = ($lines -join $dest.NewLine) + $dest.NewLine
$csvBytes = (Get-Cp932Encoding).GetBytes($csvText)


# ============================================================
#  5. ② を特定して、前回との差を見る
# ============================================================

Write-Heading '② 出力先'

if ([string]::IsNullOrWhiteSpace($DestinationFile)) {
    $destDir = Resolve-KeiriPath -Segments $dest.Segments
    if (-not $destDir.Resolved) {
        throw @"
② のフォルダにたどり着けませんでした。
  見つからなかった階層 : $($destDir.Missing)
  実在する一番深い階層 : $($destDir.Deepest)
"@
    }

    $wantedName = $dest.FileNameFormat -f $Corporation
    $DestinationFile = Join-Path $destDir.Resolved $wantedName

    if (-not (Test-Path -LiteralPath $DestinationFile -PathType Leaf)) {
        # 「流し込み用CSV」「流し込みCSV」のような表記ゆれを拾う
        $candidates = @(
            Get-ChildItem -LiteralPath $destDir.Resolved -File -Filter ($dest.FileSearchPattern -f $Corporation) -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -notlike '~$*' -and $_.Extension -eq '.csv' } |
                Sort-Object LastWriteTime -Descending
        )

        if ($candidates.Count -eq 1) {
            $DestinationFile = $candidates[0].FullName
            Write-Warning "「$wantedName」は無く、『$($candidates[0].Name)』が見つかったので、こちらに書き込みます。"
        }
        elseif ($candidates.Count -gt 1) {
            # 上書き先を勝手に選ばない。人に決めてもらう。
            Write-Host ''
            Write-Host "  上書き先の候補が $($candidates.Count) 件あり、どれか判断できませんでした。" -ForegroundColor Red
            Write-Host "  フォルダ: $($destDir.Resolved)" -ForegroundColor DarkGray
            Write-Host ''
            foreach ($c in $candidates) {
                Write-Host ("    {0}   （更新 {1:yyyy/MM/dd HH:mm}）" -f $c.Name, $c.LastWriteTime)
            }
            Write-Host ''
            Write-Host '  どれに書き込むかを指定して実行してください:' -ForegroundColor Cyan
            Write-Host ("    powershell -File scripts\New-YunyuFurikaeCsv.ps1 -DestinationFile ""{0}""" -f $candidates[0].FullName) -ForegroundColor DarkGray
            Write-Host ''
            Write-Host '  （毎月これで迷うようなら、使わないファイルの名前を変えるか、' -ForegroundColor DarkGray
            Write-Host '    scripts\Common.ps1 の FileSearchPattern を絞ってください）' -ForegroundColor DarkGray
            Write-Host ''
            exit 1
        }
        else {
            Write-Warning "「$wantedName」が無いので、新しく作成します。"
        }
    }
}

Write-Host "  $DestinationFile"

$previous = Read-ExistingCsv -Path $DestinationFile
if ($previous.Rows.Count -gt 0) {
    $prevTotal   = ($previous.Rows | Measure-Object -Property Amount -Sum).Sum
    $prevClinics = @($previous.Rows.ClinicId | Select-Object -Unique)
    $nowClinics  = @($records.ClinicId | Select-Object -Unique)

    Write-Host ''
    Write-Host ("  前回 : {0,4} 行  合計 {1,16:N0}" -f $previous.Rows.Count, $prevTotal)
    Write-Host ("  今回 : {0,4} 行  合計 {1,16:N0}   （増減 {2:N0}）" -f $records.Count, $total, ($total - $prevTotal))

    $added   = @($nowClinics  | Where-Object { $_ -notin $prevClinics })
    $removed = @($prevClinics | Where-Object { $_ -notin $nowClinics })
    if ($added.Count -gt 0) {
        Write-Host ("  今回から増えた医院ID : {0}" -f ($added -join ', ')) -ForegroundColor Yellow
    }
    if ($removed.Count -gt 0) {
        Write-Host ("  今回で無くなった医院ID : {0}" -f ($removed -join ', ')) -ForegroundColor Yellow
    }

    if ($previous.Dates -contains $dateText) {
        Write-Warning "前回のファイルにも計上日 $dateText の行があります。同じ月を二重に作っていないか確認してください。"
    }
}
else {
    Write-Host '  （前回のファイルは無い、または読めませんでした）'
}


# ============================================================
#  6. 書き込む
# ============================================================

Write-Heading '生成内容（先頭 3 行）'
$lines | Select-Object -First 4 | ForEach-Object { Write-Host "  $_" -ForegroundColor DarkGray }
Write-Host ("  … 全 {0} 行（ヘッダー 1 + データ {1}）" -f $lines.Count, $records.Count) -ForegroundColor DarkGray

if ($DryRun) {
    $preview = Join-Path ([IO.Path]::GetTempPath()) ("{0}_{1:yyyyMMdd}_preview.csv" -f $Corporation, $postingDate)
    [IO.File]::WriteAllBytes($preview, $csvBytes)
    Write-Heading '試算のみ（-DryRun）'
    Write-Host '  ② は書き換えていません。生成結果は下記に置きました。'
    Write-Host "  $preview"
    Write-Host ''
    exit 0
}

if (-not $Force) {
    Write-Host ''
    $answer = Read-Host "  ② を上書きします。よろしいですか？ (Y/N)"
    if ($answer -notmatch '^[YyＹｙ]') {
        Write-Host '  中止しました。② は書き換えていません。' -ForegroundColor Yellow
        Write-Host ''
        exit 1
    }
}

# 上書き前に退避（Box のバージョン履歴とは別に手元にも残す）
if (Test-Path -LiteralPath $DestinationFile -PathType Leaf) {
    $backupDir = Join-Path (Split-Path -Parent $DestinationFile) $dest.BackupFolderName
    if (-not (Test-Path -LiteralPath $backupDir -PathType Container)) {
        New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
    }
    $backupName = '{0}_{1:yyyyMMdd_HHmmss}{2}' -f `
        [IO.Path]::GetFileNameWithoutExtension($DestinationFile), (Get-Date), [IO.Path]::GetExtension($DestinationFile)
    $backupPath = Join-Path $backupDir $backupName
    Copy-Item -LiteralPath $DestinationFile -Destination $backupPath -Force
    Write-Host ''
    Write-Host "  退避しました: $backupPath" -ForegroundColor DarkGray
}

try {
    [IO.File]::WriteAllBytes($DestinationFile, $csvBytes)
}
catch [IO.IOException] {
    throw @"
② に書き込めませんでした。Excel などで開いたままになっていませんか。
  $DestinationFile
  ($($_.Exception.Message))
"@
}

# 書けたものを読み直して一致を確認する
$written = [IO.File]::ReadAllBytes($DestinationFile)
if ($written.Length -ne $csvBytes.Length) {
    throw "書き込み後の照合に失敗しました（サイズ不一致: $($written.Length) / $($csvBytes.Length)）。"
}

Write-Heading '完了'
Write-Host ("  {0} 行を書き込みました（Shift_JIS / CRLF）。" -f $records.Count) -ForegroundColor Green
Write-Host "  $DestinationFile"
Write-Host ''
Write-Host '  このあと、このファイルを勘定奉行へ流し込んでください。' -ForegroundColor Cyan
Write-Host ''
exit 0
