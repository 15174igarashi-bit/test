<#
.SYNOPSIS
    Box の深い階層を潜らずに、目的のフォルダ／ブックを開く。

.DESCRIPTION
    開ける先は 3 つ。
      既定             ① 法人間処理データ（ｺﾋﾟｰ）フォルダ
      -OpenWorkbook    ① ☆入力用輸入商品YYYY年M月各院経費.xlsm
      -CsvFolder       ② 輸入振替フォルダ（法人間_個別流し込みCSV\輸入振替）

    処理年月（例 2026年09月処理）に対して、データ月はその 1 か月前
    （例 ☆入力用輸入商品2026年8月各院経費.xlsm）として扱う。

.PARAMETER Month
    処理年月。省略時は今日の年月。202610 / 2026-10 / 2026年10月 / -1 / +1

.PARAMETER OpenWorkbook
    ① のブック本体を開く。

.PARAMETER CsvFolder
    ② の輸入振替フォルダを開く（月に依存しない固定パス）。

.PARAMETER ShowPath
    開かずにパスだけ表示する。

.EXAMPLE
    .\Open-KeiriMonth.ps1
.EXAMPLE
    .\Open-KeiriMonth.ps1 -OpenWorkbook
.EXAMPLE
    .\Open-KeiriMonth.ps1 -Month -1 -ShowPath
#>
[CmdletBinding()]
param(
    [string]$Month,
    [switch]$OpenWorkbook,
    [switch]$CsvFolder,
    [switch]$ShowPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'Common.ps1')

$src = $KeiriConfig.Source

function Open-InExplorer {
    param([Parameter(Mandatory)][string]$Path)
    Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$Path`""
}


# ---- ② の固定フォルダ ---------------------------------------------
if ($CsvFolder) {
    $path = Resolve-KeiriPath -Segments $KeiriConfig.Destination.Segments

    if ($ShowPath) {
        Write-Host ''
        Write-Host '  想定パス:' -ForegroundColor DarkGray
        Write-Host "    $($path.Expected)"
        Write-Host '  実在する一番深い階層:' -ForegroundColor DarkGray
        Write-Host "    $($path.Deepest)"
        if ($path.Missing) { Write-Host "  見つからなかった階層: $($path.Missing)" -ForegroundColor Yellow }
        Write-Host ''
        exit 0
    }

    if (-not $path.Resolved) {
        Write-Warning "「$($path.Missing)」が見つかりませんでした。実在する一番深い階層を開きます:`r`n    $($path.Deepest)"
        Open-InExplorer -Path $path.Deepest
        exit 1
    }

    Write-Host ''
    Write-Host '  フォルダを開きます:' -ForegroundColor Green
    Write-Host "    $($path.Resolved)"
    Write-Host ''
    Open-InExplorer -Path $path.Resolved
    exit 0
}


# ---- ① 月ごとのフォルダ／ブック ------------------------------------
try { $processMonth = Resolve-MonthArgument -Value $Month }
catch {
    Write-Host ''
    Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ''
    exit 1
}

$dataMonth       = $processMonth.AddMonths($KeiriConfig.DataMonthOffset)
$monthFolderName = $src.MonthFolderFormat -f $processMonth.Year, $processMonth.Month
$workbookName    = $src.WorkbookNameFormat -f $dataMonth.Year, $dataMonth.Month
$segments        = @($src.SegmentsBeforeMonth) + @($monthFolderName) + @($src.SegmentsAfterMonth)

Write-Host ''
Write-Host ("  処理年月 : {0}" -f $monthFolderName) -ForegroundColor Cyan
Write-Host ("  データ月 : {0}年{1}月" -f $dataMonth.Year, $dataMonth.Month) -ForegroundColor Cyan
Write-Host ''

$path = Resolve-KeiriPath -Segments $segments

if ($ShowPath) {
    Write-Host '  想定パス:' -ForegroundColor DarkGray
    Write-Host "    $($path.Expected)"
    Write-Host '  実在する一番深い階層:' -ForegroundColor DarkGray
    Write-Host "    $($path.Deepest)"
    if ($path.Missing) { Write-Host "  見つからなかった階層: $($path.Missing)" -ForegroundColor Yellow }
    Write-Host ''
    exit 0
}

if (-not $path.Resolved) {
    $hint = if ($path.Missing -eq $monthFolderName) {
        "  → $monthFolderName フォルダがまだ作られていない可能性があります。"
    } else {
        '  → フォルダ名が変わったかもしれません。scripts\Common.ps1 の設定を確認してください。'
    }
    Write-Warning @"
「$($path.Missing)」が見つかりませんでした。
$hint
  代わりに、実在する一番深い階層を開きます:
    $($path.Deepest)
"@
    Open-InExplorer -Path $path.Deepest
    exit 1
}

if (-not $OpenWorkbook) {
    Write-Host '  フォルダを開きます:' -ForegroundColor Green
    Write-Host "    $($path.Resolved)"
    Write-Host ''
    Open-InExplorer -Path $path.Resolved
    exit 0
}

$workbookPath = Join-Path $path.Resolved $workbookName

if (-not (Test-Path -LiteralPath $workbookPath -PathType Leaf)) {
    $candidates = @(
        Get-ChildItem -LiteralPath $path.Resolved -File -Filter $src.WorkbookSearchPattern -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -notlike '~$*' } |
            Sort-Object LastWriteTime -Descending
    )

    if ($candidates.Count -eq 0) {
        Write-Warning "「$workbookName」が見つかりませんでした。フォルダだけ開きます:`r`n    $($path.Resolved)"
        Open-InExplorer -Path $path.Resolved
        exit 1
    }

    $workbookPath = $candidates[0].FullName
    Write-Warning "「$workbookName」が無かったので、最新の『$($candidates[0].Name)』を開きます。"
}

Write-Host '  ブックを開きます:' -ForegroundColor Green
Write-Host "    $workbookPath"
Write-Host ''
Invoke-Item -LiteralPath $workbookPath
exit 0
