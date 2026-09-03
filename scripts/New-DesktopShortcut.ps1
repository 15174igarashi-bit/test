<#
.SYNOPSIS
    デスクトップにショートカットを作る（初回に 1 回だけ実行すればよい）。

.DESCRIPTION
    「今月の経理フォルダを開く」「今月の入力用ブックを開く」の 2 つを
    デスクトップに置きます。以後はデスクトップのアイコンをダブルクリック
    するだけで、Box の深い階層を潜らずに目的地へ行けます。

    リポジトリごと別の場所へ移動した場合は、もう一度実行してください。
#>
[CmdletBinding()]
param(
    [switch]$Remove
)

$ErrorActionPreference = 'Stop'

$root    = Split-Path -Parent $PSScriptRoot
$desktop = [Environment]::GetFolderPath('Desktop')

$links = @(
    @{ Name = '今月の経理フォルダを開く.lnk';   Target = (Join-Path $root '1_今月のフォルダを開く.bat') }
    @{ Name = '今月の入力用ブックを開く.lnk';   Target = (Join-Path $root '2_今月の入力用ブックを開く.bat') }
)

if ($Remove) {
    foreach ($link in $links) {
        $path = Join-Path $desktop $link.Name
        if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force
            Write-Host "  削除しました: $($link.Name)" -ForegroundColor Yellow
        }
    }
    return
}

$shell = New-Object -ComObject WScript.Shell

foreach ($link in $links) {
    if (-not (Test-Path -LiteralPath $link.Target -PathType Leaf)) {
        Write-Warning "元ファイルが見つかりません: $($link.Target)"
        continue
    }

    $path = Join-Path $desktop $link.Name
    $sc = $shell.CreateShortcut($path)
    $sc.TargetPath       = $link.Target
    $sc.WorkingDirectory = $root
    $sc.WindowStyle      = 7          # 最小化で起動（黒い画面を出さない）
    $sc.IconLocation     = 'shell32.dll,3'
    $sc.Description      = '横浜経理部 法人間処理データ（ｺﾋﾟｰ）へのショートカット'
    $sc.Save()

    Write-Host "  作成しました: $path" -ForegroundColor Green
}

[void][Runtime.InteropServices.Marshal]::ReleaseComObject($shell)

Write-Host ''
Write-Host '  デスクトップのアイコンをダブルクリックすれば、その月のフォルダ／ブックが開きます。' -ForegroundColor Cyan
Write-Host ''
