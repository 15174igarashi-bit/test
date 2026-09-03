<#
    横浜経理部 月次処理ツール 共通部品

    ・設定（フォルダ名・ファイル名・CSV の固定値）
    ・Box 配下のパス解決（表記ゆれ吸収つき）
    ・.xlsm を Excel なしで読むリーダー

    他のスクリプトから . (ドットソース) で読み込んで使う。
#>

Set-StrictMode -Version Latest

# ============================================================
#  設定 ここから
#  フォルダ名・ファイル名・CSV の固定値が変わったら、ここだけ直す。
# ============================================================

$KeiriUserProfile = if ($env:USERPROFILE) { $env:USERPROFILE } elseif ($env:HOME) { $env:HOME } else { $null }

$KeiriConfig = @{

    # Box のルート候補（上から順に探す）
    BoxRootCandidates = @(
        $(if ($KeiriUserProfile) { Join-Path $KeiriUserProfile 'Box' })
        $(if ($KeiriUserProfile) { Join-Path $KeiriUserProfile 'Box Sync' })
        'C:\Box'
    )

    # 処理年月とデータ月の差（か月）。2026年09月処理 → 2026年8月データ なので -1
    DataMonthOffset = -1

    # 既定の対象法人
    Corporation = '樹慶会'

    # ---- ① 入力元：☆入力用輸入商品YYYY年M月各院経費.xlsm ----
    Source = @{
        # Box 直下から「YYYY年MM月処理」の 1 つ上まで
        SegmentsBeforeMonth = @(
            '横浜経理部'
            '内部関係'
            '1_日)■■奉行 流し込み、各種作業□□'
            '★☆せるふばんく各種振込作業【新宿･横浜】、法人間振替請求書'
        )
        # 「YYYY年MM月処理」から目的地まで
        SegmentsAfterMonth = @(
            '請求書'
            '根拠ﾃﾞｰﾀ'
            '医療法人'
            '法人間処理データ（ｺﾋﾟｰ）'
        )
        MonthFolderFormat     = '{0}年{1:00}月処理'
        WorkbookNameFormat    = '☆入力用輸入商品{0}年{1}月各院経費.xlsm'
        WorkbookSearchPattern = '☆入力用輸入商品*各院経費.xls*'

        SheetName    = '新TWE'
        FirstDataRow = 3
        # 読み取る列（シート上の列記号）
        ColumnClinic    = 'A'   # 医院
        ColumnCorp      = 'B'   # 法人        ← ここでフィルタする
        ColumnClinicId  = 'C'   # 医院ID      → CSV C列
        ColumnAccountId = 'D'   # 勘定科目ID  → CSV D列
        ColumnAmount    = 'E'   # 金額（小数あり）
        # F列 小計 = ROUNDDOWN(E,0)。F の計算結果に依存せず E から作り直し、
        # F にキャッシュ値があれば突合して食い違いを警告する。
        ColumnSubtotal  = 'F'

        # 明細シート。新TWE の E 列は「商品ごと」を (医院ID, 勘定科目ID) で
        # 集計した値を貼り付けたもの（数式ではない）。貼り直し漏れを検知する
        # ため、ここから集計し直して突き合わせる。
        DetailSheetName       = '商品ごと'
        DetailFirstDataRow    = 2
        DetailColumnCorp      = 'A'
        DetailColumnClinic    = 'B'
        DetailColumnClinicId  = 'C'
        DetailColumnAccountId = 'D'
        DetailColumnSubtotal  = 'G'
        DetailTotalLabel      = '合計'   # A列がこれの行は総計行なので除外する
    }

    # ---- ② 出力先：輸入振替_<法人>流し込み用CSV.csv ----
    Destination = @{
        # Box 直下から出力フォルダまで（月フォルダなし＝固定パス）
        Segments = @(
            '横浜経理部'
            '内部関係'
            '14_バクラク'
            '請求書発行'
            '法人間'
            '法人間_個別流し込みCSV'
            '輸入振替'
        )
        # 新規作成するときの名前
        FileNameFormat      = '輸入振替_{0}流し込み用CSV.csv'
        # 既存ファイルを探すときの緩いパターン。「流し込み用CSV」「流し込みCSV」
        # のような表記のゆれで空振りしないようにする。
        FileSearchPattern   = '輸入振替*{0}*CSV*'
        BackupFolderName    = '_backup'

        # 勘定奉行の取込仕様にあわせた固定値
        CodePage    = 932                 # Shift_JIS
        NewLine     = "`r`n"
        Header      = 'OBCD001,CSJS005,CSJS200,CSJS201,CSJS202,CSJS206,CSJS213,CSJS300,CSJS301,CSJS302,CSJS306,CSJS313,CSJS100'
        FirstMark   = '*'                 # 先頭データ行の OBCD001。2 行目以降は空
        DateFormat  = 'yyyy/M/d'          # 2026/8/31（ゼロ埋めなし）
        Summary     = "SBC商品 科目振替  " # CSJS100（末尾の半角スペース 2 つも原本どおり）
        CSJS202     = '901'
        CSJS206     = '2'
        CSJS300     = '0'
        CSJS301     = '171'
        CSJS302     = '2'
        CSJS306     = '2'
    }
}

# ============================================================
#  設定 ここまで
# ============================================================


function ConvertTo-LooseKey {
    <#
        フォルダ名／ファイル名を「ゆれを吸収した比較用キー」に変換する。
        NFKC 正規化で 半角ｶﾀｶﾅ→全角カタカナ、全角（）→半角() などを揃え、
        さらに空白を落とす。
          例) 根拠ﾃﾞｰﾀ               → 根拠データ
              法人間処理データ（ｺﾋﾟｰ） → 法人間処理データ(コピー)
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)

    return ($Text.Normalize([Text.NormalizationForm]::FormKC) -replace '[\s　]', '')
}

function Resolve-ChildDirectory {
    <#
        Parent の下から Name のフォルダを探す。
        完全一致 → ゆれ吸収の一致 → ゆれ吸収の先頭一致 の順。
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

    # 記号も落としたうえでの先頭一致（★☆ や ■□ の増減に耐える）。
    # ただし数字を含む名前は使わない。「2026年12月処理」が「2026年10月処理」に
    # 化けるような取り違えを起こすため、そこは見つからないと言い切る方が安全。
    # 候補が 1 つに絞れないときも、勝手に選ばない。
    if ($Name -notmatch '\d') {
        $stripSymbol = { param($s) ((ConvertTo-LooseKey -Text $s) -replace '[^\p{L}\p{N}]', '') }
        $head = & $stripSymbol $Name
        if ($head.Length -ge 4) {
            $prefix = $head.Substring(0, [Math]::Min(6, $head.Length))
            $hits = @($children | Where-Object { (& $stripSymbol $_.Name).StartsWith($prefix) })
            if ($hits.Count -eq 1) { return $hits[0].FullName }
        }
    }

    return $null
}

function Find-BoxRoot {
    foreach ($candidate in $KeiriConfig.BoxRootCandidates) {
        if ($candidate -and (Test-Path -LiteralPath $candidate -PathType Container)) {
            return $candidate
        }
    }

    $list = ($KeiriConfig.BoxRootCandidates | ForEach-Object { "    $_" }) -join [Environment]::NewLine
    throw @"
Box フォルダが見つかりませんでした。探した場所:
$list
  Box Drive が起動しているか確認するか、scripts\Common.ps1 の BoxRootCandidates を直してください。
"@
}

function Resolve-KeiriPath {
    <#
        Box ルートから Segments を 1 段ずつ解決する。

        戻り値:
          Expected  想定していたフルパス
          Resolved  全部たどれたときの実パス（たどれなければ $null）
          Deepest   実在した一番深い階層
          Missing   見つからなかった階層名（全部たどれたら $null）
    #>
    param([Parameter(Mandatory)][string[]]$Segments)

    $boxRoot = Find-BoxRoot
    $current = $boxRoot
    $missing = $null

    foreach ($segment in $Segments) {
        $next = Resolve-ChildDirectory -Parent $current -Name $segment
        if (-not $next) { $missing = $segment; break }
        $current = $next
    }

    $expected = $boxRoot
    foreach ($segment in $Segments) { $expected = Join-Path $expected $segment }

    return [pscustomobject]@{
        Expected = $expected
        Resolved = $(if ($missing) { $null } else { $current })
        Deepest  = $current
        Missing  = $missing
    }
}

function Resolve-MonthArgument {
    <#
        引数を処理年月の DateTime（その月の 1 日）に変換する。
          省略 / 空          … 今月
          -1 / +2            … 今月からの相対
          202610 / 2026-10 /
          2026/10 / 2026年10月 … 絶対
    #>
    param([AllowEmptyString()][AllowNull()][string]$Value)

    $now   = Get-Date
    $today = Get-Date -Year $now.Year -Month $now.Month -Day 1 -Hour 0 -Minute 0 -Second 0

    if ([string]::IsNullOrWhiteSpace($Value)) { return $today }

    $v = $Value.Trim()

    if ($v -match '^[+-]\d+$') { return $today.AddMonths([int]$v) }

    if ($v -match '^(?<y>\d{4})\D?(?<m>\d{1,2})月?$') {
        $y = [int]$Matches['y']
        $m = [int]$Matches['m']
        if ($m -lt 1 -or $m -gt 12) { throw "月の指定が不正です: $Value" }
        return Get-Date -Year $y -Month $m -Day 1 -Hour 0 -Minute 0 -Second 0
    }

    throw @"
年月の指定を解釈できませんでした: $Value
  例) 202610 / 2026-10 / 2026年10月 / -1（先月） / +1（来月）
"@
}

function Get-SourceWorkbookPath {
    <#
        処理年月から ① のブックを探す。
        正式名が無ければ 輸入商品ブックらしきものの最新を拾う（拾ったら警告）。
    #>
    param([Parameter(Mandatory)][datetime]$ProcessMonth)

    $src        = $KeiriConfig.Source
    $dataMonth  = $ProcessMonth.AddMonths($KeiriConfig.DataMonthOffset)
    $monthFolder = $src.MonthFolderFormat -f $ProcessMonth.Year, $ProcessMonth.Month
    $segments   = @($src.SegmentsBeforeMonth) + @($monthFolder) + @($src.SegmentsAfterMonth)

    $path = Resolve-KeiriPath -Segments $segments
    if (-not $path.Resolved) {
        throw @"
① のフォルダにたどり着けませんでした。
  見つからなかった階層 : $($path.Missing)
  実在する一番深い階層 : $($path.Deepest)
$(if ($path.Missing -eq $monthFolder) { "  → $monthFolder がまだ作られていない可能性があります。" })
"@
    }

    $wantedName = $src.WorkbookNameFormat -f $dataMonth.Year, $dataMonth.Month
    $wanted     = Join-Path $path.Resolved $wantedName
    if (Test-Path -LiteralPath $wanted -PathType Leaf) { return $wanted }

    $fallback = @(
        Get-ChildItem -LiteralPath $path.Resolved -File -Filter $src.WorkbookSearchPattern -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -notlike '~$*' } |
            Sort-Object LastWriteTime -Descending
    )
    if ($fallback.Count -eq 0) {
        throw "「$wantedName」が $($path.Resolved) に見つかりませんでした。"
    }

    Write-Warning "「$wantedName」が無かったので、更新日時が最新の『$($fallback[0].Name)』を使います。"
    return $fallback[0].FullName
}

function Import-XlsxSheet {
    <#
        .xlsx / .xlsm を Excel を起動せずに読む。
        ブックは zip なので、共有文字列とシート XML を直接読む。
        数式セルはキャッシュされた計算結果を返す。

        戻り値: 各行の [hashtable]（キー = 列記号 'A','B',… ＋ '_row' に行番号）
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$SheetName
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "ブックが見つかりません: $Path"
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

    # Excel で開いたままでも読めるよう、共有読み取りでコピーしてから開く
    $temp = Join-Path ([IO.Path]::GetTempPath()) ("keiri_" + [Guid]::NewGuid().ToString('N') + [IO.Path]::GetExtension($Path))
    $in   = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    try {
        $out = [IO.File]::Create($temp)
        try { $in.CopyTo($out) } finally { $out.Dispose() }
    }
    finally { $in.Dispose() }

    $zip = $null
    try {
        $zip = [IO.Compression.ZipFile]::OpenRead($temp)

        $readXml = {
            param($entryName)
            $entry = $zip.Entries | Where-Object { $_.FullName -eq $entryName } | Select-Object -First 1
            if (-not $entry) { return $null }
            $reader = New-Object IO.StreamReader($entry.Open(), [Text.Encoding]::UTF8)
            try { [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
        }

        $ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        $rn = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

        # --- シート名 → 実体の XML パス ---
        $workbook = & $readXml 'xl/workbook.xml'
        if (-not $workbook) { throw "xl/workbook.xml が読めません: $Path" }

        $nsm = New-Object Xml.XmlNamespaceManager($workbook.NameTable)
        $nsm.AddNamespace('m', $ns)
        $nsm.AddNamespace('r', $rn)

        $sheetNode = $workbook.SelectSingleNode("//m:sheets/m:sheet[@name='$SheetName']", $nsm)
        if (-not $sheetNode) {
            $names = ($workbook.SelectNodes('//m:sheets/m:sheet', $nsm) | ForEach-Object { $_.name }) -join ' / '
            throw "シート「$SheetName」が見つかりません。ブック内のシート: $names"
        }
        $relId = $sheetNode.GetAttribute('id', $rn)

        $rels = & $readXml 'xl/_rels/workbook.xml.rels'
        $relNode = $rels.SelectNodes('//*') | Where-Object { $_.LocalName -eq 'Relationship' -and $_.Id -eq $relId } | Select-Object -First 1
        if (-not $relNode) { throw "シート「$SheetName」の実体が特定できません。" }

        $target = $relNode.Target -replace '^/xl/', '' -replace '^/', ''
        $sheetPath = if ($target -like 'xl/*') { $target } else { "xl/$target" }

        # --- 共有文字列 ---
        $shared = @()
        $sst = & $readXml 'xl/sharedStrings.xml'
        if ($sst) {
            $nsmS = New-Object Xml.XmlNamespaceManager($sst.NameTable)
            $nsmS.AddNamespace('m', $ns)
            $shared = @(
                foreach ($si in $sst.SelectNodes('//m:sst/m:si', $nsmS)) {
                    # リッチテキストは <r><t> が複数並ぶので連結する
                    (($si.SelectNodes('.//m:t', $nsmS) | ForEach-Object { $_.InnerText }) -join '')
                }
            )
        }

        # --- シート本体 ---
        $sheet = & $readXml $sheetPath
        if (-not $sheet) { throw "シート XML が読めません: $sheetPath" }

        $nsmW = New-Object Xml.XmlNamespaceManager($sheet.NameTable)
        $nsmW.AddNamespace('m', $ns)

        $result = New-Object Collections.Generic.List[hashtable]

        foreach ($row in $sheet.SelectNodes('//m:sheetData/m:row', $nsmW)) {
            $bag = @{ '_row' = [int]$row.r }

            foreach ($cell in $row.SelectNodes('m:c', $nsmW)) {
                if ($cell.r -notmatch '^([A-Z]+)') { continue }
                $col = $Matches[1]

                $type = $cell.GetAttribute('t')
                $value = $null

                switch ($type) {
                    's' {
                        $vNode = $cell.SelectSingleNode('m:v', $nsmW)
                        if ($vNode) {
                            $idx = [int]$vNode.InnerText
                            if ($idx -ge 0 -and $idx -lt $shared.Count) { $value = $shared[$idx] }
                        }
                    }
                    'inlineStr' {
                        $value = (($cell.SelectNodes('.//m:t', $nsmW) | ForEach-Object { $_.InnerText }) -join '')
                    }
                    'b' {
                        $vNode = $cell.SelectSingleNode('m:v', $nsmW)
                        if ($vNode) { $value = ($vNode.InnerText -eq '1') }
                    }
                    'e' {
                        $vNode = $cell.SelectSingleNode('m:v', $nsmW)
                        if ($vNode) { $value = $vNode.InnerText }   # #REF! などのエラー文字列
                    }
                    default {
                        # '' (数値) と 'str' (数式の文字列結果)
                        $vNode = $cell.SelectSingleNode('m:v', $nsmW)
                        if ($vNode) {
                            $raw = $vNode.InnerText
                            if ($type -eq 'str') {
                                $value = $raw
                            }
                            else {
                                $parsed = 0.0
                                if ([double]::TryParse($raw, [Globalization.NumberStyles]::Float,
                                                       [Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) {
                                    $value = $parsed
                                }
                                else { $value = $raw }
                            }
                        }
                    }
                }

                if ($null -ne $value) { $bag[$col] = $value }
            }

            $result.Add($bag) | Out-Null
        }

        return $result
    }
    finally {
        if ($zip) { $zip.Dispose() }
        Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue
    }
}

function Get-Cp932Encoding {
    <#
        Shift_JIS(CP932) のエンコーディングを返す。
        PowerShell 7 (.NET Core) では既定で載っていないので、必要なら登録する。
    #>
    try { return [Text.Encoding]::GetEncoding($KeiriConfig.Destination.CodePage) }
    catch {
        [Text.Encoding]::RegisterProvider([Text.CodePagesEncodingProvider]::Instance)
        return [Text.Encoding]::GetEncoding($KeiriConfig.Destination.CodePage)
    }
}

function Get-MonthEndDate {
    param([Parameter(Mandatory)][datetime]$Month)
    $lastDay = [DateTime]::DaysInMonth($Month.Year, $Month.Month)
    return (Get-Date -Year $Month.Year -Month $Month.Month -Day $lastDay -Hour 0 -Minute 0 -Second 0)
}

function Get-CellValue {
    param(
        [Parameter(Mandatory)][hashtable]$Row,
        [Parameter(Mandatory)][string]$Column
    )
    if ($Row.ContainsKey($Column)) { return $Row[$Column] }
    return $null
}
