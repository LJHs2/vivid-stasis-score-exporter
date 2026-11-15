#requires -Version 5.1
<#
.SYNOPSIS
    vivid/stasis 全成绩导出工具（Windows 10 / PowerShell 5.1）

.DESCRIPTION
    读取 vivid/stasis 的本地存档，导出全部曲目成绩（官方谱面 + 自制谱面 + 段位挑战）
    到 Excel(.xlsx) / CSV / JSON。

    数据来源：
      %LOCALAPPDATA%\VIVIDSTASIS\highscore_table    官方谱面成绩
      %LOCALAPPDATA%\VIVIDSTASIS\custom_highscore   自制谱面成绩
      %LOCALAPPDATA%\VIVIDSTASIS\profile            玩家资料（昵称 / Rating）
      <游戏目录>\song_information.bin               曲目资料（曲名 / 曲师 / 等级 / 谱师）
      <游戏目录>\Charts\...\*.stats                 谱面音符数等
      <游戏目录>\Custom Songs\...\info.json         自制谱面资料

    本工具只读存档，不会修改任何游戏文件。

.PARAMETER SaveDir
    存档目录。默认自动查找 %LOCALAPPDATA%\VIVIDSTASIS

.PARAMETER GameDir
    游戏安装目录。默认通过 Steam 注册表 / 库文件夹自动查找。

.PARAMETER OutDir
    导出目录。默认脚本所在目录下的“导出结果”。

.PARAMETER NoCustomSongs
    不导出自制谱面成绩。

.PARAMETER NoJson
    不生成 JSON 文件。

.PARAMETER Open
    导出完成后打开导出目录。

.EXAMPLE
    .\Export-VividStasisScores.ps1

.EXAMPLE
    .\Export-VividStasisScores.ps1 -OutDir D:\Scores -NoCustomSongs -Open
#>
[CmdletBinding()]
param(
    [string]$SaveDir,
    [string]$GameDir,
    [string]$OutDir,
    [switch]$NoCustomSongs,
    [switch]$NoJson,
    [switch]$Open
)

$ErrorActionPreference = 'Stop'

trap {
    Write-Host ''
    Write-Host ('  [错误] ' + $_.Exception.Message) -ForegroundColor Red
    Write-Host '  请对照 README.md 的“常见问题”检查存档目录 / 游戏目录。' -ForegroundColor DarkGray
    Write-Host ''
    exit 1
}

$script:Inv  = [System.Globalization.CultureInfo]::InvariantCulture
$script:Utf8 = New-Object System.Text.UTF8Encoding($false)

# ---------------------------------------------------------------- 通用工具

function Write-Head([string]$text) {
    Write-Host ''
    Write-Host ("  " + $text) -ForegroundColor Cyan
}

function Write-Info([string]$text) {
    Write-Host ("  " + $text) -ForegroundColor Gray
}

function Write-Ok([string]$text) {
    Write-Host ("  [OK] " + $text) -ForegroundColor Green
}

function Write-Warn2([string]$text) {
    Write-Host ("  [!] " + $text) -ForegroundColor Yellow
}

function Read-Text([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) { return '' }
    $bytes = [System.IO.File]::ReadAllBytes($path)
    $text  = $script:Utf8.GetString($bytes)
    if ($text.Length -gt 0 -and $text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
    return $text
}

function ToNum([object]$value) {
    if ($null -eq $value) { return 0.0 }
    $s = [string]$value
    if ($s.Length -eq 0) { return 0.0 }
    $d = 0.0
    if ([double]::TryParse($s, [System.Globalization.NumberStyles]::Float, $script:Inv, [ref]$d)) { return $d }
    return 0.0
}

function Format-Num([double]$value, [string]$fmt) {
    return $value.ToString($fmt, $script:Inv)
}

# 读取 GameMaker 风格的 INI（key="value"）
function Read-IniData([string]$path) {
    $map = @{}
    $text = Read-Text $path
    if (-not $text) { return $map }
    $cur = $null
    foreach ($line in ($text -split "`r?`n")) {
        $t = $line.Trim()
        if ($t.Length -eq 0) { continue }
        if ($t.StartsWith('[') -and $t.EndsWith(']')) {
            $cur = $t.Substring(1, $t.Length - 2)
            if (-not $map.ContainsKey($cur)) { $map[$cur] = @{} }
            continue
        }
        if ($null -eq $cur) { continue }
        $idx = $t.IndexOf('=')
        if ($idx -lt 0) { continue }
        $key = $t.Substring(0, $idx).Trim()
        $val = $t.Substring($idx + 1).Trim()
        if ($val.Length -ge 2 -and $val.StartsWith('"') -and $val.EndsWith('"')) {
            $val = $val.Substring(1, $val.Length - 2)
        }
        if (-not $map.ContainsKey($cur)) { $map[$cur] = @{} }
        $map[$cur][$key] = $val
    }
    return $map
}

# ---------------------------------------------------------------- 曲目资料

# 解析 song_information.bin（VSD 二进制，token 流）
function Read-SongInfo([string]$path) {
    $songs = New-Object System.Collections.ArrayList
    if (-not (Test-Path -LiteralPath $path)) { return @() }
    $bytes = [System.IO.File]::ReadAllBytes($path)
    $len = $bytes.Length
    # 文件头必须是 "VSD"
    if ($len -lt 8 -or $bytes[0] -ne 0x56 -or $bytes[1] -ne 0x53 -or $bytes[2] -ne 0x44) { return @() }

    $pos = 5
    $cur = $null
    while ($pos -lt $len) {
        $tag = $bytes[$pos]
        if ($tag -eq 0xA0) {                      # 新曲目
            if ($null -ne $cur) { [void]$songs.Add($cur) }
            $cur = @{
                id = 0; title = ''; titleAlt = ''; artist = ''; artistAlt = ''
                chart = ''; bpm = ''; version = ''
                charts = New-Object System.Collections.ArrayList
            }
            $pos += 1
        }
        elseif ($tag -eq 0xA1) {                  # 分隔符
            $pos += 1
        }
        elseif ($tag -eq 0xA2) {                  # 字段
            if ($pos + 3 -ge $len) { break }
            $sub = $bytes[$pos + 1]
            $key = [int]$bytes[$pos + 2]
            if ($sub -eq 0xB2) {
                if ($pos + 7 -gt $len) { break }
                $val = [BitConverter]::ToUInt32($bytes, $pos + 3)
                if ($key -eq 0 -and $null -ne $cur) { $cur.id = [int]$val }
                $pos += 7
            }
            elseif ($sub -eq 0xB7) {
                if ($pos + 4 -gt $len) { break }
                $pos += 4
            }
            elseif ($sub -eq 0xB8) {
                $start = $pos + 3
                $i = $start
                while ($i -lt $len -and $bytes[$i] -ne 0) { $i++ }
                $str = $script:Utf8.GetString($bytes, $start, $i - $start)
                if ($null -ne $cur) {
                    switch ($key) {
                        1 { $cur.title     = $str }
                        2 { $cur.titleAlt  = $str }
                        3 { $cur.artist    = $str }
                        4 { $cur.chart     = $str }
                        5 { $cur.bpm       = $str }
                        6 { $cur.version   = $str }
                        9 { $cur.artistAlt = $str }
                    }
                }
                $pos = $i + 1
            }
            else { break }
        }
        elseif ($tag -eq 0xC0) {                  # 难度：等级显示 / 定数 / 谱师
            $i = $pos + 1
            while ($i -lt $len -and $bytes[$i] -ne 0) { $i++ }
            $level = $script:Utf8.GetString($bytes, $pos + 1, $i - $pos - 1)
            if ($i + 5 -gt $len) { break }
            $const = [double][BitConverter]::ToSingle($bytes, $i + 1)
            $j = $i + 5
            while ($j -lt $len -and $bytes[$j] -ne 0) { $j++ }
            $designer = $script:Utf8.GetString($bytes, $i + 5, $j - $i - 5)
            if ($null -ne $cur) {
                [void]$cur.charts.Add(@{ level = $level; const = $const; designer = $designer })
            }
            $pos = $j + 1
        }
        else { break }                            # 尾部加密数据，忽略
    }
    if ($null -ne $cur) { [void]$songs.Add($cur) }
    return @($songs)
}

# 自制曲目 info.json
function Read-CustomSongInfo([string]$root) {
    $map = @{}
    if (-not (Test-Path -LiteralPath $root)) { return $map }
    $files = @(Get-ChildItem -LiteralPath $root -Recurse -File -Filter 'info.json' -ErrorAction SilentlyContinue)
    foreach ($file in $files) {
        $folder = $file.DirectoryName
        $obj = $null
        try { $obj = ConvertFrom-Json (Read-Text $file.FullName) } catch { continue }
        if ($null -eq $obj) { continue }

        $cid = [string]$obj.chart_id
        if (-not $cid) { $cid = $file.Directory.Name }
        if (-not $cid) { continue }
        if ($map.ContainsKey($cid)) { continue }

        $title = [string]$obj.name
        if (-not $title) { $title = [string]$obj.formatted_name }
        if (-not $title) { $title = $cid }

        $entry = @{
            id      = $cid
            title   = $title
            artist  = [string]$obj.artist
            bpm     = [string]$obj.bpm_display
            folder  = $folder
            diffs   = @{}
            special = @{}
        }

        for ($i = 1; $i -le 4; $i++) {
            $disp  = $obj.("difficulty_display_$i")
            $const = $obj.("difficulty_constant_$i")
            $des   = $obj.("note_designer_$i")
            if ($null -ne $disp -or $null -ne $const) {
                $entry.diffs[$i] = @{
                    level    = [string]$disp
                    const    = $(if ($null -ne $const) { [double]$const } else { $null })
                    designer = [string]$des
                }
            }
        }

        $sp = Join-Path $folder 'shatterinfo.json'
        if (Test-Path -LiteralPath $sp) {
            try {
                $o2 = ConvertFrom-Json (Read-Text $sp)
                $name = [string]$o2.difficulty_name
                if ($name) {
                    $entry.special[$name] = @{
                        level    = [string]$o2.difficulty_number
                        const    = $null
                        designer = [string]$o2.note_designer
                    }
                }
            } catch { }
        }
        $map[$cid] = $entry
    }
    return $map
}

# ---------------------------------------------------------------- 谱面统计

function Add-StatsFromFolder([hashtable]$index, [string]$root) {
    if (-not $root -or -not (Test-Path -LiteralPath $root)) { return }
    $files = @(Get-ChildItem -LiteralPath $root -Recurse -File -Filter '*.stats' -ErrorAction SilentlyContinue)
    foreach ($file in $files) {
        $diff = $file.BaseName
        $cid  = $file.Directory.Name
        if (-not $cid -or -not $diff) { continue }
        $key = $cid + '|' + $diff
        if ($index.ContainsKey($key)) { continue }
        $data = Read-IniData $file.FullName
        if ($data.ContainsKey('stats')) {
            $n = ToNum $data['stats']['noteCount']
            if ($n -gt 0) { $index[$key] = [int]$n }
        }
    }
}

function Add-CustomStatsFromFolder([hashtable]$index, [string]$root) {
    if (-not $root -or -not (Test-Path -LiteralPath $root)) { return }
    $files = @(Get-ChildItem -LiteralPath $root -Recurse -File -Filter '*.stats_custom' -ErrorAction SilentlyContinue)
    foreach ($file in $files) {
        $diff = $file.BaseName
        $cid  = $file.Directory.Name
        if (-not $cid -or -not $diff) { continue }
        $key = $cid + '|' + $diff
        if ($index.ContainsKey($key)) { continue }
        $data = Read-IniData $file.FullName
        if ($data.ContainsKey('stats')) {
            $n = ToNum $data['stats']['noteCount']
            if ($n -gt 0) { $index[$key] = [int]$n }
        }
    }
}

# ---------------------------------------------------------------- 定位目录

function Test-SaveDir([string]$path) {
    if (-not $path) { return $false }
    return (Test-Path -LiteralPath (Join-Path $path 'highscore_table'))
}

function Find-SaveDir {
    $candidates = New-Object System.Collections.ArrayList
    if ($env:LOCALAPPDATA) { [void]$candidates.Add((Join-Path $env:LOCALAPPDATA 'VIVIDSTASIS')) }
    if ($env:USERPROFILE)  { [void]$candidates.Add((Join-Path $env:USERPROFILE 'AppData\Local\VIVIDSTASIS')) }
    foreach ($c in $candidates) { if (Test-SaveDir $c) { return $c } }
    return $null
}

function Find-GameDir {
    $candidates = New-Object System.Collections.ArrayList

    $steamPath = $null
    try {
        $steamPath = (Get-ItemProperty -Path 'HKCU:\Software\Valve\Steam' -ErrorAction Stop).SteamPath
    } catch { }
    if (-not $steamPath) {
        try {
            $steamPath = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Valve\Steam' -ErrorAction Stop).InstallPath
        } catch { }
    }

    if ($steamPath) {
        $steamPath = $steamPath -replace '/', '\'
        [void]$candidates.Add((Join-Path $steamPath 'steamapps\common\vividstasis'))
        $lf = Join-Path $steamPath 'steamapps\libraryfolders.vdf'
        if (Test-Path -LiteralPath $lf) {
            $txt = Read-Text $lf
            foreach ($m in [regex]::Matches($txt, '"path"\s+"([^"]+)"')) {
                $lib = $m.Groups[1].Value -replace '\\\\', '\'
                $manifest = Join-Path $lib 'steamapps\appmanifest_2093940.acf'
                $installdir = 'vividstasis'
                if (Test-Path -LiteralPath $manifest) {
                    $mtxt = Read-Text $manifest
                    $mm = [regex]::Match($mtxt, '"installdir"\s+"([^"]+)"')
                    if ($mm.Success -and $mm.Groups[1].Value) { $installdir = $mm.Groups[1].Value }
                }
                [void]$candidates.Add((Join-Path $lib ('steamapps\common\' + $installdir)))
            }
        }
    }

    foreach ($d in @(
        'C:\Program Files (x86)\Steam\steamapps\common\vividstasis',
        'C:\Program Files\Steam\steamapps\common\vividstasis',
        'D:\SteamLibrary\steamapps\common\vividstasis',
        'E:\SteamLibrary\steamapps\common\vividstasis',
        'F:\SteamLibrary\steamapps\common\vividstasis'
    )) { [void]$candidates.Add($d) }

    try {
        $un = Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Steam App 2093940' -ErrorAction Stop
        if ($un.InstallLocation) { [void]$candidates.Add($un.InstallLocation) }
    } catch { }

    foreach ($c in $candidates) {
        if ($c -and (Test-Path -LiteralPath (Join-Path $c 'song_information.bin'))) { return $c }
    }
    return $null
}

# ---------------------------------------------------------------- 评级

$script:StdDiffOrder = @{ 'OPENING' = 0; 'MIDDLE' = 1; 'FINALE' = 2; 'ENCORE' = 3 }

function Get-DiffOrder([string]$diff) {
    if ($script:StdDiffOrder.ContainsKey($diff)) { return $script:StdDiffOrder[$diff] }
    return 10
}

function Get-ScoreGrade([double]$score) {
    if ($score -ge 1009999.5) { return 'VS' }
    if ($score -ge 1009000)   { return 'V+' }
    if ($score -ge 1008000)   { return 'V' }
    if ($score -ge 1004000)   { return 'SS+' }
    if ($score -ge 1000000)   { return 'SS' }
    if ($score -ge 990000)    { return 'S+' }
    if ($score -ge 980000)    { return 'S' }
    if ($score -ge 950000)    { return 'AA' }
    if ($score -ge 900000)    { return 'A' }
    if ($score -ge 850000)    { return 'B' }
    if ($score -ge 800000)    { return 'C' }
    if ($score -ge 600000)    { return 'D' }
    return 'E'
}

function Get-ExGrade([double]$pct) {
    if ($pct -ge 99.9995) { return 'MAX' }
    if ($pct -ge 99)      { return 'PLATINUM' }
    if ($pct -ge 98)      { return 'GOLD' }
    if ($pct -ge 96)      { return 'SILVER' }
    if ($pct -ge 93)      { return 'BRONZE' }
    if ($pct -ge 90)      { return 'GREAT' }
    if ($pct -ge 80)      { return 'GOOD' }
    return '—'
}

function Get-DiffsFromSection([hashtable]$section) {
    $list = New-Object System.Collections.ArrayList
    if ($null -eq $section) { return @() }
    foreach ($k in $section.Keys) {
        if ($k -match '_(GAME|MAXCOMBO|COMBO|GAMEPERC)$') { continue }
        [void]$list.Add([string]$k)
    }
    return @($list)
}

function New-ScoreRow {
    param(
        [string]$Type, [string]$Title, [string]$Artist, [string]$Diff,
        [string]$Level, $Const, [string]$Designer,
        [double]$Score, [double]$Game, [double]$ExpPct,
        [double]$MaxCombo, [double]$ComboFlag, [int]$Notes,
        [string]$ChartId, $SongId, [string]$Bpm
    )
    $grade = ''
    if ($Score -gt 0) { $grade = Get-ScoreGrade $Score }
    $exGrade = ''
    if ($ExpPct -gt 0) { $exGrade = Get-ExGrade $ExpPct }

    $status = ''
    if ($ComboFlag -ge 2)     { $status = 'AC' }
    elseif ($ComboFlag -ge 1) { $status = 'FC' }
    elseif ($Score -gt 0)     { $status = 'Cleared' }
    elseif ($Game -gt 0)      { $status = 'EX Only' }

    return @{
        type = $Type; title = $Title; artist = $Artist; diff = $Diff
        level = $Level; const = $Const; designer = $Designer
        score = $Score; grade = $grade
        exScore = $Game; exp = $ExpPct; exGrade = $exGrade
        maxCombo = $MaxCombo; notes = $Notes; status = $status
        chartId = $ChartId; songId = $SongId; bpm = $Bpm
    }
}

# ---------------------------------------------------------------- XLSX 输出

function Escape-Xml([string]$s) {
    if ([string]::IsNullOrEmpty($s)) { return '' }
    $s = [regex]::Replace($s, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
    return [System.Security.SecurityElement]::Escape($s)
}

function Get-ColName([int]$n) {
    $name = ''
    while ($n -gt 0) {
        $m = ($n - 1) % 26
        $name = [char](65 + $m) + $name
        $n = [int][math]::Floor(($n - 1) / 26)
    }
    return $name
}

function New-SheetXml {
    param([object[]]$Columns, [object[]]$Rows, [bool]$Selected)

    $Rows = @($Rows)
    $colCount = $Columns.Count
    $rowCount = 1 + $Rows.Count
    $lastCol  = Get-ColName $colCount
    $dimRows  = [Math]::Max(1, $rowCount)
    $sel = ''
    if ($Selected) { $sel = ' tabSelected="1"' }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    [void]$sb.Append('<dimension ref="A1:' + $lastCol + $dimRows + '"/>')
    [void]$sb.Append('<sheetViews><sheetView' + $sel + ' workbookViewId="0">')
    [void]$sb.Append('<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>')
    [void]$sb.Append('<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>')
    [void]$sb.Append('</sheetView></sheetViews>')
    [void]$sb.Append('<sheetFormatPr defaultRowHeight="15"/>')

    [void]$sb.Append('<cols>')
    for ($i = 0; $i -lt $colCount; $i++) {
        $w = [double]$Columns[$i].w
        [void]$sb.Append('<col min="' + ($i + 1) + '" max="' + ($i + 1) + '" width="' + (Format-Num $w '0.##') + '" customWidth="1"/>')
    }
    [void]$sb.Append('</cols>')

    [void]$sb.Append('<sheetData>')

    # 表头
    [void]$sb.Append('<row r="1" ht="24" customHeight="1">')
    for ($i = 0; $i -lt $colCount; $i++) {
        $ref = (Get-ColName ($i + 1)) + '1'
        $h = Escape-Xml ([string]$Columns[$i].h)
        [void]$sb.Append('<c r="' + $ref + '" s="1" t="inlineStr"><is><t>' + $h + '</t></is></c>')
    }
    [void]$sb.Append('</row>')

    $r = 1
    foreach ($row in $Rows) {
        $r++
        if ($null -eq $row) { continue }
        [void]$sb.Append('<row r="' + $r + '">')
        $cells = @($row)
        for ($i = 0; $i -lt $colCount; $i++) {
            $v = $null
            if ($i -lt $cells.Count) { $v = $cells[$i] }
            if ($null -eq $v) { continue }
            $ref = (Get-ColName ($i + 1)) + $r
            $type = [string]$Columns[$i].t

            if ($v -is [string]) {
                if ($v.Length -eq 0) { continue }
                [void]$sb.Append('<c r="' + $ref + '" s="6" t="inlineStr"><is><t xml:space="preserve">' + (Escape-Xml $v) + '</t></is></c>')
            }
            else {
                $d = [double]$v
                $style = 2
                $fmt = '0'
                switch ($type) {
                    'num2' { $style = 3; $fmt = '0.00' }
                    'num1' { $style = 4; $fmt = '0.0' }
                    default { $style = 2; $fmt = '0' }
                }
                $txt = $d.ToString($fmt, $script:Inv)
                [void]$sb.Append('<c r="' + $ref + '" s="' + $style + '"><v>' + $txt + '</v></c>')
            }
        }
        [void]$sb.Append('</row>')
    }

    [void]$sb.Append('</sheetData>')
    if ($Rows.Count -gt 0) {
        [void]$sb.Append('<autoFilter ref="A1:' + $lastCol + $rowCount + '"/>')
    }
    [void]$sb.Append('</worksheet>')
    return $sb.ToString()
}

function Add-ZipText($zip, [string]$name, [string]$text) {
    $entry  = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    try {
        $bytes = $script:Utf8.GetBytes($text)
        $stream.Write($bytes, 0, $bytes.Length)
    } finally {
        $stream.Dispose()
    }
}

function Get-StylesXml {
    return @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="2"><numFmt numFmtId="164" formatCode="0.00"/><numFmt numFmtId="165" formatCode="0.0"/></numFmts>
<fonts count="3">
<font><sz val="10"/><color rgb="FF222222"/><name val="Microsoft YaHei"/><charset val="134"/></font>
<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Microsoft YaHei"/><charset val="134"/></font>
<font><b/><sz val="11"/><color rgb="FF1F3864"/><name val="Microsoft YaHei"/><charset val="134"/></font>
</fonts>
<fills count="4">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFF2F4F8"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="2">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color rgb="FFD6D9E0"/></left><right style="thin"><color rgb="FFD6D9E0"/></right><top style="thin"><color rgb="FFD6D9E0"/></top><bottom style="thin"><color rgb="FFD6D9E0"/></bottom><diagonal/></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="7">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="3" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
'@
}

function New-XlsxFile {
    param([string]$Path, [object[]]$Sheets)

    Add-Type -AssemblyName System.IO.Compression | Out-Null
    Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null

    $count = $Sheets.Count

    $ct = New-Object System.Text.StringBuilder
    [void]$ct.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$ct.Append('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">')
    [void]$ct.Append('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>')
    [void]$ct.Append('<Default Extension="xml" ContentType="application/xml"/>')
    [void]$ct.Append('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>')
    [void]$ct.Append('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')
    for ($i = 1; $i -le $count; $i++) {
        [void]$ct.Append('<Override PartName="/xl/worksheets/sheet' + $i + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
    }
    [void]$ct.Append('</Types>')

    $rootRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>
'@

    $wb = New-Object System.Text.StringBuilder
    [void]$wb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$wb.Append('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    [void]$wb.Append('<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="24000" windowHeight="14000"/></bookViews>')
    [void]$wb.Append('<sheets>')
    for ($i = 1; $i -le $count; $i++) {
        $name = Escape-Xml ([string]$Sheets[$i - 1].Name)
        [void]$wb.Append('<sheet name="' + $name + '" sheetId="' + $i + '" r:id="rId' + $i + '"/>')
    }
    [void]$wb.Append('</sheets>')
    [void]$wb.Append('</workbook>')

    $rel = New-Object System.Text.StringBuilder
    [void]$rel.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$rel.Append('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
    for ($i = 1; $i -le $count; $i++) {
        [void]$rel.Append('<Relationship Id="rId' + $i + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + $i + '.xml"/>')
    }
    [void]$rel.Append('<Relationship Id="rId' + ($count + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')
    [void]$rel.Append('</Relationships>')

    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            Add-ZipText $zip '[Content_Types].xml' $ct.ToString()
            Add-ZipText $zip '_rels/.rels' $rootRels
            Add-ZipText $zip 'xl/workbook.xml' $wb.ToString()
            Add-ZipText $zip 'xl/_rels/workbook.xml.rels' $rel.ToString()
            Add-ZipText $zip 'xl/styles.xml' (Get-StylesXml)
            for ($i = 1; $i -le $count; $i++) {
                $s = $Sheets[$i - 1]
                $xml = New-SheetXml -Columns $s.Columns -Rows $s.Rows -Selected ($i -eq 1)
                Add-ZipText $zip ('xl/worksheets/sheet' + $i + '.xml') $xml
            }
        } finally {
            $zip.Dispose()
        }
    } finally {
        $fs.Dispose()
    }
}

# ---------------------------------------------------------------- CSV

function ConvertTo-CsvField([string]$value) {
    if ($null -eq $value) { return '' }
    if ($value -match '[",\r\n]') {
        return '"' + $value.Replace('"', '""') + '"'
    }
    return $value
}

# ================================================================ 主流程

$startTime = Get-Date

Write-Host ''
Write-Host '  ============================================' -ForegroundColor Magenta
Write-Host '   vivid/stasis  全成绩导出工具' -ForegroundColor Magenta
Write-Host '  ============================================' -ForegroundColor Magenta

# ---- 1. 目录 ----
Write-Head '[1/5] 定位存档与游戏目录'

if (-not $SaveDir) { $SaveDir = Find-SaveDir }
if (-not $SaveDir) {
    throw "未找到 vivid/stasis 存档目录（%LOCALAPPDATA%\VIVIDSTASIS）。请用 -SaveDir 指定。"
}
if (-not (Test-SaveDir $SaveDir)) {
    throw "存档目录无效：$SaveDir（缺少 highscore_table）"
}
$SaveDir = (Resolve-Path -LiteralPath $SaveDir).Path

if (-not $GameDir) { $GameDir = Find-GameDir }
if ($GameDir -and -not (Test-Path -LiteralPath (Join-Path $GameDir 'song_information.bin'))) {
    Write-Warn2 ("指定的游戏目录无效（缺少 song_information.bin）：" + $GameDir)
    $GameDir = $null
}
if ($GameDir) {
    $GameDir = (Resolve-Path -LiteralPath $GameDir).Path
    Write-Ok "存档目录: $SaveDir"
    Write-Ok "游戏目录: $GameDir"
} else {
    Write-Ok "存档目录: $SaveDir"
    Write-Warn2 '未找到游戏安装目录：将只能导出分数，缺少曲名/等级等资料。可用 -GameDir 手动指定。'
}

$running = @(Get-Process -Name 'VIVIDSTASIS', 'otherside' -ErrorAction SilentlyContinue)
if ($running.Count -gt 0) {
    Write-Warn2 '检测到 vivid/stasis 正在运行：存档可能尚未完全写入，建议关闭游戏后再导出。'
}

if (-not $OutDir) { $OutDir = Join-Path $PSScriptRoot '导出结果' }
try {
    if (-not (Test-Path -LiteralPath $OutDir)) {
        New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
    }
} catch {
    $fallback = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'vivid-stasis-成绩'
    Write-Warn2 ('无法写入 ' + $OutDir + '，改用 ' + $fallback)
    $OutDir = $fallback
    New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
}
$OutDir = (Resolve-Path -LiteralPath $OutDir).Path
Write-Ok "导出目录: $OutDir"

# ---- 2. 读取数据 ----
Write-Head '[2/5] 读取游戏数据'

$hsPath  = Join-Path $SaveDir 'highscore_table'
$chPath  = Join-Path $SaveDir 'custom_highscore'
$pfPath  = Join-Path $SaveDir 'profile'

$hs = Read-IniData $hsPath
$ch = @{}
if ((-not $NoCustomSongs) -and (Test-Path -LiteralPath $chPath)) {
    $ch = Read-IniData $chPath
}

$profile = @{}
if (Test-Path -LiteralPath $pfPath) {
    $tmp = Read-IniData $pfPath
    if ($tmp.ContainsKey('profile')) { $profile = $tmp['profile'] }
}
$playerName = [string]$profile['name']
$rating     = 0.0
if ($profile.ContainsKey('ratinglast')) { $rating = (ToNum $profile['ratinglast']) / 1000.0 }
$classLevel = [string]$profile['highest_class_v3']
$classLevelText = $classLevel
if ($classLevel -match '^-?\d+(\.\d+)?$') {
    $classLevelText = Format-Num (ToNum $classLevel) '0.###'
}

$songs = @()
$gameVersion = ''
if ($GameDir) {
    $songs = @(Read-SongInfo (Join-Path $GameDir 'song_information.bin'))
    $verPath = Join-Path $GameDir 'ver'
    if (Test-Path -LiteralPath $verPath) { $gameVersion = (Read-Text $verPath).Trim() }
}
Write-Info ("曲目资料: " + $songs.Count + " 首")

$songById = @{}
foreach ($s in $songs) {
    if ($s.id -gt 0) { $songById[[int]$s.id] = $s }
}

# 谱面统计：存档目录优先，其次游戏目录
$stats = @{}
Add-StatsFromFolder $stats (Join-Path $SaveDir 'Charts')
if ($GameDir) { Add-StatsFromFolder $stats (Join-Path $GameDir 'Charts') }
Write-Info ("官方谱面统计: " + $stats.Count + " 个")

$customStats = @{}
Add-CustomStatsFromFolder $customStats (Join-Path $SaveDir 'Custom Songs')
if ($GameDir) { Add-CustomStatsFromFolder $customStats (Join-Path $GameDir 'Custom Songs') }

$customMeta = @{}
if ($GameDir) { $customMeta = Read-CustomSongInfo (Join-Path $GameDir 'Custom Songs') }
if ($customMeta.Count -eq 0) {
    $customMeta = Read-CustomSongInfo (Join-Path $SaveDir 'Custom Songs')
}
Write-Info ("自制谱面资料: " + $customMeta.Count + " 个 / 统计 " + $customStats.Count + " 个")

# ---- 3. 生成成绩行 ----
Write-Head '[3/5] 整理成绩'

$officialRows = New-Object System.Collections.ArrayList
$customRows   = New-Object System.Collections.ArrayList

# 官方谱面
foreach ($sid in @($hs.Keys)) {
    if ($sid -notmatch '^\d+$') { continue }
    $section = $hs[$sid]
    if ($null -eq $section) { continue }
    $id = [int]$sid

    $song = $null
    if ($songById.ContainsKey($id)) { $song = $songById[$id] }

    $chartId = ''
    $title   = ('未知曲目 #' + $id)
    $artist  = ''
    $bpm     = ''
    if ($null -ne $song) {
        $chartId = [string]$song.chart
        if ($song.title) { $title = $song.title }
        $artist = [string]$song.artist
        $bpm    = [string]$song.bpm
    }

    foreach ($diff in (Get-DiffsFromSection $section)) {
        $score = ToNum $section[$diff]
        $game  = ToNum $section[$diff + '_GAME']
        if ($score -le 0 -and $game -le 0) { continue }
        $diffLabel = $diff
        if (-not $diffLabel) { $diffLabel = '(未知难度)' }

        $expRaw = ToNum $section[$diff + '_GAMEPERC']
        $expPct = $expRaw * 100.0
        $maxCombo = ToNum $section[$diff + '_MAXCOMBO']
        $comboFlag = ToNum $section[$diff + '_COMBO']

        $level = ''
        $const = $null
        $designer = ''
        $order = Get-DiffOrder $diff
        if ($order -lt 10 -and $null -ne $song -and $song.charts.Count -gt $order) {
            $slot = $song.charts[$order]
            $level = [string]$slot.level
            if ($null -ne $slot.const) { $const = [double]$slot.const }
            $designer = [string]$slot.designer
        }

        $notes = 0
        if ($chartId) {
            $key = $chartId + '|' + $diff
            if ($stats.ContainsKey($key)) { $notes = [int]$stats[$key] }
        }
        if ($notes -le 0 -and $game -gt 0 -and $expPct -gt 0) {
            $notes = [int][math]::Round($game / (3.0 * $expPct / 100.0))
        }
        if ($expPct -le 0 -and $game -gt 0 -and $notes -gt 0) {
            $expPct = $game / (3.0 * $notes) * 100.0
        }

        [void]$officialRows.Add((New-ScoreRow -Type '官方' -Title $title -Artist $artist -Diff $diffLabel `
            -Level $level -Const $const -Designer $designer `
            -Score $score -Game $game -ExpPct $expPct -MaxCombo $maxCombo -ComboFlag $comboFlag `
            -Notes $notes -ChartId $chartId -SongId $id -Bpm $bpm))
    }
}

# 自制谱面
foreach ($cid in @($ch.Keys)) {
    $section = $ch[$cid]
    if ($null -eq $section) { continue }
    $meta = $null
    if ($customMeta.ContainsKey($cid)) { $meta = $customMeta[$cid] }

    $title  = [string]$cid
    $artist = ''
    $bpm    = ''
    if ($null -ne $meta) {
        if ($meta.title)  { $title  = $meta.title }
        $artist = [string]$meta.artist
        $bpm    = [string]$meta.bpm
    }

    foreach ($diff in (Get-DiffsFromSection $section)) {
        $score = ToNum $section[$diff]
        $game  = ToNum $section[$diff + '_GAME']
        if ($score -le 0 -and $game -le 0) { continue }
        $diffLabel = $diff
        if (-not $diffLabel) { $diffLabel = '(未知难度)' }

        $expRaw = ToNum $section[$diff + '_GAMEPERC']
        $expPct = $expRaw * 100.0
        $maxCombo = ToNum $section[$diff + '_MAXCOMBO']
        $comboFlag = ToNum $section[$diff + '_COMBO']

        $level = ''
        $const = $null
        $designer = ''
        $order = Get-DiffOrder $diff
        if ($null -ne $meta) {
            if ($order -lt 10 -and $meta.diffs.ContainsKey($order + 1)) {
                $slot = $meta.diffs[$order + 1]
                $level = [string]$slot.level
                if ($null -ne $slot.const) { $const = [double]$slot.const }
                $designer = [string]$slot.designer
            }
            elseif ($meta.special.ContainsKey($diff)) {
                $slot = $meta.special[$diff]
                $level = [string]$slot.level
                $designer = [string]$slot.designer
            }
        }

        $notes = 0
        $key = ([string]$cid) + '|' + $diff
        if ($customStats.ContainsKey($key)) { $notes = [int]$customStats[$key] }
        if ($notes -le 0 -and $game -gt 0 -and $expPct -gt 0) {
            $notes = [int][math]::Round($game / (3.0 * $expPct / 100.0))
        }
        if ($expPct -le 0 -and $game -gt 0 -and $notes -gt 0) {
            $expPct = $game / (3.0 * $notes) * 100.0
        }

        [void]$customRows.Add((New-ScoreRow -Type '自制' -Title $title -Artist $artist -Diff $diffLabel `
            -Level $level -Const $const -Designer $designer `
            -Score $score -Game $game -ExpPct $expPct -MaxCombo $maxCombo -ComboFlag $comboFlag `
            -Notes $notes -ChartId ([string]$cid) -SongId '' -Bpm $bpm))
    }
}

# 排序
$officialSorted = @($officialRows | Sort-Object `
    @{ Expression = { $_.songId } }, `
    @{ Expression = { Get-DiffOrder $_.diff } }, `
    @{ Expression = { $_.diff } })
$customSorted = @($customRows | Sort-Object `
    @{ Expression = { $_.title } }, `
    @{ Expression = { Get-DiffOrder $_.diff } }, `
    @{ Expression = { $_.diff } })

$allRows = @($officialSorted) + @($customSorted)

Write-Ok ("官方谱面成绩: " + $officialSorted.Count + " 条")
Write-Ok ("自制谱面成绩: " + $customSorted.Count + " 条")

if ($allRows.Count -eq 0) {
    throw '没有读取到任何成绩记录，请确认存档目录是否正确。'
}

# ---- 4. 统计 ----
Write-Head '[4/5] 汇总统计'

$scoreSum = 0.0
$expSum = 0.0
$expCount = 0
$fcCount = 0
$acCount = 0
$ssCount = 0
$best = $null
foreach ($r in $allRows) {
    if ($r.score -gt 0) {
        $scoreSum += $r.score
        if ($r.score -ge 1000000) { $ssCount++ }
        if ($null -eq $best -or $r.score -gt $best.score) { $best = $r }
    }
    if ($r.exp -gt 0) { $expSum += $r.exp; $expCount++ }
    if ($r.status -eq 'FC') { $fcCount++ }
    if ($r.status -eq 'AC') { $acCount++ }
}
$expAvg = 0.0
if ($expCount -gt 0) { $expAvg = $expSum / $expCount }

# 难度统计
$diffStats = @{}
foreach ($r in $allRows) {
    $d = $r.diff
    if (-not $diffStats.ContainsKey($d)) {
        $diffStats[$d] = @{ count = 0; scoreSum = 0.0; scoreCount = 0; expSum = 0.0; expCount = 0 }
    }
    $item = $diffStats[$d]
    $item.count++
    if ($r.score -gt 0) { $item.scoreSum += $r.score; $item.scoreCount++ }
    if ($r.exp -gt 0) { $item.expSum += $r.exp; $item.expCount++ }
}

# 段位挑战
$courseRows = @()
if ($hs.ContainsKey('COURSE')) {
    $courseSection = $hs['COURSE']
    $courseRows = @($courseSection.Keys | Sort-Object | ForEach-Object {
        $v = ToNum $courseSection[$_]
        if ($v -le 0) { return }
        $label = [string]$_
        if ($label -match '^class(\d+)v$') { $label = 'Class ' + $Matches[1] }
        @{ name = $label; score = $v }
    })
}

# ---- 5. 输出 ----
Write-Head '[5/5] 写出文件'

$stamp = $startTime.ToString('yyyyMMdd_HHmmss')
$baseName = 'vivid-stasis-成绩_' + $stamp
$xlsxPath = Join-Path $OutDir ($baseName + '.xlsx')
$csvPath  = Join-Path $OutDir ($baseName + '.csv')
$jsonPath = Join-Path $OutDir ($baseName + '.json')

$scoreColumns = @(
    @{ h = '序号';     w = 6;  t = 'int'  },
    @{ h = '类型';     w = 8;  t = 'str'  },
    @{ h = '曲名';     w = 34; t = 'str'  },
    @{ h = '曲师';     w = 22; t = 'str'  },
    @{ h = '难度';     w = 12; t = 'str'  },
    @{ h = '等级';     w = 8;  t = 'str'  },
    @{ h = '定数';     w = 8;  t = 'num1' },
    @{ h = '谱师';     w = 22; t = 'str'  },
    @{ h = '分数';     w = 12; t = 'int'  },
    @{ h = '评级';     w = 8;  t = 'str'  },
    @{ h = 'EX 分数';  w = 10; t = 'int'  },
    @{ h = 'EX 率';    w = 9;  t = 'num2' },
    @{ h = 'EX 评级';  w = 11; t = 'str'  },
    @{ h = '最大连击'; w = 10; t = 'int'  },
    @{ h = '音符数';   w = 9;  t = 'int'  },
    @{ h = '通关';     w = 10; t = 'str'  },
    @{ h = '谱面 ID';  w = 26; t = 'str'  },
    @{ h = '歌曲 ID';  w = 9;  t = 'int'  },
    @{ h = 'BPM';      w = 9;  t = 'str'  }
)

function Convert-RowForJson($r) {
    $const = $null
    if ($null -ne $r.const) { $const = [math]::Round([double]$r.const, 3) }
    return [ordered]@{
        type       = $r.type
        title      = $r.title
        artist     = $r.artist
        difficulty = $r.diff
        level      = $r.level
        constant   = $const
        designer   = $r.designer
        score      = [math]::Round($r.score, 3)
        grade      = $r.grade
        exScore    = [int][math]::Round($r.exScore)
        exPercent  = [math]::Round($r.exp, 4)
        exGrade    = $r.exGrade
        maxCombo   = [int][math]::Round($r.maxCombo)
        notes      = [int]$r.notes
        status     = $r.status
        chartId    = $r.chartId
        songId     = $r.songId
        bpm        = $r.bpm
    }
}

function Convert-RowsToArrays($rows) {
    $list = New-Object System.Collections.ArrayList
    $i = 0
    foreach ($r in $rows) {
        $i++
        [void]$list.Add(@(
            $i, $r.type, $r.title, $r.artist, $r.diff, $r.level, $r.const, $r.designer,
            $r.score, $r.grade, $r.exScore, $r.exp, $r.exGrade, $r.maxCombo, $r.notes,
            $r.status, $r.chartId, $r.songId, $r.bpm
        ))
    }
    return @($list)
}

$allArrays    = Convert-RowsToArrays $allRows
$officialArrs = Convert-RowsToArrays $officialSorted
$customArrs   = Convert-RowsToArrays $customSorted

# 概况
$bestText = ''
if ($null -ne $best) {
    $bestText = (Format-Num $best.score '#,##0') + '  ' + $best.title + ' [' + $best.diff + ']'
}
$summaryRows = @(
    @('导出时间', $startTime.ToString('yyyy-MM-dd HH:mm:ss')),
    @('玩家昵称', $playerName),
    @('玩家 Rating', $(if ($rating -gt 0) { (Format-Num $rating '0.###') } else { '' })),
    @('最高段位等级', $classLevelText),
    @('游戏版本', $gameVersion),
    @('存档目录', $SaveDir),
    @('游戏目录', $GameDir),
    @('', ''),
    @('官方谱面成绩数', [string]$officialSorted.Count),
    @('自制谱面成绩数', [string]$customSorted.Count),
    @('成绩总数', [string]$allRows.Count),
    @('总分数', (Format-Num $scoreSum '#,##0')),
    @('平均 EX 率', ((Format-Num $expAvg '0.00') + '%')),
    @('分数 ≥ 1,000,000', [string]$ssCount),
    @('Full Combo', [string]$fcCount),
    @('All Critical', [string]$acCount),
    @('最高分', $bestText),
    @('', ''),
    @('分数说明', '准确度得分，满分 1,010,000；评级阈值 VS / V+ / V / SS+ / SS / S+ / S / AA / A / B / C / D / E'),
    @('EX 说明', 'EX 分数每音符最高 3 分；EX 率 = EX 分数 / (3 × 音符数)'),
    @('通关标记', 'FC / AC 由存档 COMBO 字段推断，仅供参考'),
    @('数据来源', 'highscore_table、custom_highscore、song_information.bin、Charts/*.stats、Custom Songs/info.json'),
    @('工具说明', '本工具只读取存档，不会修改游戏数据')
)

$diffRows = New-Object System.Collections.ArrayList
foreach ($key in @($diffStats.Keys | Sort-Object { Get-DiffOrder $_ }, { $_ })) {
    $item = $diffStats[$key]
    $avgScore = 0.0
    if ($item.scoreCount -gt 0) { $avgScore = $item.scoreSum / $item.scoreCount }
    $avgExp = 0.0
    if ($item.expCount -gt 0) { $avgExp = $item.expSum / $item.expCount }
    [void]$diffRows.Add(@($key, $item.count, $avgScore, $avgExp))
}

$courseArrs = New-Object System.Collections.ArrayList
foreach ($c in $courseRows) {
    [void]$courseArrs.Add(@($c.name, $c.score))
}

$sheets = @(
    @{ Name = '概况';     Columns = @(@{ h = '项目'; w = 22; t = 'str' }, @{ h = '内容'; w = 96; t = 'str' }); Rows = $summaryRows },
    @{ Name = '全部成绩'; Columns = $scoreColumns; Rows = $allArrays },
    @{ Name = '官方谱面'; Columns = $scoreColumns; Rows = $officialArrs },
    @{ Name = '自制谱面'; Columns = $scoreColumns; Rows = $customArrs },
    @{ Name = '难度统计'; Columns = @(
            @{ h = '难度'; w = 14; t = 'str' },
            @{ h = '谱面数'; w = 10; t = 'int' },
            @{ h = '平均分数'; w = 14; t = 'int' },
            @{ h = '平均 EX 率'; w = 12; t = 'num2' }
        ); Rows = @($diffRows) },
    @{ Name = '段位挑战'; Columns = @(
            @{ h = '段位'; w = 16; t = 'str' },
            @{ h = '分数'; w = 14; t = 'int' }
        ); Rows = @($courseArrs) }
)

New-XlsxFile -Path $xlsxPath -Sheets $sheets
Write-Ok ("Excel: " + $xlsxPath)

# CSV
$csv = New-Object System.Text.StringBuilder
$header = @($scoreColumns | ForEach-Object { ConvertTo-CsvField $_.h })
[void]$csv.AppendLine(($header -join ','))
foreach ($row in $allArrays) {
    $fields = @()
    for ($i = 0; $i -lt $row.Count; $i++) {
        $v = $row[$i]
        if ($null -eq $v) { $fields += ''; continue }
        if ($v -is [string]) {
            $fields += ConvertTo-CsvField $v
        } else {
            $type = [string]$scoreColumns[$i].t
            $fmt = '0'
            if ($type -eq 'num2') { $fmt = '0.00' }
            elseif ($type -eq 'num1') { $fmt = '0.0' }
            $fields += ([double]$v).ToString($fmt, $script:Inv)
        }
    }
    [void]$csv.AppendLine(($fields -join ','))
}
$csvText = $csv.ToString()
[System.IO.File]::WriteAllText($csvPath, $csvText, (New-Object System.Text.UTF8Encoding($true)))
Write-Ok ("CSV: " + $csvPath)

# JSON
if (-not $NoJson) {
    $jsonOfficial = @($officialSorted | ForEach-Object { Convert-RowForJson $_ })
    $jsonCustom   = @($customSorted   | ForEach-Object { Convert-RowForJson $_ })
    $jsonCourse   = @($courseRows | ForEach-Object { [ordered]@{ name = $_.name; score = [math]::Round($_.score, 3) } })

    $jsonObj = [ordered]@{
        tool        = 'vivid-stasis-score-exporter'
        generatedAt = $startTime.ToString('yyyy-MM-dd HH:mm:ss')
        saveDir     = $SaveDir
        gameDir     = $GameDir
        gameVersion = $gameVersion
        player      = [ordered]@{ name = $playerName; rating = $rating; classLevel = $classLevelText }
        summary     = [ordered]@{
            official = $officialSorted.Count
            custom   = $customSorted.Count
            total    = $allRows.Count
            scoreSum = [math]::Round($scoreSum, 3)
            expAvg   = [math]::Round($expAvg, 3)
            fc       = $fcCount
            ac       = $acCount
        }
        official    = $jsonOfficial
        custom      = $jsonCustom
        course      = $jsonCourse
    }
    $jsonText = ConvertTo-Json -InputObject $jsonObj -Depth 6
    [System.IO.File]::WriteAllText($jsonPath, $jsonText, (New-Object System.Text.UTF8Encoding($true)))
    Write-Ok ("JSON: " + $jsonPath)
}

# ---- 控制台汇总 ----
Write-Host ''
Write-Host '  --------------------------------------------' -ForegroundColor DarkGray
Write-Host ('   成绩总数  ' + $allRows.Count + ' 条    (官方 ' + $officialSorted.Count + ' / 自制 ' + $customSorted.Count + ')') -ForegroundColor White
Write-Host ('   平均 EX   ' + (Format-Num $expAvg '0.00') + '%    FC ' + $fcCount + '    AC ' + $acCount) -ForegroundColor White
if ($null -ne $best) {
    Write-Host ('   最高分    ' + (Format-Num $best.score '#,##0') + '  ' + $best.title + ' [' + $best.diff + ']') -ForegroundColor White
}
Write-Host '  --------------------------------------------' -ForegroundColor DarkGray

$top = @($allRows | Where-Object { $_.score -gt 0 } | Sort-Object -Property @{ Expression = { $_.score } } -Descending | Select-Object -First 8)
if ($top.Count -gt 0) {
    Write-Host ''
    Write-Host '  分数 TOP 8' -ForegroundColor Cyan
    foreach ($r in $top) {
        $line = '   {0,12}  {1} [{2}]' -f (Format-Num $r.score '#,##0'), $r.title, $r.diff
        Write-Host $line -ForegroundColor Gray
    }
}

$elapsed = ((Get-Date) - $startTime).TotalSeconds
Write-Host ''
Write-Host ('  完成，用时 ' + [math]::Round($elapsed, 1) + ' 秒') -ForegroundColor Green
Write-Host ''

if ($Open) { Start-Process -FilePath 'explorer.exe' -ArgumentList ('"' + $OutDir + '"') }
