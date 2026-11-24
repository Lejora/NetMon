#Requires -Version 5.1
<#
.SYNOPSIS
  NetMon の記録から、管理会社・回線事業者に提出するレポート (HTML + CSV) を作成します。

.EXAMPLE
  .\Make-Report.ps1                       # 直近7日間
  .\Make-Report.ps1 -Days 14              # 直近14日間
  .\Make-Report.ps1 -From 2026-10-01 -To 2026-10-09
#>
[CmdletBinding()]
param(
    [int]$Days = 7,
    [datetime]$From,
    [datetime]$To,
    [double]$SlowMbps = 10,     # 下り速度がこれ未満の測定を「低速」として数える
    [string]$LogDir = '',
    [string]$OutDir = '',
    [switch]$NoOpen
)
$ErrorActionPreference = 'Stop'
if (-not $LogDir) { $LogDir = Join-Path $PSScriptRoot 'logs' }
if (-not $OutDir) { $OutDir = Join-Path $PSScriptRoot 'reports' }
if ($PSBoundParameters.ContainsKey('To')) {
    if ($To.TimeOfDay.Ticks -eq 0) { $To = $To.AddDays(1).AddSeconds(-1) }   # 日付だけ指定ならその日の終わりまで
} else { $To = Get-Date }
if (-not $PSBoundParameters.ContainsKey('From')) { $From = $To.Date.AddDays(-($Days - 1)) }
if (-not (Test-Path -LiteralPath (Join-Path $LogDir 'minute'))) { throw "記録が見つかりません: $LogDir  (先に NetMon を起動して計測してください)" }

$Inv = [Globalization.CultureInfo]::InvariantCulture
$Enc = New-Object System.Text.UTF8Encoding($true)
$NaN = [double]::NaN
$FromS = $From.ToString('yyyy-MM-dd HH:mm:ss')
$ToS = $To.ToString('yyyy-MM-dd HH:mm:ss')

# ------------------------------------------------------------
#  読み込み
# ------------------------------------------------------------
function Read-SharedText([string]$Path) {
    # 監視中で書き込み中のファイルも読めるよう共有モードで開く
    $fs = New-Object IO.FileStream($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    $sr = New-Object IO.StreamReader($fs, [Text.Encoding]::UTF8, $true)
    try { return $sr.ReadToEnd() } finally { $sr.Dispose() }
}

$CsvRx = [regex]'(?:^|,)(?:"((?:[^"]|"")*)"|([^,]*))'
function Split-Csv([string]$Line) {
    if ($Line.IndexOf('"') -lt 0) { return , $Line.Split(',') }
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($m in $CsvRx.Matches($Line)) {
        if ($m.Groups[1].Success) { $out.Add($m.Groups[1].Value.Replace('""', '"')) } else { $out.Add($m.Groups[2].Value) }
    }
    return , $out.ToArray()
}

# CSV を読み、期間内の行だけを返す (各行は 列名→値 のハッシュテーブル)
function Import-LogCsv([string[]]$Paths, [string]$TimeColumn) {
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($p in $Paths) {
        $lines = (Read-SharedText $p) -split "`r?`n"
        if ($lines.Count -lt 2) { continue }
        $hdr = Split-Csv $lines[0]
        $ti = [array]::IndexOf($hdr, $TimeColumn)
        for ($i = 1; $i -lt $lines.Count; $i++) {
            if (-not $lines[$i]) { continue }
            $f = Split-Csv $lines[$i]
            $t = $f[$ti]
            if ($t.Length -eq 16) { $t += ':00' }
            if ([string]::CompareOrdinal($t, $FromS) -lt 0 -or [string]::CompareOrdinal($t, $ToS) -gt 0) { continue }
            $h = @{}
            for ($k = 0; $k -lt $hdr.Count -and $k -lt $f.Count; $k++) { $h[$hdr[$k]] = $f[$k] }
            $rows.Add($h)
        }
    }
    return , $rows
}

function Get-DayFiles([string]$Pattern) {
    $d0 = $From.ToString('yyyyMMdd'); $d1 = $To.ToString('yyyyMMdd')
    @(Get-ChildItem -LiteralPath (Join-Path $LogDir 'minute') -Filter $Pattern | Where-Object {
        $d = $_.BaseName.Split('_')[1]; ($d -ge $d0) -and ($d -le $d1)
    } | Sort-Object Name | ForEach-Object { $_.FullName })
}

function Num($s) {
    if ([string]::IsNullOrEmpty($s)) { return $NaN }
    return [double]::Parse($s, $Inv)
}

$EnvInfo = $null
$envPath = Join-Path $LogDir 'environment.json'
if (Test-Path -LiteralPath $envPath) { $EnvInfo = (Read-SharedText $envPath) | ConvertFrom-Json }
$Thr = 2.0; $LowMbps = 5.0
if ($EnvInfo) { $Thr = [double]$EnvInfo.loss_threshold_pct; $LowMbps = [double]$EnvInfo.low_traffic_mbps }

Write-Host '記録を読み込んでいます...'
$Overview = Import-LogCsv (Get-DayFiles 'overview_*.csv') 'time'
$Incidents = New-Object System.Collections.Generic.List[object]
$incPath = Join-Path $LogDir 'incidents.csv'
if (Test-Path -LiteralPath $incPath) { $Incidents = Import-LogCsv @($incPath) 'start' }
$Speeds = New-Object System.Collections.Generic.List[object]
$spdPath = Join-Path $LogDir 'speedtest.csv'
if (Test-Path -LiteralPath $spdPath) { $Speeds = Import-LogCsv @($spdPath) 'time' }
if ($Overview.Count -eq 0) { throw "指定期間 ($FromS 〜 $ToS) の記録がありません。" }

# ------------------------------------------------------------
#  区間(監視対象)ごとの集計  ※行数が多いので逐次集計
# ------------------------------------------------------------
$TargetAgg = [ordered]@{}
$summaryRows = New-Object System.Collections.Generic.List[string]
foreach ($p in (Get-DayFiles 'summary_*.csv')) {
    $lines = (Read-SharedText $p) -split "`r?`n"
    if ($summaryRows.Count -eq 0 -and $lines.Count) { $summaryRows.Add($lines[0]) }
    for ($i = 1; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        if (-not $line) { continue }
        $f = Split-Csv $line
        $t = $f[0] + ':00'
        if ([string]::CompareOrdinal($t, $FromS) -lt 0 -or [string]::CompareOrdinal($f[0], $ToS) -gt 0) { continue }
        $summaryRows.Add($line)
        # time,id,layer,name,address,kind,sent,lost,loss_pct,rtt_min,rtt_avg,rtt_p95,rtt_max,jitter,other_responder
        $key = $f[1] + '|' + $f[4]
        $a = $TargetAgg[$key]
        if (-not $a) {
            $a = @{ Id = $f[1]; Layer = $f[2]; Name = $f[3]; Address = $f[4]; Sent = 0; Lost = 0; RttSum = 0.0; RttN = 0; RttMax = 0.0; BadMin = 0; Min = 0 }
            $TargetAgg[$key] = $a
        }
        $s = [int]$f[6]; $l = [int]$f[7]
        if ($s -eq 0) { continue }
        $a.Sent += $s; $a.Lost += $l; $a.Min++
        if ($f[10]) { $a.RttSum += [double]::Parse($f[10], $Inv) * ($s - $l); $a.RttN += ($s - $l) }
        if ($f[12]) { $mx = [double]::Parse($f[12], $Inv); if ($mx -gt $a.RttMax) { $a.RttMax = $mx } }
        if (100.0 * $l / $s -ge $Thr) { $a.BadMin++ }
    }
}

# ------------------------------------------------------------
#  1分ごとの概要から全体・時間帯別に集計
# ------------------------------------------------------------
$Tot = @{ Min = 0; Measured = 0; ES = 0; EL = 0; GS = 0; GL = 0; BadMin = 0; LowES = 0; LowEL = 0; LowMin = 0; HighRttMin = 0; Wifi = $false }
$Cells = @{}      # 'yyyy-MM-dd|H' -> 集計
$Dates = New-Object System.Collections.Generic.SortedSet[string]
$HourES = New-Object double[] 24; $HourEL = New-Object double[] 24
$HourRtt = New-Object double[] 24; $HourRttN = New-Object double[] 24
$BySection = @{}
$MinuteAvgs = New-Object System.Collections.Generic.List[double]
$Worst = New-Object System.Collections.Generic.List[object]

foreach ($r in $Overview) {
    $time = $r['time']
    $date = $time.Substring(0, 10); $hour = [int]$time.Substring(11, 2)
    [void]$Dates.Add($date)
    $es = [int]$r['ext_sent']; $el = [int]$r['ext_lost']; $gs = [int]$r['gw_sent']; $gl = [int]$r['gw_lost']
    $rtt = Num $r['ext_rtt_avg']; $rx = Num $r['rx_mbps_avg']; $tx = Num $r['tx_mbps_avg']
    $Tot.Min++; $Tot.Measured += [int]$r['measured_s']
    $Tot.ES += $es; $Tot.EL += $el; $Tot.GS += $gs; $Tot.GL += $gl
    if ($r['wifi_signal']) { $Tot.Wifi = $true }
    if ($r['rtt_high'] -eq '1') { $Tot.HighRttMin++ }
    $pct = $NaN; if ($es -gt 0) { $pct = 100.0 * $el / $es }
    if ($es -gt 0 -and $pct -ge $Thr) { $Tot.BadMin++ }
    if ($rx -lt $LowMbps -and $tx -lt $LowMbps) { $Tot.LowES += $es; $Tot.LowEL += $el; $Tot.LowMin++ }

    $key = '{0}|{1}' -f $date, $hour
    $c = $Cells[$key]
    if (-not $c) { $c = @{ ES = 0; EL = 0; GS = 0; GL = 0; RttSum = 0.0; RttN = 0; Down = (New-Object System.Collections.Generic.List[double]) }; $Cells[$key] = $c }
    $c.ES += $es; $c.EL += $el; $c.GS += $gs; $c.GL += $gl
    $HourES[$hour] += $es; $HourEL[$hour] += $el
    if (-not [double]::IsNaN($rtt)) {
        $ok = $es - $el
        $c.RttSum += $rtt * $ok; $c.RttN += $ok
        $HourRtt[$hour] += $rtt * $ok; $HourRttN[$hour] += $ok
        if ($pct -lt $Thr) { $MinuteAvgs.Add($rtt) }
    }

    $code = $r['verdict_code']
    if ($code -and $code -ne 'OK' -and $code -ne 'NODATA' -and $code -ne 'ICMPLIMIT') {
        $sec = $r['verdict_section']
        if (-not $BySection.ContainsKey($sec)) { $BySection[$sec] = @{ Min = 0; Lost = 0; Code = $code } }
        $BySection[$sec].Min++; $BySection[$sec].Lost += $el
    }
    if ($es -ge 20 -and $el -gt 0) { $Worst.Add([pscustomobject]@{ Time = $time; Pct = $pct; ES = $es; EL = $el; GS = $gs; GL = $gl; Verdict = $r['verdict']; Rx = $rx; Rtt = $rtt }) }
}

$RttBase = $NaN
if ($MinuteAvgs.Count) { $b = $MinuteAvgs.ToArray(); [Array]::Sort($b); $RttBase = $b[[int][Math]::Floor($b.Length * 0.1)] }

# 速度測定
$SpeedOk = New-Object System.Collections.Generic.List[object]
foreach ($s in $Speeds) {
    $down = Num $s['down_mbps']; $up = Num $s['up_mbps']
    $o = [pscustomobject]@{ Time = $s['time']; Down = $down; Up = $up; BgRx = (Num $s['bg_rx_mbps']); Idle = (Num $s['idle_rtt_ms']); Loaded = (Num $s['loaded_rtt_ms']); Error = $s['error'] }
    if ($down -ge 0) {
        $SpeedOk.Add($o)
        $key = '{0}|{1}' -f $o.Time.Substring(0, 10), [int]$o.Time.Substring(11, 2)
        if ($Cells.ContainsKey($key)) { $Cells[$key].Down.Add($down) }
    }
}
$HourDown = New-Object double[] 24; $HourDownN = New-Object double[] 24
foreach ($o in $SpeedOk) { $h = [int]$o.Time.Substring(11, 2); $HourDown[$h] += $o.Down; $HourDownN[$h]++ }

function Get-Median([double[]]$v) {
    $a = @($v | Where-Object { -not [double]::IsNaN($_) -and $_ -ge 0 })
    if ($a.Count -eq 0) { return $NaN }
    [Array]::Sort($a)
    if ($a.Count % 2) { return $a[($a.Count - 1) / 2] }
    return ($a[$a.Count / 2 - 1] + $a[$a.Count / 2]) / 2
}
$DownMed = Get-Median ($SpeedOk | ForEach-Object { $_.Down })
$UpMed = Get-Median ($SpeedOk | ForEach-Object { $_.Up })

# ------------------------------------------------------------
#  HTML 部品
# ------------------------------------------------------------
function Esc([string]$s) { [System.Net.WebUtility]::HtmlEncode($s) }
function F([double]$v, [string]$fmt = '0.0') { if ([double]::IsNaN($v)) { return '-' } else { return $v.ToString($fmt, $Inv) } }
function Pct([double]$num, [double]$den, [string]$fmt = '0.00') { if ($den -le 0) { return $NaN } else { return 100.0 * $num / $den } }

function Get-LossBg([double]$p) {
    if ([double]::IsNaN($p)) { return '#eeeeee' }
    if ($p -le 0) { return '#cfeedd' }
    if ($p -lt 0.5) { return '#eef3b8' }
    if ($p -lt $Thr) { return '#ffd978' }
    if ($p -lt 5) { return '#ffa257' }
    if ($p -lt 15) { return '#ee5a45' }
    return '#a3122a'
}
function Get-RatioBg([double]$ratio, [bool]$HigherIsBetter) {
    if ([double]::IsNaN($ratio)) { return '#eeeeee' }
    if ($HigherIsBetter) { if ($ratio -ge 0.7) { return '#cfeedd' }; if ($ratio -ge 0.4) { return '#ffd978' }; if ($ratio -ge 0.2) { return '#ffa257' }; return '#ee5a45' }
    if ($ratio -le 1.3) { return '#cfeedd' }; if ($ratio -le 2) { return '#ffd978' }; if ($ratio -le 3) { return '#ffa257' }; return '#ee5a45'
}
function Get-Fg([string]$bg) { if ($bg -in @('#ee5a45', '#a3122a')) { return '#fff' } return '#222' }

function New-Heatmap([string]$Title, [string]$Kind) {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<h3>$(Esc $Title)</h3><div class='hm-wrap'><table class='hm'><tr><th>日付</th>")
    for ($h = 0; $h -lt 24; $h++) { [void]$sb.Append("<th>$h</th>") }
    [void]$sb.Append('</tr>')
    foreach ($d in $Dates) {
        $dt = [datetime]::ParseExact($d, 'yyyy-MM-dd', $Inv)
        $wd = '日月火水木金土'[[int]$dt.DayOfWeek]
        [void]$sb.Append("<tr><th>$($dt.ToString('M/d'))($wd)</th>")
        for ($h = 0; $h -lt 24; $h++) {
            $c = $Cells["$d|$h"]
            $val = $NaN; $txt = ''; $tip = ''
            if ($c) {
                switch ($Kind) {
                    'ext' { $val = Pct $c.EL $c.ES; $bg = Get-LossBg $val; $txt = F $val '0.0'; $tip = "外部宛先 $($c.EL)/$($c.ES) 発損失" }
                    'gw' { $val = Pct $c.GL $c.GS; $bg = Get-LossBg $val; $txt = F $val '0.0'; $tip = "宅内ルーター $($c.GL)/$($c.GS) 発損失" }
                    'rtt' { if ($c.RttN) { $val = $c.RttSum / $c.RttN }; $bg = Get-RatioBg ($val / $RttBase) $false; $txt = F $val '0'; $tip = "平均 $(F $val) ms" }
                    'down' { if ($c.Down.Count) { $val = ($c.Down | Measure-Object -Average).Average }; $bg = Get-RatioBg ($val / $DownMed) $true; $txt = F $val '0'; $tip = "下り平均 $(F $val) Mbps ($($c.Down.Count)回)" }
                }
            }
            if ([double]::IsNaN($val)) { [void]$sb.Append("<td class='na'></td>") }
            else { [void]$sb.Append("<td style='background:$bg;color:$(Get-Fg $bg)' title='$(Esc $tip)'>$txt</td>") }
        }
        [void]$sb.Append('</tr>')
    }
    [void]$sb.Append('</table></div>')
    $sb.ToString()
}

function New-BarSvg([double[]]$Vals, [string]$Unit, [string]$Color, [double]$Ref = [double]::NaN, [string]$RefLabel = '') {
    $w = 760; $h = 200; $pl = 48; $pr = 12; $pt = 26; $pb = 26
    $max = 0.0
    foreach ($v in $Vals) { if (-not [double]::IsNaN($v) -and $v -gt $max) { $max = $v } }
    if (-not [double]::IsNaN($Ref) -and $Ref -gt $max) { $max = $Ref }
    if ($max -le 0) { $max = 1 }
    $max *= 1.15
    $bw = ($w - $pl - $pr) / 24.0
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<svg viewBox='0 0 $w $h' class='bar' role='img'>")
    foreach ($frac in 0, 0.5, 1) {
        $y = $pt + ($h - $pt - $pb) * (1 - $frac)
        [void]$sb.Append("<line x1='$pl' x2='$($w-$pr)' y1='$(F $y)' y2='$(F $y)' stroke='#ddd'/>")
        [void]$sb.Append("<text x='$($pl-6)' y='$(F ($y+4))' text-anchor='end' class='ax'>$(F ($max*$frac) '0.#')</text>")
    }
    for ($i = 0; $i -lt 24; $i++) {
        $x = $pl + $i * $bw
        [void]$sb.Append("<text x='$(F ($x+$bw/2))' y='$($h-8)' text-anchor='middle' class='ax'>$i</text>")
        $v = $Vals[$i]
        if ([double]::IsNaN($v)) { continue }
        $bh = ($h - $pt - $pb) * $v / $max
        [void]$sb.Append("<rect x='$(F ($x+2))' y='$(F ($h-$pb-$bh))' width='$(F ($bw-4))' height='$(F $bh)' fill='$Color'><title>${i}時台: $(F $v '0.00') $Unit</title></rect>")
    }
    if (-not [double]::IsNaN($Ref)) {
        $y = $pt + ($h - $pt - $pb) * (1 - $Ref / $max)
        [void]$sb.Append("<line x1='$pl' x2='$($w-$pr)' y1='$(F $y)' y2='$(F $y)' stroke='#c0392b' stroke-dasharray='5 4'/>")
        [void]$sb.Append("<text x='$($w-$pr)' y='$(F ($y-4))' text-anchor='end' class='ref'>$(Esc $RefLabel)</text>")
    }
    [void]$sb.Append("<text x='4' y='12' class='ax'>$(Esc $Unit)</text></svg>")
    $sb.ToString()
}

# ------------------------------------------------------------
#  結論 (自動生成)
# ------------------------------------------------------------
$extPct = Pct $Tot.EL $Tot.ES
$gwPct = Pct $Tot.GL $Tot.GS
$lowPct = Pct $Tot.LowEL $Tot.LowES
$hours = $Tot.Measured / 3600.0
$Findings = New-Object System.Collections.Generic.List[string]

$Findings.Add(('計測期間 {0} 〜 {1}、有効な計測時間は <b>{2}</b>です (1秒ごとに外部サーバへ ping を送信し、計 {3:N0} 発)。' -f $From.ToString('yyyy/MM/dd HH:mm'), $To.ToString('yyyy/MM/dd HH:mm'), $(if ($hours -ge 1) { '{0:0.0} 時間' -f $hours } else { '{0:0} 分' -f ($hours * 60) }), $Tot.ES))
$Findings.Add(('外部サーバへのパケットロス率は <b class="bad">{0}%</b> ({1:N0} 発が消失) でした。ロス率が {2}% 以上だった時間は <b>{3} 分</b> (計測時間の {4}%) あります。' -f (F $extPct '0.00'), $Tot.EL, $Thr, $Tot.BadMin, (F (Pct $Tot.BadMin $Tot.Min) '0.0')))
$g = '宅内ルーター (このPCが直接つながっている機器) までのロス率は <b>{0}%</b> です。' -f (F $gwPct '0.00')
if ($gwPct -lt 0.1 -and $extPct -ge 0.3) { $g += ' <b>宅内の LAN・PC・ルーターは正常で、ロスは宅外 (マンション設備または回線事業者側) で発生</b>しています。' }
elseif ($gwPct -ge $Thr) { $g += ' 宅内でもロスが出ているため、LANケーブル・ルーター・PC 側も確認が必要です。' }
$Findings.Add($g)

$secList = @($BySection.GetEnumerator() | Sort-Object { $_.Value.Min } -Descending)
$secTotal = ($secList | ForEach-Object { $_.Value.Min } | Measure-Object -Sum).Sum
if ($secList.Count -gt 0) {
    $top = $secList[0]
    $Findings.Add(('ロスが発生した時間のうち <b>{0}%</b> ({1} 分) で、ロスの起点は <b class="bad">「{2}」</b> と判定されました (詳細は「ロス発生箇所の判定」)。' -f (F (Pct $top.Value.Min $secTotal) '0'), $top.Value.Min, (Esc $top.Key)))
}
if ($Tot.LowES -gt 0) {
    $t = 'このPCの通信量が {0} Mbps 未満 (ほぼ通信していない) の時間帯 {1} 分に限っても、外部へのロス率は <b>{2}%</b> です。' -f $LowMbps, $Tot.LowMin, (F $lowPct '0.00')
    if ($lowPct -ge 0.3) { $t += ' <b>利用者側の通信過多が原因ではありません</b>。' }
    $Findings.Add($t)
}
$hp = for ($h = 0; $h -lt 24; $h++) { [pscustomobject]@{ H = $h; P = (Pct $HourEL[$h] $HourES[$h]); N = $HourES[$h] } }
$worstHour = $hp | Where-Object { $_.N -ge 600 } | Sort-Object P -Descending | Select-Object -First 1
$bestHour = $hp | Where-Object { $_.N -ge 600 } | Sort-Object P | Select-Object -First 1
if ($worstHour -and $worstHour.P -gt 0) {
    $Findings.Add(('時間帯別では <b>{0}時台</b> が最も悪く、ロス率 {1}% (最も良い {2}時台は {3}%)。' -f $worstHour.H, (F $worstHour.P '0.00'), $bestHour.H, (F $bestHour.P '0.00')))
}
if (-not [double]::IsNaN($RttBase) -and $Tot.HighRttMin -gt 0) {
    $Findings.Add(('平常時の応答時間 (約 {0} ms) の2倍以上に遅くなった時間が {1} 分ありました。' -f (F $RttBase), $Tot.HighRttMin))
}
if ($SpeedOk.Count) {
    $mn = ($SpeedOk | Measure-Object Down -Minimum).Minimum
    $slow = @($SpeedOk | Where-Object { $_.Down -lt $SlowMbps }).Count
    $Findings.Add(('速度測定 {0} 回: 下り 中央値 <b>{1} Mbps</b> / 最低 {2} Mbps、上り 中央値 {3} Mbps。下り {4} Mbps 未満は <b>{5} 回</b>。' -f $SpeedOk.Count, (F $DownMed), (F $mn), (F $UpMed), $SlowMbps, $slow))
}
if ($Incidents.Count) {
    $outages = @($Incidents | Where-Object { [int]$_['max_full_outage_s'] -ge 3 })
    $longest = ($Incidents | ForEach-Object { [int]$_['max_full_outage_s'] } | Measure-Object -Maximum).Maximum
    $Findings.Add(('連続したロス (ロス事象) は {0} 件。うち全宛先が3秒以上応答しない完全断は {1} 件 (最長 {2} 秒)。' -f $Incidents.Count, $outages.Count, $longest))
}
if ($Tot.Wifi -or ($EnvInfo -and $EnvInfo.nic_type -eq 'Wi-Fi')) {
    $Findings.Add('<span class="warn">注意: 計測PCが Wi-Fi 接続の期間があります。Wi-Fi 自体の不調と区別するため、有線 LAN での計測を推奨します。</span>')
}

# ------------------------------------------------------------
#  HTML 組み立て
# ------------------------------------------------------------
$stamp = Get-Date
$name = 'report_' + $From.ToString('yyyyMMdd') + '-' + $To.ToString('yyyyMMdd') + '_' + $stamp.ToString('HHmmss')
$dir = Join-Path $OutDir $name
[void][IO.Directory]::CreateDirectory($dir)

$html = New-Object System.Text.StringBuilder
function A([string]$s) { [void]$script:html.Append($s) }

A @"
<!doctype html><html lang="ja"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>回線品質レポート</title>
<style>
:root{--fg:#1d2330;--mut:#5d6575;--line:#d9dde5;--bg:#fff;--soft:#f5f7fa;--accent:#1f5fa8;--bad:#c0392b}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--fg);font:14px/1.65 "Yu Gothic UI","Meiryo",system-ui,sans-serif}
main{max-width:1100px;margin:0 auto;padding:28px 20px 60px}
h1{font-size:24px;margin:0 0 4px}
h2{font-size:18px;margin:36px 0 10px;padding-bottom:6px;border-bottom:2px solid var(--accent)}
h3{font-size:15px;margin:22px 0 6px}
.sub{color:var(--mut);margin:0}
.box{background:var(--soft);border:1px solid var(--line);border-radius:8px;padding:14px 18px}
.box li{margin:4px 0}
.bad{color:var(--bad)}
.warn{color:#9a6200}
table.t{border-collapse:collapse;width:100%;margin:6px 0;font-size:13px}
table.t th,table.t td{border:1px solid var(--line);padding:4px 7px;text-align:left;vertical-align:top}
table.t th{background:var(--soft);font-weight:600}
table.t td.n{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
.hm-wrap{overflow-x:auto}
table.hm{border-collapse:collapse;font-size:11px}
table.hm th,table.hm td{border:1px solid #fff;padding:2px 0;width:34px;text-align:center;font-variant-numeric:tabular-nums}
table.hm th{background:var(--soft);font-weight:600;padding:2px 6px;white-space:nowrap}
table.hm td.na{background:#f7f7f7}
.legend{font-size:12px;color:var(--mut);margin:4px 0}
.legend span{display:inline-block;padding:0 6px;margin-right:4px;border-radius:3px}
svg.bar{width:100%;max-width:760px;height:auto;display:block}
svg .ax{font-size:11px;fill:#666}
svg .ref{font-size:11px;fill:#c0392b}
.note{font-size:12.5px;color:var(--mut)}
code{background:var(--soft);padding:0 4px;border-radius:3px}
@media print{main{padding:0}h2{break-after:avoid}table,svg{break-inside:avoid}}
</style></head><body><main>
<h1>インターネット回線 品質計測レポート</h1>
<p class="sub">計測期間: $($From.ToString('yyyy年M月d日 HH:mm')) 〜 $($To.ToString('yyyy年M月d日 HH:mm'))　／　作成: $($stamp.ToString('yyyy年M月d日 HH:mm'))　／　NetMon による自動計測</p>
<h2>1. 結論</h2>
<div class="box"><ul>
"@
foreach ($f in $Findings) { A "<li>$f</li>" }
A '</ul></div>'

# ---- 計測環境と経路 ----
A '<h2>2. 計測環境と経路</h2>'
if ($EnvInfo) {
    A ("<p>計測PC の接続: <b>{0}</b> ({1}, リンク速度 {2} Mbps)　／　判定しきい値: ロス率 {3}%</p>" -f (Esc $EnvInfo.nic_type), (Esc $EnvInfo.nic_name), $EnvInfo.nic_speed_mbps, $Thr)
    A '<table class="t"><tr><th>ID</th><th>区間</th><th>機器・宛先</th><th>IPアドレス</th><th>ホスト名</th><th>計測方式</th></tr>'
    foreach ($t in $EnvInfo.targets) {
        $method = 'ping'
        if ($t.kind -eq 'tcp') { $method = "TCP接続 (port $($t.port))" } elseif ($t.ttl -gt 0) { $method = "TTL=$($t.ttl) の ping (tracert 方式)" }
        $hn = $t.ptr; if (-not $t.is_hop) { $hn = $t.host }
        A ("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td></tr>" -f (Esc $t.id), (Esc $t.layer), (Esc $t.name), (Esc $t.address), (Esc $hn), $method)
    }
    A '</table><p class="note">H1〜 は「このPC → 宅内ルーター → マンション設備 → 回線事業者(ISP)網」と外へ向かう順に並んだ中継機器、E1〜 は外部のサーバです。</p>'
}

# ---- 区間別 ----
A '<h2>3. 区間別の計測結果</h2><table class="t"><tr><th>ID</th><th>区間</th><th>機器・宛先</th><th>IP</th><th>送信</th><th>損失</th><th>ロス率</th><th>平均応答</th><th>最大応答</th><th>ロス率≧しきい値の分数</th></tr>'
$sorted = $TargetAgg.Values | Sort-Object @{ E = { if ($_.Id.StartsWith('H')) { 0 } else { 1 } } }, @{ E = { [int]$_.Id.Substring(1) } }
foreach ($a in $sorted) {
    $p = Pct $a.Lost $a.Sent
    $avg = $NaN; if ($a.RttN) { $avg = $a.RttSum / $a.RttN }
    $bg = Get-LossBg $p
    A ("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td class='n'>{4:N0}</td><td class='n'>{5:N0}</td><td class='n' style='background:{6};color:{7}'>{8}%</td><td class='n'>{9} ms</td><td class='n'>{10} ms</td><td class='n'>{11}</td></tr>" -f (Esc $a.Id), (Esc $a.Layer), (Esc $a.Name), (Esc $a.Address), $a.Sent, $a.Lost, $bg, (Get-Fg $bg), (F $p '0.00'), (F $avg), (F $a.RttMax '0'), $a.BadMin)
}
A '</table><p class="note">中継機器 (H2 以降) は、機器が ping への応答を後回しにする仕様のため、その機器だけロスが多く見えることがあります。<b>ある機器から先のすべての地点で同じようにロスしている場合</b>に、その区間が原因と判断できます (「5. ロス発生箇所の判定」はこの考え方で自動判定しています)。</p>'

# ---- 時間帯別 ----
A '<h2>4. 日付×時間帯の状況</h2>'
A ("<p class='legend'>ロス率(%): <span style='background:#cfeedd'>0</span><span style='background:#eef3b8'>0.5未満</span><span style='background:#ffd978'>{0}未満</span><span style='background:#ffa257'>5未満</span><span style='background:#ee5a45;color:#fff'>15未満</span><span style='background:#a3122a;color:#fff'>15以上</span>　空欄=計測なし　(セルにマウスを乗せると件数を表示)</p>" -f $Thr)
A (New-Heatmap '外部サーバへのパケットロス率 (%)' 'ext')
A (New-Heatmap '宅内ルーターへのパケットロス率 (%) ― 宅内が正常かの確認用' 'gw')
A (New-Heatmap ('外部サーバへの平均応答時間 (ms)  ※平常時 ' + (F $RttBase) + ' ms との比で色分け') 'rtt')
if ($SpeedOk.Count) { A (New-Heatmap ('下り速度 (Mbps)  ※期間中の中央値 ' + (F $DownMed) + ' Mbps との比で色分け') 'down') }

A '<h3>時刻別の平均 (全日合算)</h3><p class="note">外部サーバへのパケットロス率 (%)</p>'
$hv = New-Object double[] 24; for ($h = 0; $h -lt 24; $h++) { $hv[$h] = Pct $HourEL[$h] $HourES[$h] }
A (New-BarSvg $hv '%' '#e06a4f' $Thr ("しきい値 $Thr%"))
$hv = New-Object double[] 24; for ($h = 0; $h -lt 24; $h++) { $hv[$h] = $NaN; if ($HourRttN[$h]) { $hv[$h] = $HourRtt[$h] / $HourRttN[$h] } }
A '<p class="note">外部サーバへの平均応答時間 (ms)</p>'
A (New-BarSvg $hv 'ms' '#4f7fc0' $RttBase '平常時')
if ($SpeedOk.Count) {
    $hv = New-Object double[] 24; for ($h = 0; $h -lt 24; $h++) { $hv[$h] = $NaN; if ($HourDownN[$h]) { $hv[$h] = $HourDown[$h] / $HourDownN[$h] } }
    A '<p class="note">下り速度の平均 (Mbps)</p>'
    A (New-BarSvg $hv 'Mbps' '#3a9a6a' $DownMed '期間中央値')
}

# ---- 判定 ----
A '<h2>5. ロス発生箇所の判定</h2>'
if ($secList.Count) {
    A '<table class="t"><tr><th>ロスの起点 (自動判定)</th><th>該当時間</th><th>割合</th><th>その間の外部宛先ロス</th></tr>'
    foreach ($e in $secList) {
        A ("<tr><td>{0}</td><td class='n'>{1} 分</td><td class='n'>{2}%</td><td class='n'>{3:N0} 発</td></tr>" -f (Esc $e.Key), $e.Value.Min, (F (Pct $e.Value.Min $secTotal) '0'), $e.Value.Lost)
    }
    A '</table>'
} else { A '<p>しきい値を超えるロスは検出されませんでした。</p>' }
A @"
<p class="note">判定方法: 1分ごとに各地点のロス率を比べ、<br>
・宅内ルーターでロス → 「宅内」<br>
・「A 〜 B」 → A まではロスがなく、B から先の全地点 (外部サーバまで) でロスしている。<b>A と B の間の回線、または機器 B が原因</b>の可能性が高い<br>
・「監視中の中継機器より先」 → 確認できた中継機器はすべて正常で、それより先 (ISPの上位網など) でロス<br>
・「○○ のみ」 → 特定の宛先だけでロス (その宛先側の要因)</p>
"@

# ---- ワースト ----
A '<h2>6. ロスがひどかった時間 (ワースト20分)</h2><table class="t"><tr><th>時刻</th><th>外部ロス率</th><th>外部 損失/送信</th><th>宅内ルーター 損失/送信</th><th>平均応答</th><th>自PC受信量</th><th>判定</th></tr>'
foreach ($w in ($Worst | Sort-Object Pct -Descending | Select-Object -First 20)) {
    $bg = Get-LossBg $w.Pct
    A ("<tr><td>{0}</td><td class='n' style='background:{1};color:{2}'>{3}%</td><td class='n'>{4}/{5}</td><td class='n'>{6}/{7}</td><td class='n'>{8} ms</td><td class='n'>{9} Mbps</td><td>{10}</td></tr>" -f $w.Time, $bg, (Get-Fg $bg), (F $w.Pct), $w.EL, $w.ES, $w.GL, $w.GS, (F $w.Rtt), (F $w.Rx '0.00'), (Esc $w.Verdict))
}
A '</table>'

# ---- ロス事象 ----
A ('<h2>7. ロス事象の一覧 (損失の多い順 上位30件 / 全{0}件)</h2>' -f $Incidents.Count)
if ($Incidents.Count) {
    A '<table class="t"><tr><th>開始</th><th>終了</th><th>秒数</th><th>外部 損失/送信</th><th>完全断</th><th>自PC受信量</th><th>判定</th></tr>'
    foreach ($i in ($Incidents | Sort-Object { [int]$_['ext_lost'] } -Descending | Select-Object -First 30)) {
        A ("<tr><td>{0}</td><td>{1}</td><td class='n'>{2}</td><td class='n'>{3}/{4}</td><td class='n'>{5} 秒</td><td class='n'>{6} Mbps</td><td>{7}</td></tr>" -f $i['start'], $i['end'].Substring(11), $i['duration_s'], $i['ext_lost'], $i['ext_sent'], $i['max_full_outage_s'], (F (Num $i['rx_mbps_avg']) '0.00'), (Esc $i['verdict']))
    }
    A '</table><p class="note">「ロス事象」= ロスが5秒以上途切れずに続いた一連の期間。完全断 = 全外部サーバが連続で無応答だった最長秒数。</p>'
} else { A '<p>ロス事象はありません。</p>' }

# ---- 速度 ----
A '<h2>8. 速度測定の結果</h2>'
if ($SpeedOk.Count) {
    A ('<p>Cloudflare の速度測定サーバへ {0} 分ごとに測定 (下り 最大25MB、上り 10MB)。全 {1} 回 (失敗 {2} 回)。</p>' -f $(if ($EnvInfo) { $EnvInfo.speedtest_every_min } else { 30 }), $Speeds.Count, ($Speeds.Count - $SpeedOk.Count))
    A '<h3>遅かった測定 (下り速度の遅い順 15件)</h3><table class="t"><tr><th>時刻</th><th>下り</th><th>上り</th><th>測定直前の自PC受信量</th><th>通常時の応答</th><th>測定中の応答</th></tr>'
    foreach ($o in ($SpeedOk | Sort-Object Down | Select-Object -First 15)) {
        A ("<tr><td>{0}</td><td class='n'>{1} Mbps</td><td class='n'>{2} Mbps</td><td class='n'>{3} Mbps</td><td class='n'>{4} ms</td><td class='n'>{5} ms</td></tr>" -f $o.Time, (F $o.Down), (F $o.Up), (F $o.BgRx '0.00'), (F $o.Idle), (F $o.Loaded))
    }
    A '</table><p class="note">「測定中の応答」が通常時より大きく増える場合、回線の混雑 (帯域不足) を示します。速度測定中の秒はパケットロスの集計から除外しています。</p>'
} else { A '<p>期間内の速度測定はありません。</p>' }

# ---- 方法 ----
$iv = 1000; $tmo = 1000; if ($EnvInfo) { $iv = $EnvInfo.interval_ms; $tmo = $EnvInfo.timeout_ms }
A @"
<h2>9. 計測方法</h2>
<ul class="note">
<li>計測ソフト: NetMon (Windows PowerShell)。Windows 標準の ping 機能 (ICMP Echo) を使用。</li>
<li>$iv ミリ秒ごとに全地点へ同時に送信し、$tmo ミリ秒以内に応答がなければ「損失」として記録。</li>
<li>中継機器は tracert / mtr と同じく TTL を制限した ping で、外部サーバへ向かう実際の経路上の各機器を計測。</li>
<li>同時にこのPCの送受信量を記録し、自分の通信による混雑と区別。速度測定中・LAN切断中・スリープ中は集計から除外。</li>
<li>すべての生データ (1秒ごとの応答時間) を同梱の CSV / zip に保存しています。</li>
</ul>
<h2>10. 同梱データ</h2>
<table class="t">
<tr><td><code>minutes.csv</code></td><td>1分ごとの概要 (外部・宅内のロス数、平均応答、判定、自PC通信量)</td></tr>
<tr><td><code>targets_by_minute.csv</code></td><td>1分ごと・地点ごとの送信数、損失数、応答時間 (最小/平均/95%/最大)、揺らぎ</td></tr>
<tr><td><code>incidents.csv</code></td><td>ロス事象の一覧</td></tr>
<tr><td><code>speedtest.csv</code></td><td>速度測定の結果</td></tr>
<tr><td><code>raw\</code></td><td>1秒ごとの全地点の応答時間 (生データ。LOSS = 応答なし)</td></tr>
</table>
</main></body></html>
"@

$htmlPath = Join-Path $dir 'report.html'
[IO.File]::WriteAllText($htmlPath, $html.ToString(), (New-Object System.Text.UTF8Encoding($false)))

# ------------------------------------------------------------
#  CSV 同梱
# ------------------------------------------------------------
function Export-Filtered([string[]]$Paths, [string]$Out, [int]$TimeIndex) {
    $sb = New-Object System.Text.StringBuilder
    $first = $true
    foreach ($p in $Paths) {
        $lines = (Read-SharedText $p) -split "`r?`n"
        if ($lines.Count -eq 0) { continue }
        if ($first) { [void]$sb.AppendLine($lines[0]); $first = $false }
        for ($i = 1; $i -lt $lines.Count; $i++) {
            if (-not $lines[$i]) { continue }
            $t = $lines[$i].Split(',')[0]
            if ($t.Length -eq 16) { $t += ':00' }   # 1分集計の行は秒がない
            if ([string]::CompareOrdinal($t, $FromS) -lt 0 -or [string]::CompareOrdinal($t, $ToS) -gt 0) { continue }
            [void]$sb.AppendLine($lines[$i])
        }
    }
    [IO.File]::WriteAllText($Out, $sb.ToString(), $Enc)
}
Export-Filtered (Get-DayFiles 'overview_*.csv') (Join-Path $dir 'minutes.csv') 0
[IO.File]::WriteAllText((Join-Path $dir 'targets_by_minute.csv'), (($summaryRows -join "`r`n") + "`r`n"), $Enc)
if (Test-Path -LiteralPath $incPath) { Export-Filtered @($incPath) (Join-Path $dir 'incidents.csv') 0 }
if (Test-Path -LiteralPath $spdPath) { Export-Filtered @($spdPath) (Join-Path $dir 'speedtest.csv') 0 }

# 生データ (期間に重なる日のもの)
$rawOut = Join-Path $dir 'raw'
[void][IO.Directory]::CreateDirectory($rawOut)
$d0 = $From.ToString('yyyyMMdd'); $d1 = $To.ToString('yyyyMMdd')
foreach ($f in Get-ChildItem -LiteralPath (Join-Path $LogDir 'raw') -Filter 'raw_*') {
    $d = $f.BaseName.Split('_')[1]
    if ($d -lt $d0 -or $d -gt $d1) { continue }
    $src = New-Object IO.FileStream($f.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    try {
        $dst = [IO.File]::Create((Join-Path $rawOut $f.Name))
        try { $src.CopyTo($dst) } finally { $dst.Dispose() }
    } finally { $src.Dispose() }
}

# 提出用に zip へまとめる
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zipPath = "$dir.zip"
if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath }
[IO.Compression.ZipFile]::CreateFromDirectory($dir, $zipPath)

Write-Host ''
Write-Host 'レポートを作成しました:' -ForegroundColor Green
Write-Host "  $htmlPath"
Write-Host "  $zipPath  (提出用: レポートと全データ一式)"
if (-not $NoOpen) { Start-Process $htmlPath }
