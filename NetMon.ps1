#Requires -Version 5.1
<#
.SYNOPSIS
  NetMon 回線品質モニタ
  パケットロス・遅延・速度低下を常時記録し、経路上のどこで起きているかを判定します。

.DESCRIPTION
  - 1秒ごとに「宅内ルーター → マンション設備 → ISP網 → 外部サーバ」の各地点へ ping を送信
    (中継機器は tracert/mtr と同じ TTL 制限方式で計測)
  - 結果はすべて logs\ 以下に CSV で保存 (1秒ごとの生データ / 1分集計 / ロス事象 / 速度測定)
  - 30分ごとに Cloudflare で下り/上り速度を測定 (測定中の秒はロス集計から除外)
  - 自PCの通信量も同時記録し「自分の通信が原因ではない」ことを示せるようにする

.EXAMPLE
  .\NetMon.ps1                    # 監視開始
  .\NetMon.ps1 -InstallStartup    # ログオン時に最小化で自動起動するよう登録
  .\NetMon.ps1 -UninstallStartup  # 自動起動を解除
#>
[CmdletBinding()]
param(
    [switch]$InstallStartup,
    [switch]$UninstallStartup,
    [int]$RunSeconds = 0,      # テスト用: 指定秒数で自動終了
    [string]$LogDir = ''
)

# ============================================================
#  設定 (必要に応じて変更してください)
# ============================================================
$Config = @{
    IntervalMs        = 1000       # 計測間隔 (ミリ秒)
    TimeoutMs         = 1000       # 応答待ち上限。超えたら「ロス」として記録
    LossThresholdPct  = 2.0        # この値以上のロス率を「異常」と判定
    TraceTarget       = '1.1.1.1'  # 経路(中継機器)を調べる宛先
    MaxHops           = 15
    MaxMonitoredHops  = 8          # 常時監視する中継機器の数 (手前から)
    # 中継機器の表示名を変えたい場合 (例: ルーターを置いていない部屋で H1 がマンション設備のとき)
    #   HopLabels = @{ 'H1' = 'マンション共用ルーター' }
    HopLabels         = @{}
    Endpoints = @(
        @{ Name = 'Cloudflare';  Host = '1.1.1.1'; Kind = 'icmp'; Layer = '外部' }
        @{ Name = 'Google DNS';  Host = '8.8.8.8'; Kind = 'icmp'; Layer = '外部' }
        @{ Name = 'Azure東日本'; Host = 'japaneast.tts.speech.microsoft.com';    Kind = 'icmp'; Layer = 'ゲーム想定' }
        @{ Name = 'AWS東京';     Host = 'dynamodb.ap-northeast-1.amazonaws.com'; Kind = 'icmp'; Layer = 'ゲーム想定' }
        # ping に応答しないサーバは TCP 接続時間で計測できます:
        # @{ Name = 'Azure東日本TCP'; Host = 'japaneast.tts.speech.microsoft.com'; Kind = 'tcp'; Port = 443; Layer = 'ゲーム想定' }
    )
    SpeedTestEveryMin = 30         # 速度測定の間隔 (分)。0 で無効
    FirstSpeedTestSec = 60         # 起動後、初回の速度測定までの秒数
    SpeedDownUrl      = 'https://speed.cloudflare.com/__down?bytes=25000000'
    SpeedUpUrl        = 'https://speed.cloudflare.com/__up'
    SpeedUpBytes      = 10000000
    SpeedMaxSeconds   = 15         # 下り/上りそれぞれの測定上限秒
    LowTrafficMbps    = 5          # レポートで「自分の通信が少ない」とみなす上限 (Mbps)
    IncidentGapSec    = 5          # ロスが途切れてこの秒数経ったら1件の「ロス事象」として確定
    IncidentMaxSec    = 300        # 長く続くロス事象はこの秒数で区切って記録
    LogDir            = (Join-Path $PSScriptRoot 'logs')
}
if ($LogDir) { $Config.LogDir = $LogDir }
$Version = '1.0'

# ============================================================
#  自動起動の登録/解除
# ============================================================
$ShortcutPath = Join-Path ([Environment]::GetFolderPath('Startup')) 'NetMon 回線品質モニタ.lnk'
if ($InstallStartup) {
    $ws = New-Object -ComObject WScript.Shell
    $sc = $ws.CreateShortcut($ShortcutPath)
    $sc.TargetPath = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $sc.Arguments = '-NoProfile -ExecutionPolicy Bypass -File "' + $PSCommandPath + '"'
    $sc.WorkingDirectory = $PSScriptRoot
    $sc.WindowStyle = 7   # 最小化
    $sc.Description = 'NetMon 回線品質モニタ'
    $sc.Save()
    Write-Host "ログオン時に最小化で自動起動するよう登録しました。" -ForegroundColor Green
    Write-Host "  $ShortcutPath"
    Write-Host "※ NetMon フォルダを移動した場合は、もう一度登録してください。" -ForegroundColor Yellow
    return
}
if ($UninstallStartup) {
    if (Test-Path -LiteralPath $ShortcutPath) {
        Remove-Item -LiteralPath $ShortcutPath
        Write-Host "自動起動を解除しました。" -ForegroundColor Green
    } else {
        Write-Host "自動起動は登録されていません。"
    }
    return
}

# ============================================================
#  計測エンジン (C#)
# ============================================================
$CSharp = @'
using System;
using System.Diagnostics;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;

namespace NetMon
{
    public class ProbeSpec
    {
        public string Kind = "icmp";  // icmp | tcp
        public string Address;
        public int Ttl;               // >0 : TTL limited probe (reply comes from the router at that hop)
        public int Port;
    }

    public class ProbeResult
    {
        public bool Ok;
        public double Rtt = -1;
        public string From = "";
        public string Status = "";
    }

    public static class Prober
    {
        static readonly byte[] Payload = new byte[32];

        // Sends all probes in parallel and waits for every reply (or timeout).
        public static ProbeResult[] Run(ProbeSpec[] specs, int timeoutMs)
        {
            var tasks = new Task<ProbeResult>[specs.Length];
            for (int i = 0; i < specs.Length; i++)
                tasks[i] = specs[i].Kind == "tcp" ? Tcp(specs[i], timeoutMs) : Icmp(specs[i], timeoutMs);
            try { Task.WaitAll(tasks, timeoutMs + 2000); } catch (AggregateException) { }
            var res = new ProbeResult[specs.Length];
            for (int i = 0; i < specs.Length; i++)
            {
                if (tasks[i].Status == TaskStatus.RanToCompletion) res[i] = tasks[i].Result;
                else { res[i] = new ProbeResult(); res[i].Status = tasks[i].IsFaulted ? "Error" : "Hung"; }
            }
            return res;
        }

        static async Task<ProbeResult> Icmp(ProbeSpec s, int timeoutMs)
        {
            var r = new ProbeResult();
            try
            {
                using (var p = new Ping())
                {
                    var opt = new PingOptions(s.Ttl > 0 ? s.Ttl : 128, true);
                    var sw = Stopwatch.StartNew();
                    PingReply rep = await p.SendPingAsync(IPAddress.Parse(s.Address), timeoutMs, Payload, opt).ConfigureAwait(false);
                    sw.Stop();
                    r.Status = rep.Status.ToString();
                    if (rep.Address != null) r.From = rep.Address.ToString();
                    if (rep.Status == IPStatus.Success)
                    {
                        r.Ok = true;
                        // RTT measured by the OS (same value as Windows ping). Below 1ms the OS reports 0.
                        r.Rtt = rep.RoundtripTime > 0 ? rep.RoundtripTime : Math.Min(sw.Elapsed.TotalMilliseconds, 0.9);
                    }
                    else if (rep.Status == IPStatus.TtlExpired)
                    {
                        // The OS does not report RTT for TTL-expired replies, so use our own stopwatch.
                        r.Ok = true;
                        r.Rtt = sw.Elapsed.TotalMilliseconds;
                    }
                }
            }
            catch (Exception e) { r.Status = "Error:" + e.GetBaseException().Message; }
            return r;
        }

        static void Observe(Task t)
        {
            t.ContinueWith(x => { var ignored = x.Exception; }, TaskContinuationOptions.OnlyOnFaulted);
        }

        static async Task<ProbeResult> Tcp(ProbeSpec s, int timeoutMs)
        {
            var r = new ProbeResult();
            r.From = s.Address;
            var c = new TcpClient(AddressFamily.InterNetwork);
            try
            {
                var sw = Stopwatch.StartNew();
                Task ct = c.ConnectAsync(IPAddress.Parse(s.Address), s.Port);
                Task done = await Task.WhenAny(ct, Task.Delay(timeoutMs)).ConfigureAwait(false);
                sw.Stop();
                if (done == ct && !ct.IsFaulted)
                {
                    r.Ok = true; r.Rtt = sw.Elapsed.TotalMilliseconds; r.Status = "Success";
                }
                else if (done == ct)
                {
                    var se = ct.Exception.GetBaseException() as SocketException;
                    if (se != null && se.SocketErrorCode == SocketError.ConnectionRefused)
                    { r.Ok = true; r.Rtt = sw.Elapsed.TotalMilliseconds; r.Status = "Refused"; }
                    else r.Status = "Error";
                }
                else
                {
                    r.Status = "TimedOut";
                    Observe(ct);
                }
            }
            catch (Exception e) { r.Status = "Error:" + e.GetBaseException().Message; }
            finally { c.Close(); }
            return r;
        }
    }

    public static class Util
    {
        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        struct MIB_IPFORWARDROW
        {
            public uint Dest, Mask, Policy, NextHop, IfIndex, Type, Proto, Age, NextHopAS, Metric1, Metric2, Metric3, Metric4, Metric5;
        }

        [System.Runtime.InteropServices.DllImport("iphlpapi.dll")]
        static extern int GetBestRoute(uint dwDestAddr, uint dwSourceAddr, out MIB_IPFORWARDROW pBestRoute);

        // Returns "nexthop|ifIndex" of the route Windows would use for dest, or null.
        public static string BestRoute(string dest)
        {
            MIB_IPFORWARDROW row;
            uint d = BitConverter.ToUInt32(IPAddress.Parse(dest).GetAddressBytes(), 0);
            if (GetBestRoute(d, 0, out row) != 0) return null;
            return new IPAddress((long)row.NextHop).ToString() + "|" + row.IfIndex;
        }

        public static string ResolveV4(string host, int timeoutMs)
        {
            IPAddress ip;
            if (IPAddress.TryParse(host, out ip)) return ip.ToString();
            try
            {
                var t = Dns.GetHostAddressesAsync(host);
                if (!t.Wait(timeoutMs)) return null;
                foreach (var a in t.Result)
                    if (a.AddressFamily == AddressFamily.InterNetwork) return a.ToString();
            }
            catch (Exception) { }
            return null;
        }

        public static string[] Reverse(string[] ips, int timeoutMs)
        {
            var tasks = new Task<IPHostEntry>[ips.Length];
            for (int i = 0; i < ips.Length; i++) tasks[i] = Dns.GetHostEntryAsync(ips[i]);
            try { Task.WaitAll(tasks, timeoutMs); } catch (AggregateException) { }
            var res = new string[ips.Length];
            for (int i = 0; i < ips.Length; i++)
            {
                res[i] = "";
                if (tasks[i].Status == TaskStatus.RanToCompletion && tasks[i].Result.HostName != ips[i])
                    res[i] = tasks[i].Result.HostName;
            }
            return res;
        }

        static int CharWidth(char c)
        {
            if (c < 0x1100) return 1;
            if (c <= 0x115F) return 2;
            if (c >= 0x2E80 && c <= 0xA4CF) return 2;
            if (c >= 0xAC00 && c <= 0xD7A3) return 2;
            if (c >= 0xF900 && c <= 0xFAFF) return 2;
            if (c >= 0xFE30 && c <= 0xFE4F) return 2;
            if (c >= 0xFF00 && c <= 0xFF60) return 2;
            if (c >= 0xFFE0 && c <= 0xFFE6) return 2;
            return 1;
        }

        // Display width in console cells (full-width Japanese = 2).
        public static int Width(string s)
        {
            int w = 0;
            if (s != null) foreach (char c in s) w += CharWidth(c);
            return w;
        }

        // Truncate/pad to exactly `width` cells.
        public static string Fit(string s, int width)
        {
            if (s == null) s = "";
            var sb = new StringBuilder();
            int w = 0;
            foreach (char c in s)
            {
                int cw = CharWidth(c);
                if (w + cw > width) break;
                sb.Append(c); w += cw;
            }
            while (w < width) { sb.Append(' '); w++; }
            return sb.ToString();
        }

        public static string FitRight(string s, int width)
        {
            if (s == null) s = "";
            int w = Width(s);
            if (w >= width) return Fit(s, width);
            return new string(' ', width - w) + s;
        }
    }

    // Download/upload throughput test that runs in the background so pinging continues.
    public class SpeedTest
    {
        volatile bool running;
        public bool Running { get { return running; } }
        public double DownMbps = -1, UpMbps = -1, DownMB, UpMB;
        public string Error = "";
        public DateTime StartedAt, FinishedAt;

        public void Start(string downUrl, string upUrl, long upBytes, int maxSeconds)
        {
            if (running) return;
            running = true;
            DownMbps = -1; UpMbps = -1; DownMB = 0; UpMB = 0; Error = "";
            StartedAt = DateTime.Now;
            Task.Run(() =>
            {
                try { DownMbps = Download(downUrl, maxSeconds); }
                catch (Exception e) { Error += "down: " + e.GetBaseException().Message + " "; }
                try { if (upBytes > 0) UpMbps = Upload(upUrl, upBytes, maxSeconds); }
                catch (Exception e) { Error += "up: " + e.GetBaseException().Message; }
                FinishedAt = DateTime.Now;
                running = false;
            });
        }

        double Download(string url, int maxSeconds)
        {
            var req = (HttpWebRequest)WebRequest.Create(url);
            req.UserAgent = "NetMon/1.0";
            req.Timeout = 15000;
            req.ReadWriteTimeout = 15000;
            using (var resp = req.GetResponse())
            using (var st = resp.GetResponseStream())
            {
                var buf = new byte[65536];
                // Exclude time to first byte (TCP/TLS setup) from the throughput.
                int n = st.Read(buf, 0, buf.Length);
                if (n <= 0) throw new Exception("empty response");
                long total = 0;
                var sw = Stopwatch.StartNew();
                bool cut = false;
                while ((n = st.Read(buf, 0, buf.Length)) > 0)
                {
                    total += n;
                    if (sw.Elapsed.TotalSeconds >= maxSeconds) { cut = true; break; }
                }
                sw.Stop();
                DownMB = total / 1e6;
                double mbps = total * 8.0 / sw.Elapsed.TotalSeconds / 1e6;
                if (cut) { try { req.Abort(); } catch (Exception) { } }
                if (total < 100000) throw new Exception("too little data");
                return mbps;
            }
        }

        double Upload(string url, long bytes, int maxSeconds)
        {
            var req = (HttpWebRequest)WebRequest.Create(url);
            req.Method = "POST";
            req.ContentType = "application/octet-stream";
            req.ContentLength = bytes;
            req.AllowWriteStreamBuffering = false;
            req.UserAgent = "NetMon/1.0";
            req.Timeout = 30000;
            req.ReadWriteTimeout = 30000;
            req.ServicePoint.Expect100Continue = false;
            var buf = new byte[65536];
            new Random().NextBytes(buf);
            long sent = 0;
            bool cut = false;
            var rs = req.GetRequestStream();
            var sw = Stopwatch.StartNew();
            while (sent < bytes)
            {
                int n = (int)Math.Min(buf.Length, bytes - sent);
                rs.Write(buf, 0, n);
                sent += n;
                if (sw.Elapsed.TotalSeconds >= maxSeconds) { cut = true; break; }
            }
            double sec;
            if (cut)
            {
                sec = sw.Elapsed.TotalSeconds;
                try { req.Abort(); } catch (Exception) { }
            }
            else
            {
                rs.Close();
                using (req.GetResponse()) { }
                sec = sw.Elapsed.TotalSeconds;
            }
            UpMB = sent / 1e6;
            return sent * 8.0 / sec / 1e6;
        }
    }
}
'@

# ============================================================
#  初期化
# ============================================================
$ErrorActionPreference = 'Stop'
if (-not ('NetMon.Prober' -as [type])) { Add-Type -TypeDefinition $CSharp -Language CSharp }
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

$Inv = [Globalization.CultureInfo]::InvariantCulture
$Enc = New-Object System.Text.UTF8Encoding($true)   # Excel で文字化けしないよう BOM 付き UTF-8
$CsvSpecial = [char[]]@(',', '"', "`r", "`n")
$Dirs = @{
    Root   = $Config.LogDir
    Raw    = Join-Path $Config.LogDir 'raw'
    Minute = Join-Path $Config.LogDir 'minute'
}
foreach ($d in $Dirs.Values) { [void][IO.Directory]::CreateDirectory($d) }

# 二重起動防止
$createdNew = $false
$Mutex = New-Object System.Threading.Mutex($true, 'Local\NetMonLineQualityMonitor', [ref]$createdNew)
if (-not $createdNew) {
    Write-Host 'NetMon はすでに起動しています (タスクバーの NetMon ウィンドウを確認してください)。' -ForegroundColor Yellow
    Start-Sleep -Seconds 5
    return
}

# ============================================================
#  ログ出力
# ============================================================
function Format-CsvField($v) {
    if ($null -eq $v) { return '' }
    if ($v -is [double]) {
        if ([double]::IsNaN($v)) { return '' }
        return $v.ToString('0.###', $Inv)
    }
    $s = [string]$v
    if ($s.IndexOfAny($CsvSpecial) -ge 0) { return '"' + $s.Replace('"', '""') + '"' }
    return $s
}

function Join-Csv([object[]]$Fields) {
    $sb = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $Fields.Count; $i++) {
        if ($i) { [void]$sb.Append(',') }
        [void]$sb.Append((Format-CsvField $Fields[$i]))
    }
    $sb.ToString()
}

function Open-Log([string]$Path, [string]$Header) {
    $isNew = (-not (Test-Path -LiteralPath $Path)) -or ((Get-Item -LiteralPath $Path).Length -eq 0)
    # 監視中でも Excel やレポート作成で読めるよう共有モードで開く
    $fs = New-Object System.IO.FileStream($Path, [IO.FileMode]::Append, [IO.FileAccess]::Write, [IO.FileShare]::ReadWrite)
    $w = New-Object System.IO.StreamWriter($fs, $Enc)
    $w.AutoFlush = $true
    if ($isNew) { $w.WriteLine($Header) }
    return $w
}

$EventLog    = Open-Log (Join-Path $Dirs.Root 'events.csv') 'time,type,message'
$IncidentLog = Open-Log (Join-Path $Dirs.Root 'incidents.csv') 'start,end,duration_s,ext_lost,ext_sent,ext_loss_pct,max_full_outage_s,verdict_code,verdict_section,verdict,rx_mbps_avg,tx_mbps_avg,detail'
$SpeedLog    = Open-Log (Join-Path $Dirs.Root 'speedtest.csv') 'time,down_mbps,up_mbps,down_mb,up_mb,bg_rx_mbps,bg_tx_mbps,idle_rtt_ms,loaded_rtt_ms,error'
$RouteLog    = Open-Log (Join-Path $Dirs.Root 'routes.csv') 'time,id,ttl,address,ptr,layer,name'

$Events = New-Object System.Collections.Generic.List[object]
function Write-Event([string]$Type, [string]$Message, [ConsoleColor]$Color = [ConsoleColor]::Gray) {
    $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    try { $script:EventLog.WriteLine((Join-Csv @($ts, $Type, $Message))) } catch { }
    $script:Events.Add([pscustomobject]@{ Time = $ts.Substring(5); Type = $Type; Message = $Message; Color = $Color })
    while ($script:Events.Count -gt 40) { $script:Events.RemoveAt(0) }
}

function Open-DayLogs {
    $day = (Get-Date).ToString('yyyyMMdd')
    foreach ($w in @($script:SummaryLog, $script:OverviewLog)) { if ($w) { $w.Dispose() } }
    $script:SummaryLog = Open-Log (Join-Path $Dirs.Minute "summary_$day.csv") 'time,id,layer,name,address,kind,sent,lost,loss_pct,rtt_min,rtt_avg,rtt_p95,rtt_max,jitter,other_responder'
    $script:OverviewLog = Open-Log (Join-Path $Dirs.Minute "overview_$day.csv") 'time,verdict_code,verdict_section,verdict,gw_sent,gw_lost,dest_sent,dest_lost,ext_sent,ext_lost,ext_loss_pct,ext_rtt_avg,ext_rtt_max,rtt_high,rx_mbps_avg,tx_mbps_avg,rx_mbps_max,measured_s,excluded_s,linkdown_s,wifi_signal'
    $script:CurDay = (Get-Date).Date
}

function Open-RawLog {
    if ($script:RawLog) { $script:RawLog.Dispose() }
    $script:RawPath = Join-Path $Dirs.Raw ('raw_' + (Get-Date).ToString('yyyyMMdd_HHmmss') + '.csv')
    $cols = @('time', 'link_up', 'speedtest', 'rx_mbps', 'tx_mbps')
    foreach ($t in $script:Targets) { $cols += ('{0} {1}' -f $t.Id, $t.Label) }
    $script:RawLog = Open-Log $script:RawPath (Join-Csv $cols)
    Compress-OldRaw
}

# 過去の生データは zip に圧縮 (1日あたり数MB → 1MB未満)
function Compress-OldRaw {
    foreach ($f in Get-ChildItem -LiteralPath $Dirs.Raw -Filter 'raw_*.csv') {
        if ($f.FullName -eq $script:RawPath) { continue }
        $zip = [IO.Path]::ChangeExtension($f.FullName, '.zip')
        try {
            if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip }
            $za = [IO.Compression.ZipFile]::Open($zip, [IO.Compression.ZipArchiveMode]::Create)
            try { [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($za, $f.FullName, $f.Name, [IO.Compression.CompressionLevel]::Optimal) }
            finally { $za.Dispose() }
            Remove-Item -LiteralPath $f.FullName
        } catch { }
    }
}

# ============================================================
#  経路の検出と監視対象の構築
# ============================================================
function Test-PrivateIp([string]$Ip) {
    $b = [Net.IPAddress]::Parse($Ip).GetAddressBytes()
    return ($b[0] -eq 10) -or ($b[0] -eq 172 -and $b[1] -ge 16 -and $b[1] -le 31) -or
           ($b[0] -eq 192 -and $b[1] -eq 168) -or ($b[0] -eq 100 -and $b[1] -ge 64 -and $b[1] -le 127)
}

function New-Target($Id, $Layer, $Name, $Address, $Ptr, $IsHop, $Kind, $Ttl, $Port, $HostName) {
    $spec = New-Object NetMon.ProbeSpec
    $spec.Kind = $Kind; $spec.Address = $Address; $spec.Ttl = $Ttl; $spec.Port = $Port
    $ring = New-Object double[] $script:RingLen
    for ($i = 0; $i -lt $ring.Length; $i++) { $ring[$i] = [double]::NaN }
    [pscustomobject]@{
        Id = $Id; Layer = $Layer; Name = $Name; Address = $Address; Ptr = $Ptr; Host = $HostName
        Label = '{0}({1})' -f $Name, $Address
        IsHop = $IsHop; Kind = $Kind; Ttl = $Ttl; Port = $Port; Spec = $spec
        Ring = $ring
        MSent = 0; MLost = 0; MOther = 0; MRtts = (New-Object System.Collections.Generic.List[double])
        TSent = 0; TLost = 0
    }
}

function Initialize-Targets {
    # 既定ルート (どのNIC・どのルーターから外へ出るか)
    $destIp = [NetMon.Util]::ResolveV4($Config.TraceTarget, 3000)
    if (-not $destIp) { return $false }
    $rt = [NetMon.Util]::BestRoute($destIp)
    if (-not $rt) { return $false }
    $gwIp, $ifIndex = $rt.Split('|')
    if ($gwIp -eq '0.0.0.0') { return $false }
    $nic = $null
    foreach ($n in [Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces()) {
        try { if ($n.GetIPProperties().GetIPv4Properties().Index -eq [int]$ifIndex) { $nic = $n; break } } catch { }
    }
    if (-not $nic) { return $false }

    # TTL を 1..MaxHops に変えた ping を3回ずつ並列送信して中継機器を調べる (tracert と同じ原理)
    $specs = foreach ($ttl in 1..$Config.MaxHops) {
        $s = New-Object NetMon.ProbeSpec; $s.Address = $destIp; $s.Ttl = $ttl; $s
    }
    $resp = @{}
    $reach = 0
    for ($round = 0; $round -lt 3; $round++) {
        $res = [NetMon.Prober]::Run([NetMon.ProbeSpec[]]$specs, $Config.TimeoutMs)
        for ($i = 0; $i -lt $res.Length; $i++) {
            $r = $res[$i]; $ttl = $i + 1
            if (-not $r.Ok -or -not $r.From) { continue }
            if ($r.Status -eq 'Success' -and ($reach -eq 0 -or $ttl -lt $reach)) { $reach = $ttl }
            if (-not $resp.ContainsKey($ttl)) { $resp[$ttl] = @{} }
            $resp[$ttl][$r.From] = 1 + [int]$resp[$ttl][$r.From]
        }
    }
    if ($reach -eq 0) { $reach = $Config.MaxHops + 1 }

    $found = @{}
    foreach ($ttl in $resp.Keys) {
        if ($ttl -lt 2 -or $ttl -ge $reach) { continue }
        $found[$ttl] = ($resp[$ttl].GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Key
    }
    # 応答を間引く中継機器はたまたま無応答になることがあるため、
    # 同じTTLで別の機器が応答した場合だけ「経路変更」とみなし、それ以外は前回の結果と統合する
    $changed = $false
    foreach ($ttl in $found.Keys) {
        if ($script:KnownHops.ContainsKey($ttl) -and $script:KnownHops[$ttl] -ne $found[$ttl]) { $changed = $true }
    }
    if (-not $changed) {
        foreach ($ttl in $script:KnownHops.Keys) {
            if ($ttl -lt $reach -and -not $found.ContainsKey($ttl)) { $found[$ttl] = $script:KnownHops[$ttl] }
        }
    }
    $hops = @()
    foreach ($ttl in ($found.Keys | Sort-Object)) {
        if ($hops.Count -ge ($Config.MaxMonitoredHops - 1)) { break }
        if ($found[$ttl] -eq $gwIp) { continue }
        $hops += [pscustomobject]@{ Ttl = $ttl; Ip = $found[$ttl] }
    }
    $ptrs = [NetMon.Util]::Reverse([string[]](@($gwIp) + @($hops | ForEach-Object { $_.Ip })), 3000)

    $targets = New-Object System.Collections.Generic.List[object]
    # H1 = 宅内ルーター (デフォルトゲートウェイ) は通常の ping で計測
    $targets.Add((New-Target 'H1' '宅内' '宅内ルーター' $gwIp $ptrs[0] $true 'icmp' 0 0 $gwIp))
    $seenPublic = $false
    for ($k = 0; $k -lt $hops.Count; $k++) {
        $h = $hops[$k]
        if (-not (Test-PrivateIp $h.Ip)) { $seenPublic = $true }
        if ($seenPublic) { $layer = 'ISP/上位網'; $name = 'ISP網' } else { $layer = 'マンション'; $name = 'マンション設備' }
        $targets.Add((New-Target ('H' + $h.Ttl) $layer $name $h.Ip $ptrs[$k + 1] $true 'icmp' $h.Ttl 0 $destIp))
    }
    foreach ($t in $targets) {
        if ($Config.HopLabels.ContainsKey($t.Id)) { $t.Name = $Config.HopLabels[$t.Id]; $t.Label = '{0}({1})' -f $t.Name, $t.Address }
    }

    $destIndex = -1
    $e = 0
    foreach ($ep in $Config.Endpoints) {
        $e++
        # DNS が毎回違うIPを返すサーバもあるため、一度決めた宛先IPは起動中は固定する
        $ip = $script:EndpointIps[$ep.Host]
        if (-not $ip) { $ip = [NetMon.Util]::ResolveV4($ep.Host, 3000); if ($ip) { $script:EndpointIps[$ep.Host] = $ip } }
        if (-not $ip) { Write-Event '警告' ("名前解決できないため監視対象から外しました: {0} ({1})" -f $ep.Name, $ep.Host) Yellow; continue }
        $port = 0; if ($ep.ContainsKey('Port')) { $port = [int]$ep.Port }
        $targets.Add((New-Target ('E' + $e) $ep.Layer $ep.Name $ip '' $false $ep.Kind 0 $port $ep.Host))
        if ($destIndex -lt 0 -and $ip -eq $destIp -and $ep.Kind -eq 'icmp') { $destIndex = $targets.Count - 1 }
    }
    if ($destIndex -lt 0) {
        for ($i = 0; $i -lt $targets.Count; $i++) { if (-not $targets[$i].IsHop) { $destIndex = $i; break } }
    }

    $script:KnownHops = @{}
    foreach ($h in $hops) { $script:KnownHops[$h.Ttl] = $h.Ip }
    $sig = ($targets | ForEach-Object { $_.Id + '=' + $_.Address }) -join ' '
    $script:Nic = $nic
    $script:NicId = $nic.Id
    $script:IsWifi = ($nic.NetworkInterfaceType -eq [Net.NetworkInformation.NetworkInterfaceType]::Wireless80211)
    $script:NeedRetrace = $false
    if ($script:Ready -and $sig -eq $script:RouteSig) {
        return $true   # 経路は変わっていない → 監視対象と累計値はそのまま
    }

    # ---- 切替 ----
    if ($script:Ready) { Close-Incident; Save-Minute }
    $script:Targets = $targets.ToArray()
    $script:Specs = [NetMon.ProbeSpec[]]@($script:Targets | ForEach-Object { $_.Spec })
    $script:GwIndex = 0
    $script:DestIndex = $destIndex
    $script:Nic = $nic
    $script:NicId = $nic.Id
    $script:IsWifi = ($nic.NetworkInterfaceType -eq [Net.NetworkInformation.NetworkInterfaceType]::Wireless80211)
    $script:LastRx = -1
    for ($i = 0; $i -lt $script:RingLen; $i++) { $script:StateRing[$i] = -1 }
    $script:NeedRetrace = $false
    $script:Ready = $true

    $sig = ($script:Targets | ForEach-Object { $_.Id + '=' + $_.Address }) -join ' '
    $path = ($script:Targets | Where-Object { $_.IsHop } | ForEach-Object { '{0} {1}' -f $_.Id, $_.Address }) -join ' > '
    if ($sig -ne $script:RouteSig) {
        if ($script:RouteSig) { Write-Event '経路変更' ("経路が変わりました: $path") Yellow }
        else { Write-Event '経路' ("経路を検出: $path") Cyan }
        $now = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        foreach ($t in $script:Targets) {
            $script:RouteLog.WriteLine((Join-Csv @($now, $t.Id, $t.Ttl, $t.Address, $t.Ptr, $t.Layer, $t.Name)))
        }
        $script:RouteSig = $sig
    }
    Save-Environment
    Open-RawLog
    return $true
}

# レポート作成用に計測環境を保存
function Save-Environment {
    $nic = $script:Nic
    $localIp = ''
    try { $localIp = ($nic.GetIPProperties().UnicastAddresses | Where-Object { $_.Address.AddressFamily -eq 'InterNetwork' } | Select-Object -First 1).Address.ToString() } catch { }
    $envInfo = [ordered]@{
        version     = $Version
        updated     = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        nic_name    = $nic.Name
        nic_desc    = $nic.Description
        nic_type    = $(if ($script:IsWifi) { 'Wi-Fi' } else { [string]$nic.NetworkInterfaceType })
        nic_speed_mbps = [math]::Round($nic.Speed / 1e6)
        local_ip    = $localIp
        trace_target = $Config.TraceTarget
        interval_ms = $Config.IntervalMs
        timeout_ms  = $Config.TimeoutMs
        loss_threshold_pct = $Config.LossThresholdPct
        low_traffic_mbps   = $Config.LowTrafficMbps
        speedtest_every_min = $Config.SpeedTestEveryMin
        speedtest_down_url  = $Config.SpeedDownUrl
        speedtest_up_url    = $Config.SpeedUpUrl
        targets     = @($script:Targets | ForEach-Object {
            [ordered]@{ id = $_.Id; layer = $_.Layer; name = $_.Name; address = $_.Address; ptr = $_.Ptr; host = $_.Host; kind = $_.Kind; ttl = $_.Ttl; port = $_.Port; is_hop = $_.IsHop }
        })
    }
    [IO.File]::WriteAllText((Join-Path $Dirs.Root 'environment.json'), ($envInfo | ConvertTo-Json -Depth 5), $Enc)
}

# ============================================================
#  判定ロジック
#   ・宅内ルーターで落ちていれば → 宅内の問題
#   ・宛先で落ちている場合、ロスが「最初に現れ、かつ宛先まで続いている」中継機器を探す
#     (途中の機器だけ落ちて先が正常なのは、機器が ping 応答を間引いているだけで実害なし)
# ============================================================
function New-Verdict($Code, $Section, $Text, $Level) {
    [pscustomobject]@{ Code = $Code; Section = $Section; Text = $Text; Level = $Level }
}

function Get-Verdict([int[]]$Sent, [int[]]$Lost) {
    $TL = $script:Targets
    $thr = $Config.LossThresholdPct
    $n = $TL.Count
    $pct = New-Object double[] $n
    $endSent = 0
    for ($i = 0; $i -lt $n; $i++) {
        if ($Sent[$i] -gt 0) { $pct[$i] = 100.0 * $Lost[$i] / $Sent[$i] } else { $pct[$i] = -1 }
        if (-not $TL[$i].IsHop) { $endSent += $Sent[$i] }
    }
    if ($endSent -eq 0) { return New-Verdict 'NODATA' 'データなし' 'データなし' 0 }
    # 1発だけのロスは偶発的なものとして扱い、異常判定には2発以上のロスを必要とする
    $pctS = New-Object double[] $n
    for ($i = 0; $i -lt $n; $i++) { if ($Lost[$i] -ge 2) { $pctS[$i] = $pct[$i] } else { $pctS[$i] = [Math]::Min($pct[$i], 0) } }

    $g = $script:GwIndex
    if ($pctS[$g] -ge $thr) {
        return New-Verdict 'HOME' ('宅内 (このPC〜{0})' -f $TL[$g].Label) ('宅内でロス {0:0.0}%: このPC〜{1} の間 (LANケーブル/ルーター/PCを確認)' -f $pct[$g], $TL[$g].Label) 3
    }
    $d = $script:DestIndex
    if ($d -ge 0 -and $pctS[$d] -ge $thr) {
        $need = [Math]::Max($thr, $pct[$d] * 0.5)
        $prev = $g
        for ($h = 0; $h -lt $n; $h++) {
            if (-not $TL[$h].IsHop -or $h -eq $g -or $pct[$h] -lt 0) { continue }
            if ($pct[$h] -ge $need) {
                $sec = '{0} 〜 {1}' -f $TL[$prev].Label, $TL[$h].Label
                return New-Verdict ('HOP:' + $TL[$h].Id) $sec ('{0} の区間からロス発生 → 宛先まで波及 ({1} {2:0.0}%)' -f $sec, $TL[$d].Name, $pct[$d]) 3
            }
            $prev = $h
        }
        return New-Verdict 'BEYOND' '監視中の中継機器より先' ('中継機器より先でロス ({0} {1:0.0}%)' -f $TL[$d].Name, $pct[$d]) 3
    }
    $bad = @()
    for ($i = 0; $i -lt $n; $i++) { if (-not $TL[$i].IsHop -and $pctS[$i] -ge $thr) { $bad += $i } }
    if ($bad.Count -ge 2) {
        $names = ($bad | ForEach-Object { '{0} {1:0.0}%' -f $TL[$_].Name, $pct[$_] }) -join ', '
        return New-Verdict 'BEYOND' '監視中の中継機器より先' ('複数の外部宛先でロス: ' + $names) 3
    }
    if ($bad.Count -eq 1) {
        $b = $bad[0]
        return New-Verdict ('SINGLE:' + $TL[$b].Id) ('{0} のみ' -f $TL[$b].Name) ('{0} のみロス {1:0.0}% (宛先側の要因の可能性)' -f $TL[$b].Name, $pct[$b]) 2
    }
    for ($i = 0; $i -lt $n; $i++) {
        if ($TL[$i].IsHop -and $pctS[$i] -ge $thr) {
            return New-Verdict 'ICMPLIMIT' '中継機器の応答間引き(実害なし)' '中継機器が ping 応答を間引いているだけ (通信への実害なし)' 1
        }
    }
    return New-Verdict 'OK' '正常' '正常' 0
}

# ============================================================
#  1分集計
# ============================================================
function Get-RttStats($List) {
    $c = $List.Count
    if ($c -eq 0) { return @([double]::NaN, [double]::NaN, [double]::NaN, [double]::NaN, [double]::NaN) }
    $a = $List.ToArray()
    $sum = 0.0; $jit = 0.0
    for ($i = 0; $i -lt $c; $i++) {
        $sum += $a[$i]
        if ($i) { $jit += [Math]::Abs($a[$i] - $a[$i - 1]) }
    }
    $jitter = [double]::NaN; if ($c -gt 1) { $jitter = $jit / ($c - 1) }
    [Array]::Sort($a)
    $p95 = $a[[Math]::Min($c - 1, [int][Math]::Ceiling($c * 0.95) - 1)]
    return @($a[0], ($sum / $c), $p95, $a[$c - 1], $jitter)
}

function Reset-Minute {
    $script:M = @{ Key = (Get-Date).ToString('yyyyMMddHHmm'); Start = (Get-Date); Secs = 0; Measured = 0; Excl = 0; Link = 0; RxSum = 0.0; TxSum = 0.0; RxMax = 0.0 }
    foreach ($t in $script:Targets) { $t.MSent = 0; $t.MLost = 0; $t.MOther = 0; $t.MRtts.Clear() }
}

function Save-Minute {
    $m = $script:M
    if (-not $m -or $m.Secs -eq 0 -or -not $script:Targets) { Reset-Minute; return }
    $ts = $m.Start.ToString('yyyy-MM-dd HH:mm')
    $TL = $script:Targets
    $sent = New-Object int[] $TL.Count
    $lost = New-Object int[] $TL.Count
    $extSent = 0; $extLost = 0
    $extR = New-Object System.Collections.Generic.List[double]
    for ($i = 0; $i -lt $TL.Count; $i++) {
        $t = $TL[$i]
        $sent[$i] = $t.MSent; $lost[$i] = $t.MLost
        $st = Get-RttStats $t.MRtts
        $lp = [double]::NaN; if ($t.MSent -gt 0) { $lp = 100.0 * $t.MLost / $t.MSent }
        $script:SummaryLog.WriteLine((Join-Csv @($ts, $t.Id, $t.Layer, $t.Name, $t.Address, $t.Kind, $t.MSent, $t.MLost, $lp, $st[0], $st[1], $st[2], $st[3], $st[4], $t.MOther)))
        if (-not $t.IsHop) { $extSent += $t.MSent; $extLost += $t.MLost; $extR.AddRange($t.MRtts) }
        # 中継機器の応答元が過半数で変わっていたら経路が変わった可能性 → 再探索
        if ($t.IsHop -and $t.Ttl -gt 0 -and $t.MSent -ge 10 -and $t.MOther -gt ($t.MSent / 2)) { $script:NeedRetrace = $true }
    }
    $v = Get-Verdict $sent $lost
    $es = Get-RttStats $extR
    $extPct = [double]::NaN; if ($extSent -gt 0) { $extPct = 100.0 * $extLost / $extSent }

    # 平常時RTT (直近24時間のロスなし分の下位10%) と比べて大きく遅いか
    $rttHigh = 0
    if (-not [double]::IsNaN($es[1])) {
        if ($script:RttBase.Count -ge 10) {
            $b = $script:RttBase.ToArray(); [Array]::Sort($b)
            $base = $b[[int][Math]::Floor($b.Length * 0.1)]
            $script:Baseline = $base
            if ($es[1] -gt [Math]::Max($base * 2, $base + 20)) { $rttHigh = 1 }
        }
        if ($extPct -lt $Config.LossThresholdPct) {
            $script:RttBase.Add($es[1])
            while ($script:RttBase.Count -gt 1440) { $script:RttBase.RemoveAt(0) }
        }
    }
    $gw = $script:GwIndex; $d = $script:DestIndex
    $dSent = 0; $dLost = 0; if ($d -ge 0) { $dSent = $sent[$d]; $dLost = $lost[$d] }
    $wifi = ''
    if ($script:IsWifi) { $wifi = Get-WifiSignal }
    $script:OverviewLog.WriteLine((Join-Csv @(
        $ts, $v.Code, $v.Section, $v.Text, $sent[$gw], $lost[$gw], $dSent, $dLost, $extSent, $extLost, $extPct,
        $es[1], $es[3], $rttHigh, ($m.RxSum / $m.Secs), ($m.TxSum / $m.Secs), $m.RxMax, $m.Measured, $m.Excl, $m.Link, $wifi)))
    $script:LastMinuteVerdict = $v
    Reset-Minute
}

function Get-WifiSignal {
    try {
        foreach ($line in (netsh wlan show interfaces)) {
            if ($line -match '^\s*(Signal|シグナル)\s*:\s*(\d+)\s*%') { return [int]$Matches[2] }
        }
    } catch { }
    return ''
}

# ============================================================
#  ロス事象 (連続したロスをひとまとめにした記録)
# ============================================================
function Update-Incident($Now, $Res, [bool]$LossNow, [bool]$AllLost, [double]$Rx, [double]$Tx) {
    $TL = $script:Targets
    if (-not $script:Inc) {
        if (-not $LossNow) { return }
        $script:Inc = @{ Start = $Now; Last = $Now; Sent = (New-Object int[] $TL.Count); Lost = (New-Object int[] $TL.Count); Run = 0; MaxRun = 0; Rx = 0.0; Tx = 0.0; Secs = 0 }
        $names = @()
        for ($i = 0; $i -lt $TL.Count; $i++) { if (-not $Res[$i].Ok -and (-not $TL[$i].IsHop -or $i -eq $script:GwIndex)) { $names += $TL[$i].Name } }
        Write-Event 'ロス' ('パケットロス発生: ' + ($names -join ', ')) Red
    }
    $inc = $script:Inc
    for ($i = 0; $i -lt $TL.Count; $i++) {
        $inc.Sent[$i]++
        if (-not $Res[$i].Ok) { $inc.Lost[$i]++ }
    }
    $inc.Secs++; $inc.Rx += $Rx; $inc.Tx += $Tx
    if ($AllLost) { $inc.Run++; if ($inc.Run -gt $inc.MaxRun) { $inc.MaxRun = $inc.Run } } else { $inc.Run = 0 }
    if ($LossNow) { $inc.Last = $Now }
    if (($Now - $inc.Last).TotalSeconds -ge $Config.IncidentGapSec -or ($Now - $inc.Start).TotalSeconds -ge $Config.IncidentMaxSec) {
        Close-Incident
    }
}

function Close-Incident {
    $inc = $script:Inc
    if (-not $inc) { return }
    $script:Inc = $null
    $TL = $script:Targets
    $v = Get-Verdict $inc.Sent $inc.Lost
    if ($v.Code -eq 'OK' -or $v.Code -eq 'ICMPLIMIT') { $v = New-Verdict 'SPORADIC' '単発のロス' '単発のロス (件数が少なく発生箇所は判定不可)' 1 }
    $es = 0; $el = 0
    $detail = @()
    for ($i = 0; $i -lt $TL.Count; $i++) {
        if (-not $TL[$i].IsHop) { $es += $inc.Sent[$i]; $el += $inc.Lost[$i] }
        $detail += '{0}:{1}/{2}' -f $TL[$i].Id, $inc.Lost[$i], $inc.Sent[$i]
    }
    $dur = [int]($inc.Last - $inc.Start).TotalSeconds + 1
    $pct = [double]::NaN; if ($es -gt 0) { $pct = 100.0 * $el / $es }
    $script:IncidentLog.WriteLine((Join-Csv @(
        $inc.Start.ToString('yyyy-MM-dd HH:mm:ss'), $inc.Last.ToString('yyyy-MM-dd HH:mm:ss'), $dur, $el, $es, $pct, $inc.MaxRun,
        $v.Code, $v.Section, $v.Text, ($inc.Rx / [Math]::Max(1, $inc.Secs)), ($inc.Tx / [Math]::Max(1, $inc.Secs)), ($detail -join ' '))))
    $script:IncToday++
    $msg = '{0}-{1} ({2}秒) 外部宛先 {3}/{4}発 損失' -f $inc.Start.ToString('HH:mm:ss'), $inc.Last.ToString('HH:mm:ss'), $dur, $el, $es
    if ($inc.MaxRun -ge 3) { $msg += ', 完全断 {0}秒' -f $inc.MaxRun }
    $msg += ' | ' + $v.Text
    $color = [ConsoleColor]::Yellow; if ($v.Level -ge 3) { $color = [ConsoleColor]::Red }
    Write-Event 'ロス事象' $msg $color
}

# ============================================================
#  速度測定
# ============================================================
function Get-NextSpeedTime([datetime]$From) {
    $every = $Config.SpeedTestEveryMin
    if ($every -le 0) { return [datetime]::MaxValue }
    $mins = [Math]::Floor(($From - $From.Date).TotalMinutes / $every + 1) * $every
    return $From.Date.AddMinutes($mins)
}

function Get-RingAvg([double[]]$Ring, [int]$Seconds) {
    $sum = 0.0; $c = 0
    for ($k = 1; $k -le $Seconds; $k++) {
        $v = $Ring[($script:RingPos - $k + $script:RingLen) % $script:RingLen]
        if (-not [double]::IsNaN($v) -and $v -ge 0) { $sum += $v; $c++ }
    }
    if ($c -eq 0) { return [double]::NaN }
    return $sum / $c
}

function Start-SpeedTest {
    if ($script:Speed.Running -or -not $script:Ready) { return }
    $rtts = @()
    foreach ($t in $script:Targets) { if (-not $t.IsHop) { $a = Get-RingAvg $t.Ring 30; if (-not [double]::IsNaN($a)) { $rtts += $a } } }
    $idle = [double]::NaN; if ($rtts.Count) { $idle = ($rtts | Measure-Object -Average).Average }
    $script:SpeedCtx = @{ BgRx = (Get-RingAvg $script:RxRing 10); BgTx = (Get-RingAvg $script:TxRing 10); Idle = $idle; Loaded = (New-Object System.Collections.Generic.List[double]) }
    $script:Speed.Start($Config.SpeedDownUrl, $Config.SpeedUpUrl, [long]$Config.SpeedUpBytes, [int]$Config.SpeedMaxSeconds)
    $script:SpeedPending = $true
}

function Complete-SpeedTest {
    $s = $script:Speed; $c = $script:SpeedCtx
    $script:SpeedPending = $false
    $script:SpeedExcludeUntil = (Get-Date).AddSeconds(3)
    $loaded = [double]::NaN
    if ($c.Loaded.Count) { $loaded = ($c.Loaded | Measure-Object -Average).Average }
    $script:SpeedLog.WriteLine((Join-Csv @(
        $s.StartedAt.ToString('yyyy-MM-dd HH:mm:ss'), $s.DownMbps, $s.UpMbps, $s.DownMB, $s.UpMB, $c.BgRx, $c.BgTx, $c.Idle, $loaded, $s.Error.Trim())))
    $script:LastSpeed = [pscustomobject]@{ Time = $s.StartedAt; Down = $s.DownMbps; Up = $s.UpMbps; Idle = $c.Idle; Loaded = $loaded; Error = $s.Error.Trim() }
    if ($s.Error) {
        Write-Event '速度' ('速度測定エラー: ' + $s.Error.Trim()) Yellow
    } else {
        Write-Event '速度' ('下り {0:0.0} Mbps / 上り {1:0.0} Mbps (測定前の自PC通信 {2:0.0} Mbps)' -f $s.DownMbps, $s.UpMbps, $c.BgRx) Cyan
    }
}

# ============================================================
#  1秒ごとの計測
# ============================================================
function Test-LinkUp {
    try {
        foreach ($n in [Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces()) {
            if ($n.Id -eq $script:NicId) { return ($n.OperationalStatus -eq [Net.NetworkInformation.OperationalStatus]::Up) }
        }
        return $false
    } catch { return $true }
}

function Invoke-Tick {
    $now = Get-Date
    $TL = $script:Targets
    $res = [NetMon.Prober]::Run($script:Specs, $Config.TimeoutMs)

    # 自PCの通信量 (Mbps)
    $rxMbps = 0.0; $txMbps = 0.0
    try {
        $st = $script:Nic.GetIPv4Statistics()
        $rx = $st.BytesReceived; $tx = $st.BytesSent
        $sec = ($now - $script:LastStatTime).TotalSeconds
        if ($script:LastRx -ge 0 -and $sec -gt 0 -and $rx -ge $script:LastRx -and $tx -ge $script:LastTx) {
            $rxMbps = ($rx - $script:LastRx) * 8 / $sec / 1e6
            $txMbps = ($tx - $script:LastTx) * 8 / $sec / 1e6
        }
        $script:LastRx = $rx; $script:LastTx = $tx; $script:LastStatTime = $now
    } catch { }

    # 宅内ルーターが無応答のときだけ LAN のリンク状態を確認
    $link = $true
    if (-not $res[$script:GwIndex].Ok) { $link = Test-LinkUp }
    if ($link -ne $script:LinkUp) {
        if ($link) {
            Write-Event 'リンク' ('LAN接続が復旧 (切断 {0:0}秒)' -f ($now - $script:LinkDownAt).TotalSeconds) Green
            $script:NeedRetrace = $true
        } else {
            Write-Event 'リンク' 'LANケーブル/Wi-Fi の接続が切れました (この間はロスとして集計しません)' Yellow
            $script:LinkDownAt = $now
            Close-Incident
        }
        $script:LinkUp = $link
    }

    $excluded = $script:Speed.Running -or ($now -lt $script:SpeedExcludeUntil)
    $state = 0
    if (-not $link) { $state = 2 } elseif ($excluded) { $state = 1 }

    $pos = $script:RingPos
    $script:StateRing[$pos] = $state
    $script:RxRing[$pos] = $rxMbps
    $script:TxRing[$pos] = $txMbps

    $raw = New-Object System.Text.StringBuilder
    [void]$raw.Append($now.ToString('yyyy-MM-dd HH:mm:ss')).Append(',').Append([int]$link).Append(',').Append([int]$excluded)
    [void]$raw.Append(',').Append($rxMbps.ToString('0.00', $Inv)).Append(',').Append($txMbps.ToString('0.00', $Inv))

    $endSent = 0; $endLost = 0; $gwLost = $false
    $loadSum = 0.0; $loadCnt = 0
    for ($i = 0; $i -lt $TL.Count; $i++) {
        $t = $TL[$i]; $r = $res[$i]
        if ($state -eq 2) { $t.Ring[$pos] = [double]::NaN; [void]$raw.Append(',-'); continue }
        if ($r.Ok) {
            $t.Ring[$pos] = $r.Rtt
            [void]$raw.Append(',').Append($r.Rtt.ToString('0.0', $Inv))
        } else {
            $t.Ring[$pos] = -1
            [void]$raw.Append(',LOSS')
        }
        if ($state -eq 1) {
            if (-not $t.IsHop -and $r.Ok) { $loadSum += $r.Rtt; $loadCnt++ }
            continue   # 速度測定中の秒は集計から除外
        }
        $t.MSent++; $t.TSent++
        if ($r.Ok) {
            $t.MRtts.Add($r.Rtt)
            if ($t.Ttl -gt 0 -and $r.From -ne $t.Address) { $t.MOther++ }
        } else {
            $t.MLost++; $t.TLost++
        }
        if (-not $t.IsHop) { $endSent++; if (-not $r.Ok) { $endLost++ } }
        elseif ($i -eq $script:GwIndex -and -not $r.Ok) { $gwLost = $true }
    }
    $script:RawLog.WriteLine($raw.ToString())
    $script:RingPos = ($pos + 1) % $script:RingLen

    $m = $script:M
    $m.Secs++; $m.RxSum += $rxMbps; $m.TxSum += $txMbps
    if ($rxMbps -gt $m.RxMax) { $m.RxMax = $rxMbps }
    if ($state -eq 0) { $m.Measured++ } elseif ($state -eq 1) { $m.Excl++ } else { $m.Link++ }
    if ($script:SpeedPending -and $loadCnt -gt 0) { $script:SpeedCtx.Loaded.Add($loadSum / $loadCnt) }

    if ($state -eq 0) {
        Update-Incident $now $res (($endLost -gt 0) -or $gwLost) (($endSent -gt 0) -and ($endLost -eq $endSent)) $rxMbps $txMbps
    }
}

# ============================================================
#  コンソール表示
# ============================================================
$Lines = New-Object System.Collections.Generic.List[object]
$Cur = New-Object System.Collections.Generic.List[object]
function W([string]$Text, [ConsoleColor]$Color = [ConsoleColor]::Gray) { $script:Cur.Add([object[]]@($Text, $Color)) }
function NL { $script:Lines.Add($script:Cur); $script:Cur = New-Object System.Collections.Generic.List[object] }

function Get-WindowStats {
    # 直近60秒の送信数/損失数 (速度測定中・リンク断の秒は除く)
    $TL = $script:Targets
    $sent = New-Object int[] $TL.Count; $lost = New-Object int[] $TL.Count
    $avg = New-Object double[] $TL.Count; $max = New-Object double[] $TL.Count
    for ($i = 0; $i -lt $TL.Count; $i++) {
        $ring = $TL[$i].Ring; $s = 0; $l = 0; $sum = 0.0; $mx = -1.0
        for ($k = 0; $k -lt $script:RingLen; $k++) {
            if ($script:StateRing[$k] -ne 0) { continue }
            $v = $ring[$k]
            if ([double]::IsNaN($v)) { continue }
            $s++
            if ($v -lt 0) { $l++ } else { $sum += $v; if ($v -gt $mx) { $mx = $v } }
        }
        $sent[$i] = $s; $lost[$i] = $l
        if ($s - $l -gt 0) { $avg[$i] = $sum / ($s - $l) } else { $avg[$i] = [double]::NaN }
        $max[$i] = $mx
    }
    return @{ Sent = $sent; Lost = $lost; Avg = $avg; Max = $max }
}

function Format-Ms([double]$v) {
    if ([double]::IsNaN($v) -or $v -lt 0) { return '-' }
    if ($v -lt 10) { return $v.ToString('0.0', $Inv) }
    return $v.ToString('0', $Inv)
}

function Get-LossColor([double]$Pct) {
    if ($Pct -lt 0) { return [ConsoleColor]::DarkGray }
    if ($Pct -ge $Config.LossThresholdPct) { return [ConsoleColor]::Red }
    if ($Pct -gt 0) { return [ConsoleColor]::Yellow }
    return [ConsoleColor]::Green
}

function Show-Dashboard {
    if (-not $script:UiOk) { return }
    try {
        $ScrW = [Console]::WindowWidth - 1
        $ScrH = [Console]::WindowHeight
        if ($ScrW -ne $script:LastW -or $ScrH -ne $script:LastH) { [Console]::Clear(); $script:LastW = $ScrW; $script:LastH = $ScrH; $script:PrevLines = 0 }
        $script:Lines.Clear(); $script:Cur.Clear()
        $now = Get-Date
        $up = $now - $script:StartTime

        W ' NetMon 回線品質モニタ ' Cyan
        W ('{0}  稼働 {1}:{2:mm\:ss}' -f $now.ToString('yyyy-MM-dd HH:mm:ss'), [int][Math]::Floor($up.TotalHours), $up)
        W '   [S]速度測定  [R]経路再探索  [Q]終了' DarkGray
        NL

        if (-not $script:Ready) {
            NL
            W ' ネットワークに接続されていません。接続を待っています...' Yellow; NL
        } else {
            $nic = $script:Nic
            $kind = '有線'; if ($script:IsWifi) { $kind = 'Wi-Fi' }
            W (' 接続: {0} ({1} {2}Mbps)' -f $nic.Name, $kind, [math]::Round($nic.Speed / 1e6))
            $rx = $script:RxRing[($script:RingPos - 1 + $script:RingLen) % $script:RingLen]
            $tx = $script:TxRing[($script:RingPos - 1 + $script:RingLen) % $script:RingLen]
            W ('   自PCの通信量: 受信 {0:0.00} / 送信 {1:0.00} Mbps' -f $rx, $tx) DarkGray
            if ($script:IsWifi) { W '   ※Wi-Fi接続中: 証拠として使うなら有線LAN推奨' Yellow }
            NL
            W (' ' + ('-' * ($ScrW - 1))) DarkGray; NL

            $ws = Get-WindowStats
            # 画面幅に合わせて列幅を調整 (狭いときは「最大」列を省き、機器名を詰める)
            $showMax = ($ScrW -ge 100)
            $maxW = 0; if ($showMax) { $maxW = 7 }
            $devW = [Math]::Max(14, [Math]::Min(36, $ScrW - 1 - 4 - 11 - 7 - 7 - $maxW - 9 - 9 - 2))
            $spark = [Math]::Min($script:RingLen, $ScrW - (1 + 4 + 11 + $devW + 7 + 7 + $maxW + 9 + 9 + 2))
            $hdr = ' ' + [NetMon.Util]::Fit('ID', 4) + [NetMon.Util]::Fit('区間', 11) + [NetMon.Util]::Fit('機器/宛先', $devW) +
                   [NetMon.Util]::FitRight('RTT', 7) + [NetMon.Util]::FitRight('平均', 7) + [NetMon.Util]::FitRight('最大', $maxW) +
                   [NetMon.Util]::FitRight('ロス60s', 9) + [NetMon.Util]::FitRight('累計ロス', 9) + '  '
            W $hdr DarkCyan
            if ($spark -ge 10) { W ('直近{0}秒 .正常 x=ロス' -f $spark) DarkCyan }
            NL

            $lastPos = ($script:RingPos - 1 + $script:RingLen) % $script:RingLen
            for ($i = 0; $i -lt $script:Targets.Count; $i++) {
                $t = $script:Targets[$i]
                if ($t.IsHop -and $t.Ptr) { $disp = '{0} {1}' -f $t.Address, $t.Ptr }
                elseif ($t.IsHop) { $disp = '{0} {1}' -f $t.Address, $t.Name }
                else { $disp = '{0} {1}' -f $t.Name, $t.Address }
                W (' ' + [NetMon.Util]::Fit($t.Id, 4) + [NetMon.Util]::Fit($t.Layer, 11) + [NetMon.Util]::Fit($disp, $devW))
                $last = $t.Ring[$lastPos]
                if ($last -lt 0) { W ([NetMon.Util]::FitRight('LOSS', 7)) Red }
                else { W ([NetMon.Util]::FitRight((Format-Ms $last), 7)) }
                W ([NetMon.Util]::FitRight((Format-Ms $ws.Avg[$i]), 7) + [NetMon.Util]::FitRight((Format-Ms $ws.Max[$i]), $maxW))
                $lp = -1.0; if ($ws.Sent[$i] -gt 0) { $lp = 100.0 * $ws.Lost[$i] / $ws.Sent[$i] }
                $lpText = '-'; if ($lp -ge 0) { $lpText = $lp.ToString('0.0', $Inv) + '%' }
                W ([NetMon.Util]::FitRight($lpText, 9)) (Get-LossColor $lp)
                $tp = -1.0; if ($t.TSent -gt 0) { $tp = 100.0 * $t.TLost / $t.TSent }
                $tpText = '-'; if ($tp -ge 0) { $tpText = $tp.ToString('0.00', $Inv) + '%' }
                W ([NetMon.Util]::FitRight($tpText, 9) + '  ') (Get-LossColor $tp)
                if ($spark -ge 10) {
                    $sb = New-Object System.Text.StringBuilder; $col = $null
                    for ($k = $spark; $k -ge 1; $k--) {
                        $p = ($script:RingPos - $k + $script:RingLen) % $script:RingLen
                        $s = $script:StateRing[$p]; $v = $t.Ring[$p]
                        if ($s -lt 0 -or [double]::IsNaN($v)) { $ch = ' '; $c = [ConsoleColor]::DarkGray }
                        elseif ($s -eq 2) { $ch = '-'; $c = [ConsoleColor]::DarkGray }
                        elseif ($s -eq 1) { $ch = 's'; $c = [ConsoleColor]::DarkGray }
                        elseif ($v -lt 0) { $ch = 'x'; $c = [ConsoleColor]::Red }
                        elseif ($v -ge 100) { $ch = '!'; $c = [ConsoleColor]::Yellow }
                        else { $ch = '.'; $c = [ConsoleColor]::DarkGreen }
                        if ($c -ne $col -and $sb.Length) { W $sb.ToString() $col; [void]$sb.Clear() }
                        $col = $c; [void]$sb.Append($ch)
                    }
                    if ($sb.Length) { W $sb.ToString() $col }
                }
                NL
            }
            W (' ' + ('-' * ($ScrW - 1))) DarkGray; NL

            $v = Get-Verdict $ws.Sent $ws.Lost
            $vc = @([ConsoleColor]::Green, [ConsoleColor]::DarkYellow, [ConsoleColor]::Yellow, [ConsoleColor]::Red)[$v.Level]
            W ' 判定(直近60秒): '; W $v.Text $vc; NL

            if ($script:Inc) {
                $inc = $script:Inc; $es = 0; $el = 0
                for ($i = 0; $i -lt $script:Targets.Count; $i++) { if (-not $script:Targets[$i].IsHop) { $es += $inc.Sent[$i]; $el += $inc.Lost[$i] } }
                W (' ロス発生中: {0} から {1:0}秒  外部宛先 {2}/{3}発 損失' -f $inc.Start.ToString('HH:mm:ss'), ($now - $inc.Start).TotalSeconds, $el, $es) Red
            } else {
                W (' ロス事象: 起動後 {0}件' -f $script:IncToday) DarkGray
            }
            NL

            $er = @()
            for ($i = 0; $i -lt $script:Targets.Count; $i++) {
                if (-not $script:Targets[$i].IsHop -and -not [double]::IsNaN($ws.Avg[$i])) { $er += $ws.Avg[$i] }
            }
            if ($er.Count) {
                $ea = ($er | Measure-Object -Average).Average
                W (' 遅延: 外部宛先の平均 {0:0.0} ms' -f $ea)
                if ($script:Baseline) {
                    $bc = [ConsoleColor]::DarkGray
                    if ($ea -gt [Math]::Max($script:Baseline * 2, $script:Baseline + 20)) { $bc = [ConsoleColor]::Yellow }
                    W (' (平常時 {0:0.0} ms)' -f $script:Baseline) $bc
                }
                NL
            }

            if ($Config.SpeedTestEveryMin -gt 0) {
                if ($script:Speed.Running) {
                    W (' 速度: 測定中... ({0:0}秒経過、この間はロス集計から除外)' -f ($now - $script:Speed.StartedAt).TotalSeconds) Cyan
                } elseif ($script:LastSpeed) {
                    $ls = $script:LastSpeed
                    if ($ls.Error) { W (' 速度: {0} 測定失敗 ({1})' -f $ls.Time.ToString('HH:mm'), $ls.Error) Yellow }
                    else { W (' 速度: 下り {0:0.0} / 上り {1:0.0} Mbps ({2})' -f $ls.Down, $ls.Up, $ls.Time.ToString('HH:mm')) Cyan }
                } else { W ' 速度: 未測定' DarkGray }
                W ('  次回 {0}' -f $script:NextSpeed.ToString('HH:mm')) DarkGray
                if ($ls -and -not $ls.Error -and -not [double]::IsNaN($ls.Loaded)) { W ('  測定中RTT {0:0.0} ms (通常 {1:0.0} ms)' -f $ls.Loaded, $ls.Idle) DarkGray }
                NL
            }
        }
        W (' 記録先: ' + $Config.LogDir) DarkGray; NL
        W ' --- 最近のイベント ---' DarkCyan; NL
        $room = $ScrH - 1 - $script:Lines.Count
        for ($k = $script:Events.Count - 1; $k -ge 0 -and $room -gt 0; $k--) {
            $e = $script:Events[$k]
            W (' {0}  ' -f $e.Time) DarkGray; W ('[{0}] {1}' -f $e.Type, $e.Message) $e.Color; NL
            $room--
        }

        # 描画 (画面を消さずに上書きしてちらつきを防ぐ)
        [Console]::CursorVisible = $false
        [Console]::SetCursorPosition(0, 0)
        $max = [Math]::Min($script:Lines.Count, $ScrH - 1)
        $li = 0
        for (; $li -lt $max; $li++) {
            $used = 0
            foreach ($seg in $script:Lines[$li]) {
                $txt = [string]$seg[0]
                $wd = [NetMon.Util]::Width($txt)
                if ($used + $wd -gt $ScrW) { $txt = [NetMon.Util]::Fit($txt, $ScrW - $used); $wd = $ScrW - $used }
                if ($wd -le 0) { break }
                [Console]::ForegroundColor = $seg[1]
                [Console]::Write($txt)
                $used += $wd
            }
            [Console]::ForegroundColor = [ConsoleColor]::Gray
            [Console]::Write((' ' * ($ScrW - $used)) + "`n")
        }
        for (; $li -lt [Math]::Min($script:PrevLines, $ScrH - 1); $li++) { [Console]::Write((' ' * $ScrW) + "`n") }
        $script:PrevLines = $max
        $script:UiErr = 0
    } catch {
        $script:UiErr++
        if ($script:UiErr -gt 10) { $script:UiOk = $false }
    }
}

function Update-Title {
    try {
        $title = 'NetMon 回線品質モニタ'
        if ($script:Inc) { $title = '[ロス発生中] ' + $title }
        $Host.UI.RawUI.WindowTitle = $title
    } catch { }
}

function Read-Keys {
    try {
        while ([Console]::KeyAvailable) {
            $k = [Console]::ReadKey($true)
            switch ($k.Key) {
                'S' { if ($Config.SpeedTestEveryMin -gt 0) { Write-Event '操作' '手動で速度測定を開始' Gray; Start-SpeedTest } }
                'R' { Write-Event '操作' '経路を再探索' Gray; $script:NeedRetrace = $true }
                'Q' { $script:Quit = $true }
            }
        }
    } catch { }
}

# ============================================================
#  メインループ
# ============================================================
$RingLen = 60
$RingPos = 0
$StateRing = New-Object int[] $RingLen
$RxRing = New-Object double[] $RingLen
$TxRing = New-Object double[] $RingLen
for ($i = 0; $i -lt $RingLen; $i++) { $StateRing[$i] = -1; $RxRing[$i] = [double]::NaN; $TxRing[$i] = [double]::NaN }
$Targets = @()
$Ready = $false
$NeedRetrace = $true
$RouteSig = ''
$KnownHops = @{}
$EndpointIps = @{}
$Inc = $null
$IncToday = 0
$LinkUp = $true
$LinkDownAt = Get-Date
$LastRx = -1; $LastTx = -1; $LastStatTime = Get-Date
$RttBase = New-Object System.Collections.Generic.List[double]
$Baseline = $null
$Speed = New-Object NetMon.SpeedTest
$SpeedPending = $false
$SpeedExcludeUntil = Get-Date
$LastSpeed = $null
$NextSpeed = (Get-Date).AddSeconds($Config.FirstSpeedTestSec)
if ($Config.SpeedTestEveryMin -le 0) { $NextSpeed = [datetime]::MaxValue }
$UiOk = $true; $UiErr = 0; $LastW = 0; $LastH = 0; $PrevLines = 0
$Quit = $false
$StartTime = Get-Date
$LastTick = $null
$NextRetry = Get-Date

try { $Host.UI.RawUI.WindowTitle = 'NetMon 回線品質モニタ' } catch { }
try { [Console]::Clear() } catch { $UiOk = $false }

Open-DayLogs
Reset-Minute
Write-Event '開始' ("NetMon v$Version 起動") Gray

$clock = [Diagnostics.Stopwatch]::StartNew()
$nextMs = 0
try {
    while (-not $Quit) {
        # 次の計測時刻まで待つ (待ち時間中もキー入力を受け付ける)
        while ($clock.ElapsedMilliseconds -lt $nextMs -and -not $Quit) {
            Read-Keys
            Start-Sleep -Milliseconds ([Math]::Min(50, [Math]::Max(1, $nextMs - $clock.ElapsedMilliseconds)))
        }
        if ($Quit) { break }
        if ($RunSeconds -gt 0 -and $clock.Elapsed.TotalSeconds -ge $RunSeconds) { break }
        $nextMs += $Config.IntervalMs
        if ($clock.ElapsedMilliseconds - $nextMs -gt 3 * $Config.IntervalMs) { $nextMs = $clock.ElapsedMilliseconds + $Config.IntervalMs }

        try {
            $now = Get-Date
            # スリープ等で計測が止まっていた
            if ($LastTick -and ($now - $LastTick).TotalSeconds -gt 5) {
                Write-Event '中断' ('{0} から {1:0}秒間 計測が止まっていました (スリープ等)。この間のデータはありません' -f $LastTick.ToString('HH:mm:ss'), ($now - $LastTick).TotalSeconds) DarkYellow
                Close-Incident
                $NeedRetrace = $true
            }
            $LastTick = $now

            # 分・日の切り替え
            if ($M.Key -ne $now.ToString('yyyyMMddHHmm')) {
                Save-Minute
                if ($now.Date -ne $CurDay) { Open-DayLogs; if ($Ready) { Open-RawLog } }
            }

            if ($NeedRetrace -and $now -ge $NextRetry) {
                if (-not (Initialize-Targets)) {
                    if ($Ready -or -not $script:WaitLogged) { Write-Event '待機' 'ネットワーク経路を取得できません。10秒ごとに再試行します' Yellow; $script:WaitLogged = $true }
                    $NextRetry = $now.AddSeconds(10)
                } else { $script:WaitLogged = $false }
                $LastTick = Get-Date   # 経路探索にかかった時間を「中断」と誤検知しない
            }

            if ($Ready) {
                Invoke-Tick
                if ($Config.SpeedTestEveryMin -gt 0) {
                    if ($SpeedPending -and -not $Speed.Running) { Complete-SpeedTest }
                    if ($now -ge $NextSpeed) {
                        if ($LinkUp) { Start-SpeedTest }
                        $NextSpeed = Get-NextSpeedTime $now
                    }
                }
            }
        } catch {
            Write-Event 'エラー' ($_.Exception.Message) Red
        }
        Show-Dashboard
        Update-Title
    }
} finally {
    try { Close-Incident; Save-Minute } catch { }
    try { Write-Event '停止' 'NetMon 停止' Gray } catch { }
    foreach ($w in @($RawLog, $SummaryLog, $OverviewLog, $EventLog, $IncidentLog, $SpeedLog, $RouteLog)) { if ($w) { try { $w.Dispose() } catch { } } }
    try { [Console]::CursorVisible = $true; [Console]::ForegroundColor = [ConsoleColor]::Gray } catch { }
    $Mutex.ReleaseMutex(); $Mutex.Dispose()
}
