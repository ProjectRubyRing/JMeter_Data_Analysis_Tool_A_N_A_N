#requires -Version 5.1
<#
.SYNOPSIS
    A5:SQL Mk-2 登録データベース (EC2踏み台 SSHトンネル経由 Aurora Serverless v2 MySQL 8.4.7)
    自動CSV一括エクスポート ＋ 性能試験用CSVデータ総合分析 Excelレポート生成ツール

.DESCRIPTION
    Windows 11 / PowerShell 5.1 標準機能のみで動作します。
    - A5:SQL Mk-2 のレジストリまたはポータブル版のDBリストから接続設定を読込 (MySQL + SSH2トンネル + SSL/TLS)
    - Windows 11 標準 OpenSSH (ssh.exe) によるEC2踏み台SSHポートフォワーディング
    - 純.NET実装 MySQL 8.4.7 ワイヤプロトコルクライアント (TLS 1.2/1.3, caching_sha2_password,
      mysql_native_password, mysql_clear_password対応, 読み取り専用セッション, 全テーブルCSVストリーミング出力)
    - 高速C#ストリーミング分析エンジン (CSVフォルダ解析 ＆ CSV未出力のDB直接ストリーミング解析対応。
      文字コード自動判定, 件数/容量/推定InnoDBサイズ,
      カラム型/NULL率/カーディナリティ/選択性/偏り(ジニ係数・エントロピー・Top-N占有率),
      カテゴリ分布・割合, 数値パーセンタイル・10区間ヒストグラム, 時系列(月別/曜日別/時間帯別),
      テーブル間FK自動推定・1:Nファンアウト倍率・孤立レコード検出, 性能試験警告・JMeter設計ガイド)
    - Excelレポート出力 (Meiryo UI, 文字色 #333333, アクセント #1A2B4C, 警告 #800020, モノトーン基調,
      グラフ付き。Excel COM優先、未インストール環境向けOpenXMLフォールバック対応)
    - テスト開始前・開始後の分析Excelファイル2件を比較し、更新・追加・変更なしテーブルと
      カラム/分布/数値/時系列差分をモノトーン基調Excelレポートとして出力
#>

[CmdletBinding()]
param(
    [ValidateSet('Gui', 'AnalyzeOnly', 'ExportAndAnalyze', 'ExportOnly', 'DirectAnalyze', 'CompareAnalysis')]
    [string]$Mode = 'Gui',

    [string]$A5DbName = '',
    [string]$A5RegistryRoot = 'HKCU:\Software\mmatsubara',
    [string]$ConnectionStringOverride = '',

    [string]$InputFolder = '',
    [string]$OutputPath = '',
    [string]$BeforeAnalysisPath = '',
    [string]$AfterAnalysisPath = '',

    [string]$DbHost = '',
    [int]$DbPort = 0,
    [string]$DbName = '',
    [string]$DbUser = '',
    [string]$DbPassword = '',

    [string]$SshHost = '',
    [int]$SshPort = 0,
    [string]$SshUser = '',
    [string]$SshKeyPath = '',
    [string]$SshPassphrase = '',
    [switch]$DisableSsh,

    [ValidateSet('Auto', 'Disabled', 'Preferred', 'Required', 'VerifyCA')]
    [string]$SslMode = 'Auto',
    [string]$SslCaCertPath = '',

    [int]$MaxRowsPerTable = 0,
    [string[]]$IncludeTables = @(),
    [string[]]$ExcludeTables = @(),

    [ValidateSet('Auto', 'Com', 'OpenXml')]
    [string]$ExcelEngine = 'Auto',

    [switch]$OpenAfterCreate,
    [switch]$NoGui,
    [string]$A5DbListPath = '',
    [switch]$Resume,
    [switch]$DetailedAnalysis
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ==============================================================================
#  Design System Constants (Strict Monotone + Midnight Navy + Burgundy)
#  High-contrast palette: avoids same-color-family text/background pairings.
# ==============================================================================
$script:THEME = @{
    FontName        = 'Meiryo UI'
    TextHex         = '#333333'   # 濃いグレー (本文・軸・ラベル)
    MutedHex        = '#595959'   # ミディアムグレー (補足・ラベル, 白背景比 7.0:1)
    AccentHex       = '#1A2B4C'   # ミッドナイトネイビー (アクセント・ヘッダー・主要系列)
    AccentLightHex  = '#F2F2F2'   # 純ニュートラル薄グレー (小見出し・サブタイトル背景)
    TotalBgHex      = '#EAEAEA'   # 純ニュートラルグレー (合計行背景)
    WarnHex         = '#800020'   # バーガンディ (警告・注意文字色)
    WarnBgHex       = '#F5F5F5'   # 純ニュートラル薄グレー (警告セル背景: 非同系統で高コントラスト)
    ZebraHex        = '#F7F7F7'   # 純モノトーン交互行
    CardBgHex       = '#FFFFFF'   # KPIカード純白背景
    WhiteHex        = '#FFFFFF'
    BorderHex       = '#D9D9D9'   # 罫線 (薄グレー)
    BorderDarkHex   = '#8C8C8C'   # 強調罫線 (濃グレー)
    GridlineHex     = '#E6E6E6'   # グラフ目盛線
    HeaderBorderHex = '#4D5B75'   # ヘッダー区切り罫線
    ChartPalette    = @('#1A2B4C', '#4D4D4D', '#737373', '#999999', '#BFBFBF', '#333333', '#666666', '#A6A6A6')
}

function ConvertTo-ExcelOleColor {
    param([Parameter(Mandatory = $true)][string]$Hex)
    $clean = $Hex.TrimStart('#')
    if ($clean.Length -ne 6) { throw "Invalid hex color: $Hex" }
    $r = [Convert]::ToInt32($clean.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($clean.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($clean.Substring(4, 2), 16)
    # Excel COM expects BGR (OLE_COLOR)
    return ($r + ($g * 256) + ($b * 65536))
}

function ConvertTo-DrawingColor {
    param([Parameter(Mandatory = $true)][string]$Hex)
    $clean = $Hex.TrimStart('#')
    $r = [Convert]::ToInt32($clean.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($clean.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($clean.Substring(4, 2), 16)
    return [System.Drawing.Color]::FromArgb($r, $g, $b)
}

# ==============================================================================
#  Module 1 & 2: Embedded C# 5.0 Core
#   - AuroraAnalyzer.MySql84Client (Pure .NET 4.8 MySQL 8.4.7 Wire Protocol + TLS)
#   - AuroraAnalyzer.CsvAnalysisEngine (High-speed Streaming CSV Profiler)
# ==============================================================================
$script:CSharpCoreSource = @'
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net.Security;
using System.Net.Sockets;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Security.Authentication;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;

namespace AuroraAnalyzer
{
    public static class AtomicFile
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool MoveFileEx(string existing, string destination, uint flags);

        // The staging file and destination are always on the same volume.
        public static void Publish(string stagingPath, string destination)
        {
            const uint ReplaceExisting = 1, WriteThrough = 8;
            if (!MoveFileEx(stagingPath, destination, ReplaceExisting | WriteThrough))
            {
                var error = new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
                throw new IOException("Cannot publish " + destination + ": " + error.Message, error);
            }
        }
    }

    public class OperationCancelState
    {
        private volatile bool _cancelled;
        private readonly object _sync = new object();
        private readonly List<Action> _abortActions = new List<Action>();

        public bool IsCancelled { get { return _cancelled; } }

        public void ThrowIfCancelled()
        {
            if (_cancelled)
            {
                throw new OperationCanceledException("ユーザーによって処理が中断されました。");
            }
        }

        public IDisposable RegisterAbort(Action abortAction)
        {
            if (abortAction == null) return new NoopDisposable();
            bool invokeNow = false;
            lock (_sync)
            {
                if (_cancelled) invokeNow = true;
                else _abortActions.Add(abortAction);
            }
            if (invokeNow)
            {
                try { abortAction(); } catch { }
                return new NoopDisposable();
            }
            return new AbortRegistration(this, abortAction);
        }

        public IDisposable RegisterProcessKill(Process proc)
        {
            if (proc == null) return new NoopDisposable();
            return RegisterAbort(delegate()
            {
                try
                {
                    if (!proc.HasExited) proc.Kill();
                }
                catch { }
            });
        }

        public IDisposable AttachConsoleCancel()
        {
            ConsoleCancelEventHandler handler = delegate(object sender, ConsoleCancelEventArgs e)
            {
                e.Cancel = true;
                Cancel();
            };
            try
            {
                Console.CancelKeyPress += handler;
            }
            catch
            {
                return new NoopDisposable();
            }
            return new ConsoleCancelRegistration(handler);
        }

        internal void UnregisterAbort(Action abortAction)
        {
            lock (_sync)
            {
                _abortActions.Remove(abortAction);
            }
        }

        public void Cancel()
        {
            List<Action> snapshot = null;
            lock (_sync)
            {
                if (_cancelled) return;
                _cancelled = true;
                snapshot = new List<Action>(_abortActions);
                _abortActions.Clear();
            }
            for (int i = 0; i < snapshot.Count; i++)
            {
                try { snapshot[i](); } catch { }
            }
        }

        private class NoopDisposable : IDisposable
        {
            public void Dispose() { }
        }

        private class AbortRegistration : IDisposable
        {
            private OperationCancelState _owner;
            private Action _action;

            public AbortRegistration(OperationCancelState owner, Action action)
            {
                _owner = owner;
                _action = action;
            }

            public void Dispose()
            {
                OperationCancelState o = _owner;
                Action a = _action;
                _owner = null;
                _action = null;
                if (o != null && a != null) o.UnregisterAbort(a);
            }
        }

        private class ConsoleCancelRegistration : IDisposable
        {
            private ConsoleCancelEventHandler _handler;

            public ConsoleCancelRegistration(ConsoleCancelEventHandler handler)
            {
                _handler = handler;
            }

            public void Dispose()
            {
                ConsoleCancelEventHandler h = _handler;
                _handler = null;
                if (h != null)
                {
                    try { Console.CancelKeyPress -= h; } catch { }
                }
            }
        }
    }

    public class BackgroundTaskHandle<T> : IDisposable
    {
        private readonly ManualResetEventSlim _done = new ManualResetEventSlim(false);
        private readonly ConcurrentQueue<string> _messages = new ConcurrentQueue<string>();
        private readonly Thread _worker;
        private T _result;
        private Exception _error;
        private long _rowsWritten;
        public long RowsWritten { get { return Interlocked.Read(ref _rowsWritten); } }

        internal BackgroundTaskHandle(Func<Action<string>, T> work)
            : this((emit, reportRows) => work(emit)) { }

        internal BackgroundTaskHandle(Func<Action<string>, Action<long>, T> work)
        {
            _worker = new Thread(delegate()
            {
                try
                {
                    _result = work(delegate(string msg)
                    {
                        if (!string.IsNullOrEmpty(msg)) _messages.Enqueue(msg);
                    }, rows => Interlocked.Exchange(ref _rowsWritten, rows));
                }
                catch (Exception ex)
                {
                    _error = ex;
                }
                finally
                {
                    _done.Set();
                }
            });
            _worker.IsBackground = true;
            _worker.Start();
        }

        public bool WaitStep(int milliseconds)
        {
            return _done.Wait(milliseconds > 0 ? milliseconds : 25);
        }

        public string[] DrainMessages()
        {
            if (_messages.IsEmpty) return new string[0];
            List<string> list = new List<string>();
            string m;
            while (_messages.TryDequeue(out m))
            {
                list.Add(m);
            }
            return list.ToArray();
        }

        public T GetResult()
        {
            _done.Wait();
            if (_error != null)
            {
                ExceptionDispatchInfo.Capture(_error).Throw();
            }
            return _result;
        }

        public void Dispose()
        {
            _worker.Join();
            _done.Dispose();
        }
    }

    public static class ComHelper
    {
        private static readonly CultureInfo EnUs = CultureInfo.GetCultureInfo("en-US");

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        public static int GetWindowProcessId(int hwnd)
        {
            uint pid;
            GetWindowThreadProcessId(new IntPtr(hwnd), out pid);
            return (int)pid;
        }

        public static void SetValue2(object range, object value)
        {
            if (range == null) return;
            range.GetType().InvokeMember(
                "Value2",
                System.Reflection.BindingFlags.SetProperty,
                null,
                range,
                new object[] { value },
                EnUs);
        }
    }

    public static class XmlText
    {
        public static string Escape(string value)
        {
            if (string.IsNullOrEmpty(value)) return value ?? "";
            bool dirty = false;
            for (int i = 0; i < value.Length; i++)
            {
                char ch = value[i];
                if ((ch <= 0x1F && ch != 0x09 && ch != 0x0A && ch != 0x0D) ||
                    ch == '&' || ch == '<' || ch == '>' || ch == '"' || ch == '\'')
                {
                    dirty = true;
                    break;
                }
            }
            if (!dirty) return value;
            StringBuilder sb = new StringBuilder(value.Length + 16);
            for (int i = 0; i < value.Length; i++)
            {
                char ch = value[i];
                if (ch <= 0x1F && ch != 0x09 && ch != 0x0A && ch != 0x0D) continue;
                switch (ch)
                {
                    case '&': sb.Append("&amp;"); break;
                    case '<': sb.Append("&lt;"); break;
                    case '>': sb.Append("&gt;"); break;
                    case '"': sb.Append("&quot;"); break;
                    case '\'': sb.Append("&apos;"); break;
                    default: sb.Append(ch); break;
                }
            }
            return sb.ToString();
        }
    }

    public sealed class ProcessOutputDrain
    {
        private readonly object _sync = new object();
        private readonly StringBuilder _stderr = new StringBuilder();
        private const int MaxStderrChars = 16384;

        public void Attach(Process proc)
        {
            if (proc == null) return;
            proc.OutputDataReceived += delegate(object s, DataReceivedEventArgs e) { };
            proc.ErrorDataReceived += delegate(object s, DataReceivedEventArgs e)
            {
                if (e.Data == null) return;
                lock (_sync)
                {
                    if (_stderr.Length < MaxStderrChars)
                    {
                        if (_stderr.Length > 0) _stderr.AppendLine();
                        int remaining = MaxStderrChars - _stderr.Length;
                        _stderr.Append(e.Data.Length <= remaining ? e.Data : e.Data.Substring(0, remaining));
                    }
                }
            };
            proc.BeginOutputReadLine();
            proc.BeginErrorReadLine();
        }

        public string GetStderrText()
        {
            lock (_sync) { return _stderr.ToString(); }
        }
    }

    public sealed class ExcelActivationException : InvalidOperationException
    {
        public ExcelActivationException(Exception cause)
            : base("Excel COM を起動できません。 " + cause.Message, cause) { }
    }

    // =========================================================================
    //  MySQL 8.4.7 Wire Protocol Client (Protocol v10 + TLS + caching_sha2_password)
    // =========================================================================
    public class MySqlColumnInfo
    {
        public string Name;
        public byte ColumnType;
        public ushort Charset;
        public uint ColumnLength;
        public ushort Flags;
    }

    public class MySqlQueryResult
    {
        public List<MySqlColumnInfo> Columns;
        public List<string[]> Rows;

        public MySqlQueryResult()
        {
            Columns = new List<MySqlColumnInfo>();
            Rows = new List<string[]>();
        }
    }

    public class MySql84Client : IDisposable
    {
        private const uint CLIENT_LONG_PASSWORD = 0x00000001;
        private const uint CLIENT_LONG_FLAG = 0x00000004;
        private const uint CLIENT_CONNECT_WITH_DB = 0x00000008;
        private const uint CLIENT_PROTOCOL_41 = 0x00000200;
        private const uint CLIENT_SSL = 0x00000800;
        private const uint CLIENT_TRANSACTIONS = 0x00002000;
        private const uint CLIENT_SECURE_CONNECTION = 0x00008000;
        private const uint CLIENT_MULTI_RESULTS = 0x00020000;
        private const uint CLIENT_PLUGIN_AUTH = 0x00080000;

        private TcpClient _tcp;
        private Stream _stream;
        private readonly byte[] _packetHeader = new byte[4];
        private byte[] _bulkPacket = new byte[8192];
        private bool _isSsl;
        private byte _seqId;
        private string _serverVersion;
        private string _caCertPath;
        private string _sslMode;
        private volatile bool _aborted;
        private OperationCancelState _cancelState;
        private IDisposable _cancelReg;

        public string ServerVersion { get { return _serverVersion; } }
        public bool IsSslActive { get { return _isSsl; } }

        public void Abort()
        {
            _aborted = true;
            Stream s = _stream;
            TcpClient t = _tcp;
            try { if (t != null) t.Close(); } catch { }
            try { if (s != null) s.Dispose(); } catch { }
        }

        public void Connect(
            string connectHost,
            int connectPort,
            string tlsTargetHost,
            string user,
            string password,
            string database,
            string sslMode,
            string caCertPath,
            int timeoutSeconds)
        {
            Connect(connectHost, connectPort, tlsTargetHost, user, password, database, sslMode, caCertPath, timeoutSeconds, null);
        }

        public void Connect(
            string connectHost,
            int connectPort,
            string tlsTargetHost,
            string user,
            string password,
            string database,
            string sslMode,
            string caCertPath,
            int timeoutSeconds,
            OperationCancelState cancelState)
        {
            _sslMode = string.IsNullOrEmpty(sslMode) ? "Preferred" : sslMode;
            _caCertPath = caCertPath;
            _isSsl = false;
            _seqId = 0;
            _aborted = false;
            _cancelState = cancelState;
            if (_cancelReg != null) { try { _cancelReg.Dispose(); } catch { } _cancelReg = null; }
            if (cancelState != null) cancelState.ThrowIfCancelled();

            _tcp = new TcpClient();
            _tcp.NoDelay = true;
            _tcp.ReceiveBufferSize = 1024 * 1024;
            _tcp.SendBufferSize = 262144;
            if (cancelState != null)
            {
                _cancelReg = cancelState.RegisterAbort(delegate() { Abort(); });
            }

            IAsyncResult ar = _tcp.BeginConnect(connectHost, connectPort, null, null);
            DateTime deadline = DateTime.UtcNow.AddSeconds(timeoutSeconds > 0 ? timeoutSeconds : 30);
            while (!ar.AsyncWaitHandle.WaitOne(100))
            {
                if (cancelState != null && cancelState.IsCancelled)
                {
                    Abort();
                    cancelState.ThrowIfCancelled();
                }
                if (DateTime.UtcNow >= deadline)
                {
                    Abort();
                    throw new TimeoutException(string.Format("MySQL TCP connection timed out ({0}:{1})", connectHost, connectPort));
                }
            }
            if (cancelState != null && cancelState.IsCancelled)
            {
                Abort();
                cancelState.ThrowIfCancelled();
            }
            try
            {
                _tcp.EndConnect(ar);
            }
            catch
            {
                if (cancelState != null && cancelState.IsCancelled) cancelState.ThrowIfCancelled();
                throw;
            }
            int ioTimeoutSeconds = Math.Max(300, timeoutSeconds > 0 ? timeoutSeconds : 300);
            _tcp.ReceiveTimeout = ioTimeoutSeconds * 1000;
            _tcp.SendTimeout = ioTimeoutSeconds * 1000;
            _stream = _tcp.GetStream();

            // 1. Read Initial Handshake Packet (v10)
            byte[] hs = ReadPacket();
            if (hs.Length == 0)
            {
                throw new IOException("Empty handshake packet from MySQL server.");
            }
            if (hs[0] == 0xFF)
            {
                throw ParseErrorPacket(hs);
            }
            if (hs[0] != 0x0A)
            {
                throw new NotSupportedException(string.Format("Unsupported MySQL protocol version: {0}", hs[0]));
            }

            int pos = 1;
            _serverVersion = ReadNullTerminatedString(hs, ref pos);
            pos += 4; // connection id

            byte[] salt1 = new byte[8];
            Array.Copy(hs, pos, salt1, 0, 8);
            pos += 8;
            pos += 1; // filler 0x00

            uint serverCaps = (uint)(hs[pos] | (hs[pos + 1] << 8));
            pos += 2;

            byte serverCharset = 45;
            if (pos < hs.Length)
            {
                serverCharset = hs[pos];
                pos += 1;
                pos += 2; // status flags
                if (pos + 2 <= hs.Length)
                {
                    uint upperCaps = (uint)(hs[pos] | (hs[pos + 1] << 8));
                    serverCaps |= (upperCaps << 16);
                    pos += 2;
                }
            }

            int authDataLen = 0;
            if (pos < hs.Length)
            {
                authDataLen = ((serverCaps & CLIENT_PLUGIN_AUTH) != 0) ? hs[pos] : 0;
                pos += 1;
            }
            pos += 10; // reserved

            byte[] nonce = salt1;
            if ((serverCaps & CLIENT_SECURE_CONNECTION) != 0 && pos < hs.Length)
            {
                int part2Len = Math.Max(13, authDataLen - 8);
                if (pos + part2Len > hs.Length) part2Len = hs.Length - pos;
                int actualPart2 = part2Len;
                if (actualPart2 > 0 && hs[pos + actualPart2 - 1] == 0)
                {
                    actualPart2--;
                }
                nonce = new byte[8 + actualPart2];
                Array.Copy(salt1, 0, nonce, 0, 8);
                Array.Copy(hs, pos, nonce, 8, actualPart2);
                pos += part2Len;
            }

            string authPlugin = "mysql_native_password";
            if ((serverCaps & CLIENT_PLUGIN_AUTH) != 0 && pos < hs.Length)
            {
                authPlugin = ReadNullTerminatedString(hs, ref pos);
            }

            // 2. Negotiate capabilities & TLS/SSL
            uint clientCaps = CLIENT_LONG_PASSWORD | CLIENT_LONG_FLAG | CLIENT_PROTOCOL_41 |
                              CLIENT_TRANSACTIONS | CLIENT_SECURE_CONNECTION |
                              CLIENT_MULTI_RESULTS | CLIENT_PLUGIN_AUTH;

            if (!string.IsNullOrEmpty(database))
            {
                clientCaps |= CLIENT_CONNECT_WITH_DB;
            }

            bool serverSupportsSsl = (serverCaps & CLIENT_SSL) != 0;
            bool wantSsl = false;
            if (string.Equals(_sslMode, "Disabled", StringComparison.OrdinalIgnoreCase))
            {
                wantSsl = false;
            }
            else if (string.Equals(_sslMode, "Required", StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(_sslMode, "VerifyCA", StringComparison.OrdinalIgnoreCase))
            {
                if (!serverSupportsSsl)
                {
                    throw new AuthenticationException("MySQL server does not advertise SSL/TLS support, but SslMode=" + _sslMode);
                }
                wantSsl = true;
            }
            else
            {
                // Preferred / Auto
                wantSsl = serverSupportsSsl;
            }

            byte clientCharset = 255; // utf8mb4_0900_ai_ci (MySQL 8.x); fallback 45 if older
            if (!string.IsNullOrEmpty(_serverVersion) && _serverVersion.StartsWith("5."))
            {
                clientCharset = 45; // utf8mb4_general_ci
            }

            if (wantSsl)
            {
                clientCaps |= CLIENT_SSL;
                // Send 32-byte SSLRequest packet
                using (MemoryStream ms = new MemoryStream())
                {
                    WriteUInt32LE(ms, clientCaps);
                    WriteUInt32LE(ms, 16777216);
                    ms.WriteByte(clientCharset);
                    ms.Write(new byte[23], 0, 23);
                    WritePacket(ms.ToArray());
                }

                string targetName = string.IsNullOrEmpty(tlsTargetHost) ? connectHost : tlsTargetHost;
                SslStream ssl = new SslStream(_stream, true, ValidateServerCertificate);
                SslProtocols protocols = SslProtocols.Tls12;
                try
                {
                    // Include TLS 1.3 (12288) on Windows 11 when supported
                    ssl.AuthenticateAsClient(targetName, null, protocols | (SslProtocols)12288, false);
                }
                catch (Exception ex)
                {
                    if (_cancelState != null && _cancelState.IsCancelled) _cancelState.ThrowIfCancelled();
                    if (ex is AuthenticationException)
                    {
                        try { ssl.Dispose(); } catch { }
                        throw;
                    }
                    try { ssl.Dispose(); } catch { }
                    ssl = new SslStream(_stream, false, ValidateServerCertificate);
                    ssl.AuthenticateAsClient(targetName, null, SslProtocols.Tls12, false);
                }
                _stream = ssl;
                _isSsl = true;
            }

            // 3. Send HandshakeResponse41
            byte[] authResp = ComputeAuthResponse(authPlugin, password ?? "", nonce);
            using (MemoryStream ms = new MemoryStream())
            {
                WriteUInt32LE(ms, clientCaps);
                WriteUInt32LE(ms, 16777216);
                ms.WriteByte(clientCharset);
                ms.Write(new byte[23], 0, 23);

                byte[] userBytes = Encoding.UTF8.GetBytes(user ?? "");
                ms.Write(userBytes, 0, userBytes.Length);
                ms.WriteByte(0);

                // length-encoded auth response
                WriteLenEncInt(ms, (ulong)authResp.Length);
                if (authResp.Length > 0)
                {
                    ms.Write(authResp, 0, authResp.Length);
                }

                if (!string.IsNullOrEmpty(database))
                {
                    byte[] dbBytes = Encoding.UTF8.GetBytes(database);
                    ms.Write(dbBytes, 0, dbBytes.Length);
                    ms.WriteByte(0);
                }

                byte[] pluginBytes = Encoding.UTF8.GetBytes(authPlugin);
                ms.Write(pluginBytes, 0, pluginBytes.Length);
                ms.WriteByte(0);

                WritePacket(ms.ToArray());
            }

            // 4. Process Authentication Loop (OK / AuthMoreData / AuthSwitchRequest / ERR)
            HandleAuthLoop(authPlugin, password ?? "", nonce);
            // Buffer complete command responses after the TLS/authentication handshake.
            _stream = new BufferedStream(_stream, 262144);
        }

        private bool ValidateServerCertificate(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        {
            if (!string.Equals(_sslMode, "VerifyCA", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            if (sslPolicyErrors == SslPolicyErrors.None)
            {
                return true;
            }
            if (!string.IsNullOrEmpty(_caCertPath) && File.Exists(_caCertPath) && certificate != null)
            {
                List<X509Certificate2> extraCerts = new List<X509Certificate2>();
                using (X509Certificate2 serverCert = new X509Certificate2(certificate))
                {
                    X509Chain customChain = new X509Chain();
                    try
                    {
                        customChain.ChainPolicy.RevocationMode = X509RevocationMode.NoCheck;
                        customChain.ChainPolicy.VerificationFlags = X509VerificationFlags.AllowUnknownCertificateAuthority;

                        string pemText = File.ReadAllText(_caCertPath);
                        MatchCollection matches = Regex.Matches(pemText, "-----BEGIN CERTIFICATE-----(.+?)-----END CERTIFICATE-----", RegexOptions.Singleline);
                        foreach (Match m in matches)
                        {
                            string b64 = Regex.Replace(m.Groups[1].Value, "\\s+", "");
                            byte[] certBytes = Convert.FromBase64String(b64);
                            X509Certificate2 caCert = new X509Certificate2(certBytes);
                            extraCerts.Add(caCert);
                            customChain.ChainPolicy.ExtraStore.Add(caCert);
                        }
                        if (customChain.Build(serverCert) && customChain.ChainElements.Count > 0)
                        {
                            // Verify root is in provided CA store
                            X509Certificate2 root = customChain.ChainElements[customChain.ChainElements.Count - 1].Certificate;
                            foreach (X509Certificate2 ca in extraCerts)
                            {
                                if (string.Equals(ca.Thumbprint, root.Thumbprint, StringComparison.OrdinalIgnoreCase))
                                {
                                    return true;
                                }
                            }
                        }
                        return false;
                    }
                    finally
                    {
                        customChain.Reset();
                        foreach (X509Certificate2 ca in extraCerts)
                        {
                            try { ca.Dispose(); } catch { }
                        }
                    }
                }
            }
            return false;
        }

        private void HandleAuthLoop(string currentPlugin, string password, byte[] currentNonce)
        {
            for (int step = 0; step < 8; step++)
            {
                byte[] pkt = ReadPacket();
                if (pkt.Length == 0) throw new IOException("Unexpected empty packet during MySQL auth.");

                byte header = pkt[0];
                if (header == 0x00)
                {
                    // OK_Packet
                    return;
                }
                if (header == 0xFF)
                {
                    throw ParseErrorPacket(pkt);
                }
                if (header == 0xFE)
                {
                    // AuthSwitchRequest
                    int p = 1;
                    string newPlugin = ReadNullTerminatedString(pkt, ref p);
                    int rem = pkt.Length - p;
                    if (rem > 0 && pkt[pkt.Length - 1] == 0) rem--;
                    byte[] newNonce = new byte[Math.Max(0, rem)];
                    if (rem > 0) Array.Copy(pkt, p, newNonce, 0, rem);

                    currentPlugin = newPlugin;
                    currentNonce = newNonce;

                    byte[] switchedResp = ComputeAuthResponse(currentPlugin, password, currentNonce);
                    WritePacket(switchedResp);
                    continue;
                }
                if (header == 0x01)
                {
                    // AuthMoreData
                    if (pkt.Length >= 2 && pkt[1] == 0x03)
                    {
                        // caching_sha2_password fast_auth_success -> next packet is OK_Packet
                        continue;
                    }
                    if (pkt.Length >= 2 && pkt[1] == 0x04)
                    {
                        // caching_sha2_password perform_full_authentication
                        if (_isSsl)
                        {
                            byte[] pwBytes = Encoding.UTF8.GetBytes(password + "\0");
                            WritePacket(pwBytes);
                        }
                        else
                        {
                            // Request RSA public key
                            WritePacket(new byte[] { 0x02 });
                            byte[] keyPkt = ReadPacket();
                            if (keyPkt.Length == 0 || keyPkt[0] == 0xFF)
                            {
                                throw ParseErrorPacket(keyPkt);
                            }
                            string pem = Encoding.ASCII.GetString(keyPkt, 1, keyPkt.Length - 1);
                            byte[] encrypted = EncryptPasswordWithRsaPem(password, currentNonce, pem);
                            WritePacket(encrypted);
                        }
                        continue;
                    }
                    throw new AuthenticationException(string.Format("Unexpected AuthMoreData byte: 0x{0:X2}", pkt.Length > 1 ? pkt[1] : 0));
                }
                throw new AuthenticationException(string.Format("Unexpected packet header during auth: 0x{0:X2}", header));
            }
            throw new AuthenticationException("MySQL authentication exceeded maximum handshake steps.");
        }

        private static byte[] ComputeAuthResponse(string plugin, string password, byte[] nonce)
        {
            if (string.Equals(plugin, "mysql_clear_password", StringComparison.OrdinalIgnoreCase))
            {
                return Encoding.UTF8.GetBytes(password + "\0");
            }
            if (string.IsNullOrEmpty(password))
            {
                return new byte[0];
            }
            if (string.Equals(plugin, "caching_sha2_password", StringComparison.OrdinalIgnoreCase))
            {
                // XOR(SHA256(password), SHA256(SHA256(SHA256(password)) + nonce))
                using (SHA256 sha = SHA256.Create())
                {
                    byte[] h1 = sha.ComputeHash(Encoding.UTF8.GetBytes(password));
                    byte[] h2 = sha.ComputeHash(h1);
                    byte[] combined = new byte[h2.Length + nonce.Length];
                    Array.Copy(h2, 0, combined, 0, h2.Length);
                    Array.Copy(nonce, 0, combined, h2.Length, nonce.Length);
                    byte[] h3 = sha.ComputeHash(combined);
                    byte[] res = new byte[h1.Length];
                    for (int i = 0; i < h1.Length; i++)
                    {
                        res[i] = (byte)(h1[i] ^ h3[i]);
                    }
                    return res;
                }
            }
            // mysql_native_password: XOR(SHA1(password), SHA1(nonce + SHA1(SHA1(password))))
            using (SHA1 sha1 = SHA1.Create())
            {
                byte[] s1 = sha1.ComputeHash(Encoding.UTF8.GetBytes(password));
                byte[] s2 = sha1.ComputeHash(s1);
                byte[] combined = new byte[nonce.Length + s2.Length];
                Array.Copy(nonce, 0, combined, 0, nonce.Length);
                Array.Copy(s2, 0, combined, nonce.Length, s2.Length);
                byte[] s3 = sha1.ComputeHash(combined);
                byte[] res = new byte[s1.Length];
                for (int i = 0; i < s1.Length; i++)
                {
                    res[i] = (byte)(s1[i] ^ s3[i]);
                }
                return res;
            }
        }

        private static byte[] EncryptPasswordWithRsaPem(string password, byte[] nonce, string pem)
        {
            byte[] pw = Encoding.UTF8.GetBytes(password + "\0");
            byte[] xored = new byte[pw.Length];
            for (int i = 0; i < pw.Length; i++)
            {
                xored[i] = (byte)(pw[i] ^ nonce[i % nonce.Length]);
            }

            Match m = Regex.Match(pem, "-----BEGIN PUBLIC KEY-----(.+?)-----END PUBLIC KEY-----", RegexOptions.Singleline);
            if (!m.Success)
            {
                throw new AuthenticationException("Invalid RSA public key PEM from MySQL server.");
            }
            string b64 = Regex.Replace(m.Groups[1].Value, "\\s+", "");
            byte[] der = Convert.FromBase64String(b64);
            RSAParameters rsaParams = ParseSubjectPublicKeyInfo(der);

            using (RSACryptoServiceProvider rsa = new RSACryptoServiceProvider())
            {
                rsa.ImportParameters(rsaParams);
                return rsa.Encrypt(xored, true); // OAEP-SHA1
            }
        }

        // Pure .NET 4.8 ASN.1 DER SubjectPublicKeyInfo -> RSAParameters parser
        private static RSAParameters ParseSubjectPublicKeyInfo(byte[] der)
        {
            int idx = 0;
            ExpectTag(der, ref idx, 0x30); // outer SEQUENCE
            ReadDerLength(der, ref idx);

            ExpectTag(der, ref idx, 0x30); // AlgorithmIdentifier SEQUENCE
            int algLen = ReadDerLength(der, ref idx);
            idx += algLen;

            ExpectTag(der, ref idx, 0x03); // BIT STRING
            ReadDerLength(der, ref idx);
            idx += 1; // unused bits byte (0x00)

            ExpectTag(der, ref idx, 0x30); // RSAPublicKey SEQUENCE
            ReadDerLength(der, ref idx);

            ExpectTag(der, ref idx, 0x02); // Modulus INTEGER
            int modLen = ReadDerLength(der, ref idx);
            byte[] mod = ReadDerUnsignedInt(der, ref idx, modLen);

            ExpectTag(der, ref idx, 0x02); // Exponent INTEGER
            int expLen = ReadDerLength(der, ref idx);
            byte[] exp = ReadDerUnsignedInt(der, ref idx, expLen);

            RSAParameters rp = new RSAParameters();
            rp.Modulus = mod;
            rp.Exponent = exp;
            return rp;
        }

        private static void ExpectTag(byte[] buf, ref int idx, byte tag)
        {
            if (idx >= buf.Length || buf[idx] != tag)
            {
                throw new InvalidDataException(string.Format("Unexpected ASN.1 tag 0x{0:X2}, expected 0x{1:X2}", idx < buf.Length ? buf[idx] : 0, tag));
            }
            idx++;
        }

        private static int ReadDerLength(byte[] buf, ref int idx)
        {
            int b = buf[idx++];
            if ((b & 0x80) == 0) return b;
            int count = b & 0x7F;
            int len = 0;
            for (int i = 0; i < count; i++)
            {
                len = (len << 8) | buf[idx++];
            }
            return len;
        }

        private static byte[] ReadDerUnsignedInt(byte[] buf, ref int idx, int len)
        {
            int start = idx;
            idx += len;
            while (len > 1 && buf[start] == 0x00)
            {
                start++;
                len--;
            }
            byte[] res = new byte[len];
            Array.Copy(buf, start, res, 0, len);
            return res;
        }

        public void ExecuteNonQuery(string sql)
        {
            _seqId = 0;
            byte[] sqlBytes = Encoding.UTF8.GetBytes(sql);
            byte[] payload = new byte[1 + sqlBytes.Length];
            payload[0] = 0x03; // COM_QUERY
            Array.Copy(sqlBytes, 0, payload, 1, sqlBytes.Length);
            WritePacket(payload);

            byte[] resp = ReadPacket();
            if (resp.Length > 0 && resp[0] == 0xFF)
            {
                throw ParseErrorPacket(resp);
            }
        }

        public MySqlQueryResult ExecuteQuery(string sql)
        {
            _seqId = 0;
            byte[] sqlBytes = Encoding.UTF8.GetBytes(sql);
            byte[] payload = new byte[1 + sqlBytes.Length];
            payload[0] = 0x03; // COM_QUERY
            Array.Copy(sqlBytes, 0, payload, 1, sqlBytes.Length);
            WritePacket(payload);

            MySqlQueryResult result = new MySqlQueryResult();
            byte[] first = ReadPacket();
            if (first.Length == 0) return result;
            if (first[0] == 0xFF) throw ParseErrorPacket(first);
            if (first[0] == 0x00) return result; // OK, no rows

            int pos = 0;
            int colCount = (int)ReadLenEncInt(first, ref pos);
            for (int i = 0; i < colCount; i++)
            {
                byte[] colPkt = ReadPacket();
                result.Columns.Add(ParseColumnDefinition(colPkt));
            }

            // Intermediate EOF packet
            byte[] eof1 = ReadPacket();
            if (eof1.Length > 0 && eof1[0] == 0xFF) throw ParseErrorPacket(eof1);

            while (true)
            {
                byte[] rowPkt = ReadPacket();
                if (rowPkt.Length == 0) break;
                if (rowPkt[0] == 0xFF) throw ParseErrorPacket(rowPkt);
                if (rowPkt[0] == 0xFE && rowPkt.Length < 9) break; // EOF

                string[] row = new string[colCount];
                int rp = 0;
                for (int c = 0; c < colCount; c++)
                {
                    if (rp >= rowPkt.Length)
                    {
                        row[c] = null;
                        continue;
                    }
                    if (rowPkt[rp] == 0xFB)
                    {
                        row[c] = null;
                        rp++;
                    }
                    else
                    {
                        ulong slen = ReadLenEncInt(rowPkt, ref rp);
                        row[c] = Encoding.UTF8.GetString(rowPkt, rp, (int)slen);
                        rp += (int)slen;
                    }
                }
                result.Rows.Add(row);
            }
            return result;
        }

        public long ExportQueryToCsv(string sql, string csvFilePath, int maxRows, bool writeUtf8Bom)
        {
            return ExportQueryToCsv(sql, csvFilePath, maxRows, writeUtf8Bom, null, null);
        }

        public long ExportQueryToCsv(
            string sql,
            string csvFilePath,
            int maxRows,
            bool writeUtf8Bom,
            OperationCancelState cancelState,
            Action<long> onRowProgress)
        {
            return ExportQueryToCsv(sql, csvFilePath, maxRows, writeUtf8Bom, cancelState, onRowProgress, null);
        }

        internal long ExportQueryToCsv(string sql, string csvFilePath, int maxRows, bool writeUtf8Bom,
            OperationCancelState cancelState, Action<long> onRowProgress, CsvExportAnalysis analysis)
        {
            OperationCancelState effCancel = cancelState ?? _cancelState;
            if (effCancel != null) effCancel.ThrowIfCancelled();

            _seqId = 0;
            byte[] sqlBytes = Encoding.UTF8.GetBytes(sql);
            byte[] payload = new byte[1 + sqlBytes.Length];
            payload[0] = 0x03; // COM_QUERY
            Array.Copy(sqlBytes, 0, payload, 1, sqlBytes.Length);
            WritePacket(payload);

            byte[] first = ReadPacket();
            if (first.Length == 0) throw new InvalidDataException("Missing CSV result set.");
            if (first[0] == 0xFF) throw ParseErrorPacket(first);
            if (first[0] == 0x00) throw new InvalidDataException("Expected CSV column definitions.");

            int pos = 0;
            int colCount = (int)ReadLenEncInt(first, ref pos);
            List<MySqlColumnInfo> cols = new List<MySqlColumnInfo>(colCount);
            for (int i = 0; i < colCount; i++)
            {
                if (effCancel != null) effCancel.ThrowIfCancelled();
                cols.Add(ParseColumnDefinition(ReadPacket()));
            }

            byte[] eof1 = ReadPacket();
            if (eof1.Length > 0 && eof1[0] == 0xFF) throw ParseErrorPacket(eof1);

            long writtenRows = 0;
            Encoding enc = new UTF8Encoding(writeUtf8Bom);
            string tempPartPath = csvFilePath + ".part";
            bool committed = false;
            Stopwatch progressClock = Stopwatch.StartNew();

            try
            {
                CsvTableProfile profile = analysis == null ? null : new CsvTableProfile(csvFilePath,
                    writeUtf8Bom ? "UTF-8 (BOM)" : "UTF-8", cols.Select(c => c.Name).ToArray());
                string[] values = profile == null ? null : new string[colCount];
                using (FileStream csvOut = new FileStream(tempPartPath, FileMode.Create, FileAccess.Write, FileShare.Read, 1048576, FileOptions.SequentialScan))
                using (StreamWriter sw = new StreamWriter(csvOut, enc, 65536))
                {
                    for (int c = 0; c < colCount; c++)
                    {
                        if (c > 0) sw.Write(',');
                        WriteCsvField(sw, cols[c].Name);
                    }
                    sw.WriteLine();

                    while (true)
                    {
                        if (effCancel != null) effCancel.ThrowIfCancelled();

                        int pktLen = ReadBulkPacket();
                        if (pktLen <= 0) throw new InvalidDataException("Missing CSV row data.");
                        if (_bulkPacket[0] == 0xFF) throw ParseErrorPacket(_bulkPacket, pktLen);
                        if (_bulkPacket[0] == 0xFE && pktLen < 9) break;

                        if (maxRows > 0 && writtenRows >= maxRows)
                        {
                            // Drain remaining packets cleanly to keep protocol state synchronized
                            continue;
                        }

                        int rp = 0;
                        for (int c = 0; c < colCount; c++)
                        {
                            if (c > 0) sw.Write(',');
                            string value = ReadTextField(_bulkPacket, pktLen, ref rp);
                            WriteCsvField(sw, value);
                            if (values != null) values[c] = value;
                        }
                        sw.WriteLine();
                        if (profile != null) profile.AddRow(values);
                        writtenRows++;
                        if (onRowProgress != null && (writtenRows == 1 || progressClock.ElapsedMilliseconds >= 400))
                        {
                            onRowProgress(writtenRows);
                            progressClock.Restart();
                        }
                    }
                }

                if (effCancel != null) effCancel.ThrowIfCancelled();
                if (profile != null) profile.Complete(new FileInfo(tempPartPath).Length);
                if (effCancel != null) effCancel.ThrowIfCancelled();
                AtomicFile.Publish(tempPartPath, csvFilePath);
                committed = true;
                if (profile != null) analysis.Profiles.Add(csvFilePath, profile);
            }
            finally
            {
                if (!committed)
                {
                    try { if (File.Exists(tempPartPath)) File.Delete(tempPartPath); } catch { }
                }
            }
            return writtenRows;
        }

        public long StreamQuery(
            string sql,
            int maxRows,
            Action<List<MySqlColumnInfo>> onColumns,
            Action<string[], long, long> onRow)
        {
            return StreamQuery(sql, maxRows, onColumns, onRow, null, null);
        }

        public long StreamQuery(
            string sql,
            int maxRows,
            Action<List<MySqlColumnInfo>> onColumns,
            Action<string[], long, long> onRow,
            OperationCancelState cancelState,
            Action<long> onRowProgress)
        {
            OperationCancelState effCancel = cancelState ?? _cancelState;
            if (effCancel != null) effCancel.ThrowIfCancelled();

            _seqId = 0;
            byte[] sqlBytes = Encoding.UTF8.GetBytes(sql);
            byte[] payload = new byte[1 + sqlBytes.Length];
            payload[0] = 0x03; // COM_QUERY
            Array.Copy(sqlBytes, 0, payload, 1, sqlBytes.Length);
            WritePacket(payload);

            byte[] first = ReadPacket();
            if (first.Length == 0) throw new InvalidDataException("Missing analysis result set.");
            if (first[0] == 0xFF) throw ParseErrorPacket(first);
            if (first[0] == 0x00) throw new InvalidDataException("Expected analysis column definitions.");

            int pos = 0;
            int colCount = (int)ReadLenEncInt(first, ref pos);
            List<MySqlColumnInfo> cols = new List<MySqlColumnInfo>(colCount);
            for (int i = 0; i < colCount; i++)
            {
                if (effCancel != null) effCancel.ThrowIfCancelled();
                cols.Add(ParseColumnDefinition(ReadPacket()));
            }

            byte[] eof1 = ReadPacket();
            if (eof1.Length > 0 && eof1[0] == 0xFF) throw ParseErrorPacket(eof1);

            if (onColumns != null)
            {
                onColumns(cols);
            }

            long rowCount = 0;
            string[] rowBuf = new string[colCount];
            long baseCsvOverhead = (colCount > 1 ? (long)(colCount - 1) : 0L) + 2L; // commas + CRLF

            while (true)
            {
                if (effCancel != null) effCancel.ThrowIfCancelled();

                int pktLen = ReadBulkPacket();
                if (pktLen <= 0) throw new InvalidDataException("Missing analysis row data.");
                if (_bulkPacket[0] == 0xFF) throw ParseErrorPacket(_bulkPacket, pktLen);
                if (_bulkPacket[0] == 0xFE && pktLen < 9) break;

                if (maxRows > 0 && rowCount >= maxRows)
                {
                    // Drain remaining packets cleanly to keep protocol state synchronized
                    continue;
                }

                long rawRowBytes = 0;
                long csvLineBytes = baseCsvOverhead;
                int rp = 0;
                for (int c = 0; c < colCount; c++)
                {
                    string val = ReadTextField(_bulkPacket, pktLen, ref rp);
                    rowBuf[c] = val;
                    if (!string.IsNullOrEmpty(val))
                    {
                        // Binary columns can decode to replacement characters. Measure the
                        // text that the CSV writer actually produces, including escaping.
                        int textBytes = ContentHashHelper.Utf8ByteCount(val);
                        rawRowBytes += textBytes;
                        csvLineBytes += textBytes + CsvEscapeOverhead(val);
                    }
                }

                if (onRow != null)
                {
                    onRow(rowBuf, rawRowBytes, csvLineBytes);
                }
                rowCount++;
                if (onRowProgress != null && (rowCount % 5000) == 0)
                {
                    onRowProgress(rowCount);
                }
            }

            return rowCount;
        }

        private int ReadBulkPacket()
        {
            int total = 0;
            while (true)
            {
                ReadExact(_packetHeader, 0, 4);
                int len = _packetHeader[0] | (_packetHeader[1] << 8) | (_packetHeader[2] << 16);
                _seqId = (byte)(_packetHeader[3] + 1);
                int need = total + len;
                if (need > _bulkPacket.Length)
                {
                    int cap = _bulkPacket.Length;
                    while (cap < need) cap = cap < 1048576 ? cap * 2 : cap + 1048576;
                    byte[] grown = new byte[cap];
                    if (total > 0) Buffer.BlockCopy(_bulkPacket, 0, grown, 0, total);
                    _bulkPacket = grown;
                }
                if (len > 0) ReadExact(_bulkPacket, total, len);
                total += len;
                if (len < 0xFFFFFF) return total;
            }
        }

        private static string ReadTextField(byte[] packet, int packetLen, ref int offset)
        {
            if (offset >= packetLen) throw new InvalidDataException("Missing MySQL row field.");
            if (packet[offset] == 0xFB)
            {
                offset++;
                return null;
            }
            ulong len = ReadLenEncInt(packet, ref offset);
            if (len > (ulong)(packetLen - offset)) throw new InvalidDataException("Truncated MySQL row field.");
            if (len == 0) return "";
            string val = Encoding.UTF8.GetString(packet, offset, (int)len);
            offset += (int)len;
            return val;
        }

        private static readonly char[] CsvSpecialChars = { ',', '"', '\r', '\n' };

        private static long CsvEscapeOverhead(string val)
        {
            int first = val.IndexOfAny(CsvSpecialChars);
            if (first < 0) return 0;
            long extra = 2;
            for (int i = first; i < val.Length; i++) if (val[i] == '"') extra++;
            return extra;
        }

        internal static long CalcCsvEncodedUtf8Length(string val)
        {
            return ContentHashHelper.Utf8ByteCount(val) + CsvEscapeOverhead(val);
        }

        private static void WriteCsvField(StreamWriter writer, string val)
        {
            if (string.IsNullOrEmpty(val)) return;
            if (val.IndexOfAny(CsvSpecialChars) < 0)
            {
                writer.Write(val);
            }
            else
            {
                writer.Write('"');
                writer.Write(val.Replace("\"", "\"\""));
                writer.Write('"');
            }
        }

        private static MySqlColumnInfo ParseColumnDefinition(byte[] pkt)
        {
            int p = 0;
            SkipLenEncString(pkt, ref p); // catalog
            SkipLenEncString(pkt, ref p); // schema
            SkipLenEncString(pkt, ref p); // table
            SkipLenEncString(pkt, ref p); // org_table
            string name = ReadLenEncString(pkt, ref p); // name
            SkipLenEncString(pkt, ref p); // org_name

            MySqlColumnInfo info = new MySqlColumnInfo();
            info.Name = name;
            if (p < pkt.Length)
            {
                p++; // fixed length fields marker (0x0C)
                if (p + 10 <= pkt.Length)
                {
                    info.Charset = (ushort)(pkt[p] | (pkt[p + 1] << 8));
                    p += 2;
                    info.ColumnLength = (uint)(pkt[p] | (pkt[p + 1] << 8) | (pkt[p + 2] << 16) | (pkt[p + 3] << 24));
                    p += 4;
                    info.ColumnType = pkt[p++];
                    info.Flags = (ushort)(pkt[p] | (pkt[p + 1] << 8));
                }
            }
            return info;
        }

        private byte[] ReadPacket()
        {
            MemoryStream full = null;
            byte[] header = _packetHeader;
            while (true)
            {
                ReadExact(header, 0, 4);
                int len = header[0] | (header[1] << 8) | (header[2] << 16);
                _seqId = (byte)(header[3] + 1);

                byte[] chunk = new byte[len];
                if (len > 0)
                {
                    ReadExact(chunk, 0, len);
                }

                if (len < 0xFFFFFF)
                {
                    if (full == null) return chunk;
                    full.Write(chunk, 0, len);
                    return full.ToArray();
                }
                if (full == null) full = new MemoryStream();
                full.Write(chunk, 0, len);
            }
        }

        private void WritePacket(byte[] payload)
        {
            if (_cancelState != null) _cancelState.ThrowIfCancelled();
            int offset = 0;
            try
            {
                do
                {
                    int chunkLen = Math.Min(0xFFFFFF, payload.Length - offset);
                    byte[] hdr = new byte[4];
                    hdr[0] = (byte)(chunkLen & 0xFF);
                    hdr[1] = (byte)((chunkLen >> 8) & 0xFF);
                    hdr[2] = (byte)((chunkLen >> 16) & 0xFF);
                    hdr[3] = _seqId++;
                    _stream.Write(hdr, 0, 4);
                    if (chunkLen > 0)
                    {
                        _stream.Write(payload, offset, chunkLen);
                    }
                    offset += chunkLen;
                } while (offset < payload.Length);
                _stream.Flush();
            }
            catch
            {
                if (_cancelState != null && _cancelState.IsCancelled) _cancelState.ThrowIfCancelled();
                throw;
            }
        }

        private void ReadExact(byte[] buffer, int offset, int count)
        {
            int total = 0;
            while (total < count)
            {
                if (_cancelState != null) _cancelState.ThrowIfCancelled();
                int r;
                try
                {
                    r = _stream.Read(buffer, offset + total, count - total);
                }
                catch
                {
                    if (_cancelState != null && _cancelState.IsCancelled) _cancelState.ThrowIfCancelled();
                    throw;
                }
                if (r <= 0)
                {
                    if (_cancelState != null && _cancelState.IsCancelled) _cancelState.ThrowIfCancelled();
                    throw new EndOfStreamException("Connection closed unexpectedly by MySQL server.");
                }
                total += r;
            }
        }

        private static Exception ParseErrorPacket(byte[] pkt)
        {
            return ParseErrorPacket(pkt, pkt == null ? 0 : pkt.Length);
        }

        private static Exception ParseErrorPacket(byte[] pkt, int length)
        {
            if (pkt == null || length < 3) return new Exception("Unknown MySQL error.");
            int errCode = pkt[1] | (pkt[2] << 8);
            int p = 3;
            string sqlState = "";
            if (p < length && pkt[p] == (byte)'#')
            {
                p++;
                if (p + 5 <= length)
                {
                    sqlState = Encoding.ASCII.GetString(pkt, p, 5);
                    p += 5;
                }
            }
            string msg = (p < length) ? Encoding.UTF8.GetString(pkt, p, length - p) : "";
            return new Exception(string.Format("MySQL Error {0} ({1}): {2}", errCode, sqlState, msg));
        }

        private static string ReadNullTerminatedString(byte[] buf, ref int pos)
        {
            int start = pos;
            while (pos < buf.Length && buf[pos] != 0) pos++;
            string s = Encoding.UTF8.GetString(buf, start, pos - start);
            if (pos < buf.Length) pos++;
            return s;
        }

        private static ulong ReadLenEncInt(byte[] buf, ref int pos)
        {
            byte b = buf[pos++];
            if (b < 0xFB) return b;
            if (b == 0xFC)
            {
                ulong v = (ulong)(buf[pos] | (buf[pos + 1] << 8));
                pos += 2;
                return v;
            }
            if (b == 0xFD)
            {
                ulong v = (ulong)(buf[pos] | (buf[pos + 1] << 8) | (buf[pos + 2] << 16));
                pos += 3;
                return v;
            }
            if (b == 0xFE)
            {
                ulong v = BitConverter.ToUInt64(buf, pos);
                pos += 8;
                return v;
            }
            return 0;
        }

        private static void WriteLenEncInt(Stream s, ulong val)
        {
            if (val < 251) { s.WriteByte((byte)val); }
            else if (val < 65536) { s.WriteByte(0xFC); s.WriteByte((byte)(val & 0xFF)); s.WriteByte((byte)((val >> 8) & 0xFF)); }
            else if (val < 16777216) { s.WriteByte(0xFD); s.WriteByte((byte)(val & 0xFF)); s.WriteByte((byte)((val >> 8) & 0xFF)); s.WriteByte((byte)((val >> 16) & 0xFF)); }
            else { s.WriteByte(0xFE); byte[] b = BitConverter.GetBytes(val); s.Write(b, 0, 8); }
        }

        private static void WriteUInt32LE(Stream s, uint v)
        {
            s.WriteByte((byte)(v & 0xFF));
            s.WriteByte((byte)((v >> 8) & 0xFF));
            s.WriteByte((byte)((v >> 16) & 0xFF));
            s.WriteByte((byte)((v >> 24) & 0xFF));
        }

        private static string ReadLenEncString(byte[] buf, ref int pos)
        {
            ulong len = ReadLenEncInt(buf, ref pos);
            string s = Encoding.UTF8.GetString(buf, pos, (int)len);
            pos += (int)len;
            return s;
        }

        private static void SkipLenEncString(byte[] buf, ref int pos)
        {
            ulong len = ReadLenEncInt(buf, ref pos);
            pos += (int)len;
        }

        public void Dispose()
        {
            IDisposable reg = _cancelReg;
            _cancelReg = null;
            if (reg != null)
            {
                try { reg.Dispose(); } catch { }
            }
            try
            {
                if (_stream != null && !_aborted && (_cancelState == null || !_cancelState.IsCancelled))
                {
                    _seqId = 0;
                    WritePacket(new byte[] { 0x01 }); // COM_QUIT
                }
            }
            catch { }
            try { if (_stream != null) _stream.Dispose(); } catch { }
            try { if (_tcp != null) _tcp.Close(); } catch { }
            _stream = null;
            _tcp = null;
            _cancelState = null;
        }
    }

    // =========================================================================
    //  Performance-Testing CSV Statistical & Distribution Analyzer
    // =========================================================================
    public struct RowSketchEntry
    {
        public ulong PkHash;
        public ulong RowHash;
    }

    public class RowChurnSketch
    {
        public string KeyColumns;
        public bool IsValid;
        public bool IsExact;
        public long TotalRows;
        public ulong MaxKeptPkHash;
        public List<RowSketchEntry> Entries;

        public RowChurnSketch()
        {
            KeyColumns = "";
            Entries = new List<RowSketchEntry>();
        }
    }

    public class DbColumnMetadata
    {
        public string ColumnName;
        public int OrdinalPosition;
        public string ColumnType;
        public string IsNullable;
        public string ColumnKey;
        public string ColumnDefault;
        public string Extra;
        public string CharacterSetName;
        public string CollationName;
        public string ColumnComment;

        public DbColumnMetadata()
        {
            ColumnName = "";
            ColumnType = "";
            IsNullable = "";
            ColumnKey = "";
            ColumnDefault = "";
            Extra = "";
            CharacterSetName = "";
            CollationName = "";
            ColumnComment = "";
        }
    }

    public class DbIndexMetadata
    {
        public string IndexName;
        public bool NonUnique;
        public string IndexType;
        public string Columns;
        public long Cardinality;

        public DbIndexMetadata()
        {
            IndexName = "";
            IndexType = "";
            Columns = "";
        }
    }

    public class DbForeignKeyMetadata
    {
        public string ConstraintName;
        public string ColumnName;
        public string ReferencedTableName;
        public string ReferencedColumnName;

        public DbForeignKeyMetadata()
        {
            ConstraintName = "";
            ColumnName = "";
            ReferencedTableName = "";
            ReferencedColumnName = "";
        }
    }

    public class DbTableMetadata
    {
        public bool HasMetadata;
        public string Engine;
        public string RowFormat;
        public long EstTableRows;
        public long AvgRowLength;
        public long DataLength;
        public long IndexLength;
        public long DataFree;
        public bool HasAutoIncrement;
        public long AutoIncrement;
        public string CreateTime;
        public string UpdateTime;
        public string TableCollation;
        public string TableComment;
        public List<DbIndexMetadata> Indexes;
        public List<DbForeignKeyMetadata> DeclaredFks;

        public DbTableMetadata()
        {
            Engine = "";
            RowFormat = "";
            CreateTime = "";
            UpdateTime = "";
            TableCollation = "";
            TableComment = "";
            Indexes = new List<DbIndexMetadata>();
            DeclaredFks = new List<DbForeignKeyMetadata>();
        }
    }

    public class DbServerMetadata
    {
        public bool HasVariables;
        public bool HasStatus;
        public Dictionary<string, string> Variables;
        public Dictionary<string, long> StatusCounters;

        public DbServerMetadata()
        {
            Variables = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            StatusCounters = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        }
    }

    internal static class ContentHashHelper
    {
        internal static bool IsNullToken(string raw)
        {
            return raw == null || raw.Length == 0 || raw == "\\N" ||
                   string.Equals(raw, "NULL", StringComparison.Ordinal) ||
                   string.Equals(raw, "(null)", StringComparison.OrdinalIgnoreCase);
        }

        internal static ulong Mix64(ulong z)
        {
            unchecked
            {
                z = (z ^ (z >> 30)) * 0xBF58476D1CE4E5B9UL;
                z = (z ^ (z >> 27)) * 0x94D049BB133111EBUL;
                return z ^ (z >> 31);
            }
        }

        internal static void HashString128(string s, out ulong h1, out ulong h2)
        {
            int ignored;
            HashString128Utf8(s, out h1, out h2, out ignored);
        }

        internal static int Utf8ByteCount(string s)
        {
            if (string.IsNullOrEmpty(s)) return 0;
            int bytes = 0;
            bool fallback = false;
            bool pair = false;
            for (int i = 0; i < s.Length; i++)
            {
                char ch = s[i];
                if (pair)
                {
                    pair = false;
                    if (!char.IsLowSurrogate(ch)) fallback = true;
                    continue;
                }
                if (ch <= 0x7F) bytes++;
                else if (ch <= 0x7FF) bytes += 2;
                else if (char.IsHighSurrogate(ch))
                {
                    if (i + 1 < s.Length && char.IsLowSurrogate(s[i + 1]))
                    {
                        bytes += 4;
                        pair = true;
                    }
                    else fallback = true;
                }
                else if (char.IsLowSurrogate(ch)) fallback = true;
                else bytes += 3;
                if (fallback) return Encoding.UTF8.GetByteCount(s);
            }
            return bytes;
        }

        // Hash matches the previous character walk, including the CRLF normalization.
        // byteLen matches Encoding.UTF8.GetByteCount for the original text.
        internal static void HashString128Utf8(string s, out ulong h1, out ulong h2, out int byteLen)
        {
            unchecked
            {
                ulong a = 0xCBF29CE484222325UL;
                ulong b = 0x9E3779B97F4A7C15UL;
                int len = s != null ? s.Length : 0;
                int normLen = 0;
                int bytes = 0;
                bool fallback = false;
                bool pair = false;
                for (int i = 0; i < len; i++)
                {
                    char ch = s[i];
                    if (ch == '\r' && i + 1 < len && s[i + 1] == '\n')
                    {
                        if (!fallback) bytes++;
                        continue;
                    }
                    a = (a ^ ch) * 0x100000001B3UL;
                    b = (b * 0x517CC1B727220A95UL) ^ ((ulong)ch + 0x9E3779B97F4A7C15UL);
                    normLen++;
                    if (fallback) continue;
                    if (pair)
                    {
                        pair = false;
                        if (!char.IsLowSurrogate(ch)) fallback = true;
                    }
                    else if (ch <= 0x7F) bytes++;
                    else if (ch <= 0x7FF) bytes += 2;
                    else if (char.IsHighSurrogate(ch))
                    {
                        if (i + 1 < len && char.IsLowSurrogate(s[i + 1]))
                        {
                            bytes += 4;
                            pair = true;
                        }
                        else fallback = true;
                    }
                    else if (char.IsLowSurrogate(ch)) fallback = true;
                    else bytes += 3;
                }
                h1 = Mix64(a ^ ((ulong)normLen * 0x9E3779B185EBCA87UL));
                h2 = Mix64(b ^ ((ulong)normLen * 0xC2B2AE3D27D4EB4FUL) ^ (h1 >> 17));
                byteLen = fallback ? Encoding.UTF8.GetByteCount(s ?? "") : bytes;
            }
        }

        internal static void BeginRowHash(int colCount, out ulong a, out ulong b)
        {
            unchecked
            {
                a = 0x84222325CBF29CE4UL ^ ((ulong)colCount * 0x9E3779B97F4A7C15UL);
                b = 0x6A09E667F3BCC908UL ^ ((ulong)colCount * 0xBF58476D1CE4E5B9UL);
            }
        }

        internal static void NullColumnHash(int index, out ulong h1, out ulong h2)
        {
            h1 = Mix64(0x1111111111111111UL ^ (ulong)(index + 1));
            h2 = Mix64(0x2222222222222222UL ^ (ulong)(index + 1));
        }

        internal static void MixRowColumn(ref ulong a, ref ulong b, int index, ulong ch1, ulong ch2)
        {
            unchecked
            {
                a = Mix64(a ^ (ch1 + (ulong)(index + 1) * 0x9E3779B97F4A7C15UL));
                b = Mix64(b ^ (ch2 + (ulong)(index + 1) * 0x517CC1B727220A95UL));
            }
        }

        internal static void ComputeRowHash128(IList<string> values, int colCount, out ulong h1, out ulong h2)
        {
            ulong a, b;
            BeginRowHash(colCount, out a, out b);
            int vCount = values != null ? values.Count : 0;
            for (int i = 0; i < colCount; i++)
            {
                string raw = i < vCount ? values[i] : null;
                ulong ch1, ch2;
                if (IsNullToken(raw)) NullColumnHash(i, out ch1, out ch2);
                else HashString128(raw, out ch1, out ch2);
                MixRowColumn(ref a, ref b, i, ch1, ch2);
            }
            h1 = a;
            h2 = b;
        }

        internal static ulong ComputeKeyHash64(IList<string> values, int[] keyCols, out bool isNullKey)
        {
            isNullKey = false;
            if (values == null || keyCols == null || keyCols.Length == 0)
            {
                isNullKey = true;
                return 0UL;
            }
            unchecked
            {
                ulong acc = 0xA0761D6478BD642FUL ^ ((ulong)keyCols.Length * 0x9E3779B97F4A7C15UL);
                for (int k = 0; k < keyCols.Length; k++)
                {
                    int ci = keyCols[k];
                    string raw = (ci >= 0 && ci < values.Count) ? values[ci] : null;
                    if (IsNullToken(raw))
                    {
                        isNullKey = true;
                        return 0UL;
                    }
                    ulong h1, h2;
                    HashString128(raw, out h1, out h2);
                    acc = Mix64(acc ^ h1 ^ (h2 * 0x9E3779B185EBCA87UL) ^ ((ulong)(k + 1) * 0xBF58476D1CE4E5B9UL));
                }
                return acc;
            }
        }

        internal static string Format256(ulong s1, ulong s2, ulong x1, ulong x2)
        {
            return s1.ToString("x16", CultureInfo.InvariantCulture) +
                   s2.ToString("x16", CultureInfo.InvariantCulture) +
                   x1.ToString("x16", CultureInfo.InvariantCulture) +
                   x2.ToString("x16", CultureInfo.InvariantCulture);
        }

        internal static string Format128(ulong s1, ulong s2)
        {
            return s1.ToString("x16", CultureInfo.InvariantCulture) +
                   s2.ToString("x16", CultureInfo.InvariantCulture);
        }

        internal static string EncodeSketchEntries(List<RowSketchEntry> entries, int maxCount)
        {
            if (entries == null || entries.Count == 0 || maxCount <= 0) return "";
            int n = Math.Min(entries.Count, maxCount);
            byte[] buf = new byte[n * 16];
            for (int i = 0; i < n; i++)
            {
                ulong pk = entries[i].PkHash;
                ulong rh = entries[i].RowHash;
                int o = i * 16;
                for (int b = 0; b < 8; b++)
                {
                    buf[o + b] = (byte)(pk >> (b * 8));
                    buf[o + 8 + b] = (byte)(rh >> (b * 8));
                }
            }
            return Convert.ToBase64String(buf);
        }

        internal static List<RowSketchEntry> DecodeSketchEntries(string base64)
        {
            List<RowSketchEntry> list = new List<RowSketchEntry>();
            if (string.IsNullOrWhiteSpace(base64)) return list;
            try
            {
                byte[] buf = Convert.FromBase64String(base64.Trim());
                int n = buf.Length / 16;
                list.Capacity = n;
                for (int i = 0; i < n; i++)
                {
                    int o = i * 16;
                    ulong pk = 0UL;
                    ulong rh = 0UL;
                    for (int b = 0; b < 8; b++)
                    {
                        pk |= ((ulong)buf[o + b]) << (b * 8);
                        rh |= ((ulong)buf[o + 8 + b]) << (b * 8);
                    }
                    RowSketchEntry entry;
                    entry.PkHash = pk;
                    entry.RowHash = rh;
                    list.Add(entry);
                }
            }
            catch { list.Clear(); }
            return list;
        }

        internal static int[] SelectSketchKeyColumns(IList<string> headers, IList<string> explicitPkCols, out string keyLabel)
        {
            keyLabel = "";
            if (headers == null || headers.Count == 0) return null;

            if (explicitPkCols != null && explicitPkCols.Count > 0)
            {
                List<int> idxs = new List<int>(explicitPkCols.Count);
                List<string> names = new List<string>(explicitPkCols.Count);
                bool allFound = true;
                for (int k = 0; k < explicitPkCols.Count; k++)
                {
                    int found = -1;
                    for (int h = 0; h < headers.Count; h++)
                    {
                        if (string.Equals(headers[h], explicitPkCols[k], StringComparison.OrdinalIgnoreCase))
                        {
                            found = h;
                            break;
                        }
                    }
                    if (found < 0) { allFound = false; break; }
                    idxs.Add(found);
                    names.Add(headers[found]);
                }
                if (allFound && idxs.Count > 0)
                {
                    keyLabel = string.Join("+", names.ToArray());
                    return idxs.ToArray();
                }
            }

            // Heuristic fallback (matches CSV & DirectDB identically when no explicit PK metadata)
            for (int h = 0; h < headers.Count; h++)
            {
                if (string.Equals(headers[h], "id", StringComparison.OrdinalIgnoreCase))
                {
                    keyLabel = headers[h];
                    return new int[] { h };
                }
            }

            string h0 = headers[0] ?? "";
            bool h0LooksPk = h0.EndsWith("_id", StringComparison.OrdinalIgnoreCase) ||
                             h0.EndsWith("_cd", StringComparison.OrdinalIgnoreCase) ||
                             h0.EndsWith("_code", StringComparison.OrdinalIgnoreCase) ||
                             h0.EndsWith("_no", StringComparison.OrdinalIgnoreCase) ||
                             h0.EndsWith("_key", StringComparison.OrdinalIgnoreCase) ||
                             h0.EndsWith("_uuid", StringComparison.OrdinalIgnoreCase) ||
                             headers.Count == 1;
            if (h0LooksPk)
            {
                keyLabel = h0;
                return new int[] { 0 };
            }
            return null;
        }
    }

    internal sealed class RowChurnSketchBuilder
    {
        internal const int DefaultCap = 2048;
        private readonly int[] _keyCols;
        private readonly string _keyLabel;
        private readonly bool _explicitPk;
        private readonly int _cap;
        private readonly RowSketchEntry[] _heap;
        private int _count;
        private long _totalRows;
        private bool _invalidKey;

        internal RowChurnSketchBuilder(int[] keyCols, string keyLabel, bool explicitPk, int cap)
        {
            _keyCols = keyCols;
            _keyLabel = keyLabel ?? "";
            _explicitPk = explicitPk;
            _cap = cap > 0 ? cap : DefaultCap;
            _heap = (keyCols != null && keyCols.Length > 0) ? new RowSketchEntry[_cap] : null;
        }

        internal bool IsExplicitPk { get { return _explicitPk; } }
        internal int[] KeyCols { get { return _keyCols; } }

        internal void Add(IList<string> rowValues, ulong rowHash)
        {
            _totalRows++;
            if (_invalidKey || _heap == null) return;
            bool isNullKey;
            ulong pkHash = ContentHashHelper.ComputeKeyHash64(rowValues, _keyCols, out isNullKey);
            if (isNullKey)
            {
                _invalidKey = true;
                return;
            }
            Insert(pkHash, rowHash);
        }

        internal void AddFromHashes(ulong[] hash1, ulong[] hash2, bool[] hashed, ulong rowHash)
        {
            _totalRows++;
            if (_invalidKey || _heap == null) return;
            int keyCount = _keyCols.Length;
            int nHash = hashed == null ? 0 : hashed.Length;
            ulong acc = 0xA0761D6478BD642FUL ^ ((ulong)keyCount * 0x9E3779B97F4A7C15UL);
            for (int k = 0; k < keyCount; k++)
            {
                int ci = _keyCols[k];
                if (ci < 0 || ci >= nHash || !hashed[ci])
                {
                    _invalidKey = true;
                    return;
                }
                ulong h1 = hash1[ci];
                ulong h2 = hash2[ci];
                acc = ContentHashHelper.Mix64(acc ^ h1 ^ (h2 * 0x9E3779B185EBCA87UL) ^ ((ulong)(k + 1) * 0xBF58476D1CE4E5B9UL));
            }
            Insert(acc, rowHash);
        }

        private void Insert(ulong pkHash, ulong rowHash)
        {
            if (_count < _cap)
            {
                RowSketchEntry entry;
                entry.PkHash = pkHash;
                entry.RowHash = rowHash;
                _heap[_count] = entry;
                SiftUp(_count);
                _count++;
            }
            else if (pkHash < _heap[0].PkHash)
            {
                RowSketchEntry entry;
                entry.PkHash = pkHash;
                entry.RowHash = rowHash;
                _heap[0] = entry;
                SiftDown(0);
            }
            else if (pkHash == _heap[0].PkHash)
            {
                _invalidKey = true;
            }
        }

        private void SiftUp(int idx)
        {
            while (idx > 0)
            {
                int parent = (idx - 1) >> 1;
                if (_heap[parent].PkHash >= _heap[idx].PkHash) break;
                RowSketchEntry tmp = _heap[parent];
                _heap[parent] = _heap[idx];
                _heap[idx] = tmp;
                idx = parent;
            }
        }

        private void SiftDown(int idx)
        {
            while (true)
            {
                int left = (idx << 1) + 1;
                if (left >= _count) break;
                int largest = left;
                int right = left + 1;
                if (right < _count && _heap[right].PkHash > _heap[left].PkHash) largest = right;
                if (_heap[idx].PkHash >= _heap[largest].PkHash) break;
                RowSketchEntry tmp = _heap[idx];
                _heap[idx] = _heap[largest];
                _heap[largest] = tmp;
                idx = largest;
            }
        }

        internal RowChurnSketch Build(TableAnalysis finalizedTable)
        {
            RowChurnSketch sk = new RowChurnSketch();
            sk.KeyColumns = _keyLabel;
            sk.TotalRows = _totalRows;
            if (_invalidKey || _heap == null || _totalRows == 0 || _count == 0)
            {
                sk.IsValid = false;
                return sk;
            }

            // For single-column heuristic key, verify uniqueness on full table when not capped
            if (!_explicitPk && _keyCols.Length == 1 && finalizedTable != null && _keyCols[0] < finalizedTable.Columns.Count)
            {
                ColumnAnalysis kc = finalizedTable.Columns[_keyCols[0]];
                if (kc.NullCount > 0 || (!kc.DistinctCapped && kc.DistinctCount < _totalRows))
                {
                    sk.IsValid = false;
                    return sk;
                }
            }

            List<RowSketchEntry> list = new List<RowSketchEntry>(_count);
            for (int i = 0; i < _count; i++) list.Add(_heap[i]);
            list.Sort(delegate(RowSketchEntry a, RowSketchEntry b) { return a.PkHash.CompareTo(b.PkHash); });
            for (int i = 1; i < list.Count; i++)
            {
                if (list[i].PkHash == list[i - 1].PkHash)
                {
                    sk.IsValid = false;
                    return sk;
                }
            }
            sk.IsValid = true;
            sk.IsExact = (_totalRows <= _cap);
            sk.MaxKeptPkHash = sk.IsExact ? ulong.MaxValue : list[list.Count - 1].PkHash;
            sk.Entries = list;
            return sk;
        }
    }

    public class ValueShareItem
    {
        public string Value;
        public long Count;
        public double Share;
        public double CumulativeShare;
    }

    public class HistogramBin
    {
        public int BinIndex;
        public double LowerBound;
        public double UpperBound;
        public string Label;
        public long Count;
        public double Share;
        public double CumulativeShare;
    }

    public class ColumnAnalysis
    {
        public string TableName;
        public string ColumnName;
        public int ColumnIndex;
        public string InferredType;
        public long TotalRows;
        public long NonNullCount;
        public long NullCount;
        public double NullRate;
        public long DistinctCount;
        public bool DistinctCapped;
        public double CardinalityRate;
        public string SelectivityRating;
        public bool IsCandidatePk;

        public int MinCharLen;
        public double AvgCharLen;
        public int MaxCharLen;
        public int MinByteLen;
        public double AvgByteLen;
        public int P95ByteLen;
        public int MaxByteLen;

        public bool IsNumeric;
        public double NumMin;
        public double NumP05;
        public double NumP10;
        public double NumP25;
        public double NumMedian;
        public double NumMean;
        public double NumP75;
        public double NumP90;
        public double NumP95;
        public double NumP99;
        public double NumMax;
        public double NumStdDev;
        public double NumSum;
        public long ZeroCount;
        public double ZeroRate;
        public long NegativeCount;
        public double NegativeRate;
        public long OutlierCount;
        public double OutlierRate;
        public List<HistogramBin> Histogram;

        public string Top1Value;
        public long Top1Count;
        public double Top1Share;
        public double Top3Share;
        public double Top5Share;
        public double Top10Share;
        public double ConcentrationRatio;
        public double GiniCoefficient;
        public double NormalizedEntropy;
        public string SkewRisk;
        public bool IsCategorical;
        public List<ValueShareItem> TopValues;

        public long WhitespaceOnlyCount;
        public long MultilineCount;
        public long NonAsciiCount;
        public double NonAsciiRate;
        public string KeyRole;
        public string MinValueText;
        public string MaxValueText;
        public string TopValuesSummary;
        public string DetailStatsSummary;
        public string RelationSummary;

        public bool IsDateTime;
        public bool HasTimeComponent;
        public DateTime? MinDate;
        public DateTime? MaxDate;
        public double DateSpanDays;
        public double AvgPerDay;
        public string PeakDay;
        public long PeakDayCount;
        public double PeakBurstRatio;
        public SortedDictionary<string, long> MonthlyCounts;
        public SortedDictionary<string, long> RecentDailyCounts;
        public long[] DayOfWeekCounts; // 0=Mon .. 6=Sun
        public long[] HourlyCounts;    // 0..23

        public string ContentHash;
        public string DistinctSetHash;
        public DbColumnMetadata DbMeta;

        public string AlertNote;

        public ColumnAnalysis()
        {
            Histogram = new List<HistogramBin>();
            TopValues = new List<ValueShareItem>();
            MonthlyCounts = new SortedDictionary<string, long>(StringComparer.Ordinal);
            RecentDailyCounts = new SortedDictionary<string, long>(StringComparer.Ordinal);
            DayOfWeekCounts = new long[7];
            HourlyCounts = new long[24];
            ContentHash = "";
            DistinctSetHash = "";
            AlertNote = "";
            Top1Value = "";
            PeakDay = "";
            KeyRole = "一般";
            MinValueText = "-";
            MaxValueText = "-";
            TopValuesSummary = "-";
            DetailStatsSummary = "-";
            RelationSummary = "-";
        }
    }

    public class TableAnalysis
    {
        public string FileName;
        public string FilePath;
        public string TableName;
        public string EncodingName;
        public long FileSizeBytes;
        public long RowCount;
        public int ColumnCount;
        public long RaggedRowCount;
        public long MinRowBytes;
        public double AvgRowBytes;
        public long MaxRowBytes;
        public long EstInnoDbBytes;
        public long EstInnoDbPages;
        public double RowsPerInnoDbPage;
        public double RowShare;
        public double SizeShare;
        public double AvgNullRate;
        public double AvgCardinalityRate;
        public string InferredRole;
        public string ColumnTypeBreakdown;
        public int AllNullColumnCount;
        public int ConstantColumnCount;
        public int HighSkewColumnCount;
        public int LargeColumnCount;
        public string QualityRiskBreakdown;
        public string CandidatePks;
        public string DateColumns;
        public string PrimaryDateRange;
        public string CategoricalColumns;
        public string RelationSummary;
        public string JMeterReadiness;
        public int HighWarningCount;
        public int WarningCount;
        public string WarningSummary;
        public string RowMultisetHash;
        public RowChurnSketch RowSketch;
        public DbTableMetadata DbMeta;
        public List<ColumnAnalysis> Columns;

        public TableAnalysis()
        {
            Columns = new List<ColumnAnalysis>();
            ColumnTypeBreakdown = "-";
            QualityRiskBreakdown = "問題なし";
            CandidatePks = "";
            DateColumns = "";
            PrimaryDateRange = "-";
            CategoricalColumns = "";
            RelationSummary = "-";
            JMeterReadiness = "対象外";
            WarningSummary = "";
            RowMultisetHash = "";
        }
    }

    public class RelationshipAnalysis
    {
        public string ParentTable;
        public string ParentColumn;
        public string ChildTable;
        public string ChildColumn;
        public long ParentDistinctCount;
        public long ChildNonNullRows;
        public long ChildDistinctKeys;
        public long MatchedRows;
        public long OrphanRows;
        public double MatchRate;
        public long FanOutMin;
        public double FanOutAvg;
        public double FanOutP50;
        public double FanOutP95;
        public double FanOutP99;
        public long FanOutMax;
        public string MaxFanOutKey;
        public string RiskLevel;
        public string Note;
    }

    public class WarningItem
    {
        public string Severity; // 高 / 中 / 低
        public string Category;
        public string TableName;
        public string ColumnName;
        public string MetricValue;
        public string PerfImpact;
        public string Recommendation;
    }

    public class JMeterCandidateItem
    {
        public string TableName;
        public string KeyColumn;
        public long UniqueCount;
        public long TotalRows;
        public string DataType;
        public string Suitability;
        public string RecommendedUsage;
    }

    public class FolderAnalysisResult
    {
        public string InputFolder;
        public bool IsDirectDb;
        public string SourceMode;
        public DateTime AnalyzedAt;
        public int TableCount;
        public long TotalRows;
        public long TotalColumns;
        public long TotalFileBytes;
        public long TotalEstInnoDbBytes;
        public long TotalEstInnoDbPages;
        public int HighWarningCount;
        public int MediumWarningCount;
        public int LowWarningCount;
        public int TotalWarningCount;
        public int EmptyTableCount;
        public long RaggedRowTotal;
        public int CandidatePkColumnCount;
        public int AllNullColumnCount;
        public int HighSkewColumnCount;
        public int LargeColumnCount;
        public long TotalOrphanRows;
        public double OverallAvgNullRate;
        public double OverallAvgRowBytes;
        public long OverallMaxRowBytes;
        public double EstMinAcuCurrent;
        public double EstMinAcu10x;

        public List<TableAnalysis> Tables;
        public List<RelationshipAnalysis> Relationships;
        public List<WarningItem> Warnings;
        public List<JMeterCandidateItem> JMeterCandidates;
        public List<string> AuroraInsights;
        public DbServerMetadata DbServer;

        // Aggregated time-series across all datetime columns (primary timestamp per table)
        public SortedDictionary<string, long> GlobalMonthlyCounts;
        public long[] GlobalDayOfWeekCounts;
        public long[] GlobalHourlyCounts;
        public string PrimaryTimeSeriesSource;

        public FolderAnalysisResult()
        {
            IsDirectDb = false;
            SourceMode = "CSV";
            Tables = new List<TableAnalysis>();
            Relationships = new List<RelationshipAnalysis>();
            Warnings = new List<WarningItem>();
            JMeterCandidates = new List<JMeterCandidateItem>();
            AuroraInsights = new List<string>();
            GlobalMonthlyCounts = new SortedDictionary<string, long>(StringComparer.Ordinal);
            GlobalDayOfWeekCounts = new long[7];
            GlobalHourlyCounts = new long[24];
            PrimaryTimeSeriesSource = "";
        }
    }

    internal class ColumnAccumulator
    {
        private const int MAX_DISTINCT_TRACK = 200000;
        private const int MAX_RESERVOIR = 50000;
        private const int MAX_BYTE_SAMPLE = 20000;

        public string ColumnName;
        public int ColumnIndex;
        public long TotalRows;
        public long NonNullCount;
        public long NullCount;

        public long IntCount;
        public long FloatCount;
        public long BoolCount;
        public long DateOnlyCount;
        public long DateTimeCount;
        public long UuidCount;
        public long JsonCount;
        public long EmailCount;
        public long WhitespaceOnlyCount;
        public long MultilineCount;
        public long NonAsciiCount;

        public int MinCharLen;
        public int MaxCharLen;
        public long SumCharLen;
        public int MinByteLen;
        public int MaxByteLen;
        public long SumByteLen;
        public List<int> ByteLenSample;

        public double NumMin;
        public double NumMax;
        public double NumMean;
        public double NumM2;
        public double NumSum;
        public double NumSumComp;
        public long NumCount;
        public long ZeroCount;
        public long NegativeCount;
        public List<double> NumSample;

        public ulong HashSum1;
        public ulong HashSum2;
        public ulong HashXor1;
        public ulong HashXor2;

        public Dictionary<string, long> Freq;
        public bool DistinctCapped;

        public DateTime? MinDate;
        public DateTime? MaxDate;
        public bool HasTime;
        public SortedDictionary<string, long> MonthlyCounts;
        public Dictionary<string, long> DailyCounts;
        public long[] DayOfWeekCounts;
        public long[] HourlyCounts;

        private Random _rng;
        private int _cachedDayKey;
        private string _cachedDayText;
        private string _cachedMonthText;

        public ColumnAccumulator(string name, int idx)
        {
            ColumnName = name;
            ColumnIndex = idx;
            MinCharLen = int.MaxValue;
            MinByteLen = int.MaxValue;
            NumMin = double.MaxValue;
            NumMax = double.MinValue;
            ByteLenSample = new List<int>(256);
            NumSample = new List<double>(256);
            Freq = new Dictionary<string, long>(32, StringComparer.Ordinal);
            MonthlyCounts = new SortedDictionary<string, long>(StringComparer.Ordinal);
            DailyCounts = new Dictionary<string, long>(StringComparer.Ordinal);
            DayOfWeekCounts = new long[7];
            HourlyCounts = new long[24];
            _rng = new Random(1337 + idx * 31);
            _cachedDayKey = -1;
        }

        private static readonly string[] DateFormats = new string[]
        {
            "yyyy-MM-dd HH:mm:ss",
            "yyyy-MM-dd HH:mm:ss.f",
            "yyyy-MM-dd HH:mm:ss.ff",
            "yyyy-MM-dd HH:mm:ss.fff",
            "yyyy-MM-dd HH:mm:ss.ffff",
            "yyyy-MM-dd HH:mm:ss.fffff",
            "yyyy-MM-dd HH:mm:ss.ffffff",
            "yyyy-MM-ddTHH:mm:ss",
            "yyyy-MM-ddTHH:mm:ss.fff",
            "yyyy-MM-ddTHH:mm:ssZ",
            "yyyy-MM-ddTHH:mm:ss.fffZ",
            "yyyy/MM/dd HH:mm:ss",
            "yyyy/MM/dd H:mm:ss",
            "yyyy/M/d HH:mm:ss",
            "yyyy/M/d H:mm:ss",
            "yyyy-MM-dd",
            "yyyy-M-d",
            "yyyy/MM/dd",
            "yyyy/M/d"
        };

        public void Add(string raw, out int rowUtf8Bytes, out ulong hash1, out ulong hash2, out bool hashed)
        {
            hash1 = 0UL;
            hash2 = 0UL;
            hashed = false;
            rowUtf8Bytes = 0;
            TotalRows++;
            if (ContentHashHelper.IsNullToken(raw))
            {
                NullCount++;
                if (raw != null && raw.Length != 0) rowUtf8Bytes = ContentHashHelper.Utf8ByteCount(raw);
                return;
            }

            NonNullCount++;
            hashed = true;
            int byteLen;
            ContentHashHelper.HashString128Utf8(raw, out hash1, out hash2, out byteLen);
            rowUtf8Bytes = byteLen;
            unchecked
            {
                HashSum1 += hash1;
                HashSum2 += hash2;
                HashXor1 ^= hash1;
                HashXor2 ^= hash2;
            }
            int charLen = raw.Length;
            if (charLen < MinCharLen) MinCharLen = charLen;
            if (charLen > MaxCharLen) MaxCharLen = charLen;
            SumCharLen += charLen;
            if (byteLen < MinByteLen) MinByteLen = byteLen;
            if (byteLen > MaxByteLen) MaxByteLen = byteLen;
            SumByteLen += byteLen;

            if (ByteLenSample.Count < MAX_BYTE_SAMPLE)
            {
                ByteLenSample.Add(byteLen);
            }
            else
            {
                int slot = _rng.Next((int)Math.Min(int.MaxValue, NonNullCount));
                if (slot < MAX_BYTE_SAMPLE) ByteLenSample[slot] = byteLen;
            }

            if (raw.IndexOf('\n') >= 0 || raw.IndexOf('\r') >= 0) MultilineCount++;
            if (byteLen > charLen) NonAsciiCount++;

            // Frequency tracking (truncate ultra-long keys to 120 chars for safe memory)
            string fkey = raw.Length <= 120 ? raw : (raw.Substring(0, 117) + "...");
            long existing;
            if (Freq.TryGetValue(fkey, out existing))
            {
                Freq[fkey] = existing + 1;
            }
            else if (Freq.Count < MAX_DISTINCT_TRACK)
            {
                Freq[fkey] = 1;
            }
            else
            {
                DistinctCapped = true;
            }

            string trimmed = raw.Trim();
            if (trimmed.Length == 0)
            {
                WhitespaceOnlyCount++;
                return;
            }

            // Check Boolean/Flag
            if (trimmed == "0" || trimmed == "1" ||
                string.Equals(trimmed, "true", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(trimmed, "false", StringComparison.OrdinalIgnoreCase))
            {
                BoolCount++;
            }

            // Check Numeric (avoid leading-zero code strings like "00123")
            bool looksLikeLeadingZeroCode = (trimmed.Length > 1 && trimmed[0] == '0' && trimmed[1] != '.');
            if (!looksLikeLeadingZeroCode)
            {
                long intVal;
                double dblVal;
                if (TryParseInt64Strict(trimmed, out intVal))
                {
                    IntCount++;
                    AddNumeric((double)intVal);
                }
                else if (ShouldTryFloat(trimmed) &&
                         double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out dblVal) &&
                         !double.IsNaN(dblVal) && !double.IsInfinity(dblVal))
                {
                    FloatCount++;
                    AddNumeric(dblVal);
                }
            }

            // Check Date/DateTime (fast pre-filter: starts with 4 digits and length >= 8)
            if (trimmed.Length >= 8 && trimmed.Length <= 32 &&
                char.IsDigit(trimmed[0]) && char.IsDigit(trimmed[1]) &&
                char.IsDigit(trimmed[2]) && char.IsDigit(trimmed[3]) &&
                (trimmed[4] == '-' || trimmed[4] == '/'))
            {
                DateTime dt;
                bool hasTimePart;
                bool parsed = TryParsePaddedDate(trimmed, out dt, out hasTimePart);
                if (!parsed && DateTime.TryParseExact(trimmed, DateFormats, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out dt))
                {
                    parsed = true;
                    hasTimePart = (trimmed.Length > 10 && trimmed.IndexOf(':') >= 0);
                }
                if (parsed && dt.Year >= 1970 && dt.Year <= 2100) NoteDate(dt, hasTimePart);
            }

            // Check UUID / JSON / Email
            if (IsUuid(trimmed))
            {
                UuidCount++;
            }
            else if (trimmed.Length >= 2 && ((trimmed[0] == '{' && trimmed[trimmed.Length - 1] == '}') || (trimmed[0] == '[' && trimmed[trimmed.Length - 1] == ']')))
            {
                JsonCount++;
            }
            else if (trimmed.Length >= 5 && trimmed.Length <= 254 && trimmed.IndexOf(' ') < 0)
            {
                int atIdx = trimmed.IndexOf('@');
                if (atIdx > 0 && atIdx < trimmed.Length - 3 &&
                    trimmed.IndexOf('@', atIdx + 1) < 0 &&
                    trimmed.IndexOf('.', atIdx + 2) > atIdx &&
                    trimmed[trimmed.Length - 1] != '.')
                {
                    EmailCount++;
                }
            }
        }

        private static bool TryParseInt64Strict(string s, out long value)
        {
            value = 0L;
            if (string.IsNullOrEmpty(s)) return false;
            int i = 0;
            bool neg = false;
            if (s[0] == '+' || s[0] == '-')
            {
                neg = s[0] == '-';
                i++;
                if (i >= s.Length) return false;
            }
            ulong acc = 0UL;
            ulong limit = neg ? 9223372036854775808UL : 9223372036854775807UL;
            for (; i < s.Length; i++)
            {
                char ch = s[i];
                if (ch < '0' || ch > '9') return false;
                ulong d = (ulong)(ch - '0');
                if (acc > limit / 10UL || (acc == limit / 10UL && d > limit % 10UL)) return false;
                acc = acc * 10UL + d;
            }
            value = neg ? unchecked(-(long)acc) : (long)acc;
            return true;
        }

        private static bool ShouldTryFloat(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            bool digit = false;
            bool mark = false;
            for (int i = 0; i < s.Length; i++)
            {
                char ch = s[i];
                if (ch >= '0' && ch <= '9') digit = true;
                else if (ch == '.' || ch == 'e' || ch == 'E' || ch == '+' || ch == '-') mark = true;
                else return false;
            }
            if (!digit) return false;
            return mark || s.Length > 18;
        }

        private static bool IsUuid(string s)
        {
            if (s == null || s.Length != 36) return false;
            for (int i = 0; i < 36; i++)
            {
                char ch = s[i];
                if (i == 8 || i == 13 || i == 18 || i == 23)
                {
                    if (ch != '-') return false;
                }
                else if (!((ch >= '0' && ch <= '9') || (ch >= 'a' && ch <= 'f') || (ch >= 'A' && ch <= 'F')))
                {
                    return false;
                }
            }
            return true;
        }

        private static bool IsAsciiDigit(char ch)
        {
            return ch >= '0' && ch <= '9';
        }

        private static int TwoDigits(string s, int i)
        {
            return (s[i] - '0') * 10 + (s[i + 1] - '0');
        }

        private static bool TryMakeUtc(int year, int month, int day, int hour, int minute, int second, long fractionTicks, out DateTime dt)
        {
            dt = default(DateTime);
            try
            {
                dt = new DateTime(year, month, day, hour, minute, second, DateTimeKind.Utc);
                if (fractionTicks != 0L) dt = dt.AddTicks(fractionTicks);
                return true;
            }
            catch (ArgumentOutOfRangeException)
            {
                return false;
            }
        }

        // Zero-padded MySQL / ISO forms only. Every other shape falls back to TryParseExact.
        private static bool TryParsePaddedDate(string s, out DateTime dt, out bool hasTime)
        {
            dt = default(DateTime);
            hasTime = false;
            int n = s.Length;
            if (n < 10 || n > 27) return false;
            char sep = s[4];
            if ((sep != '-' && sep != '/') || s[7] != sep) return false;
            if (!IsAsciiDigit(s[0]) || !IsAsciiDigit(s[1]) || !IsAsciiDigit(s[2]) || !IsAsciiDigit(s[3]) ||
                !IsAsciiDigit(s[5]) || !IsAsciiDigit(s[6]) || !IsAsciiDigit(s[8]) || !IsAsciiDigit(s[9])) return false;
            int year = (s[0] - '0') * 1000 + (s[1] - '0') * 100 + (s[2] - '0') * 10 + (s[3] - '0');
            int month = TwoDigits(s, 5);
            int day = TwoDigits(s, 8);
            if (n == 10) return TryMakeUtc(year, month, day, 0, 0, 0, 0L, out dt);

            bool tee = s[10] == 'T';
            if (s[10] != ' ' && !tee) return false;
            if (sep == '/' && tee) return false;
            if (n < 19 || s[13] != ':' || s[16] != ':') return false;
            if (!IsAsciiDigit(s[11]) || !IsAsciiDigit(s[12]) || !IsAsciiDigit(s[14]) ||
                !IsAsciiDigit(s[15]) || !IsAsciiDigit(s[17]) || !IsAsciiDigit(s[18])) return false;
            int hour = TwoDigits(s, 11);
            int minute = TwoDigits(s, 14);
            int second = TwoDigits(s, 17);
            long fractionTicks = 0L;
            if (n == 19)
            {
                hasTime = true;
                return TryMakeUtc(year, month, day, hour, minute, second, 0L, out dt);
            }
            if (n == 20 && s[19] == 'Z')
            {
                if (!tee || sep != '-') return false;
                hasTime = true;
                return TryMakeUtc(year, month, day, hour, minute, second, 0L, out dt);
            }
            if (s[19] != '.') return false;
            int fracEnd = n;
            if (s[n - 1] == 'Z')
            {
                if (!tee || sep != '-') return false;
                fracEnd = n - 1;
            }
            int digits = fracEnd - 20;
            if (digits < 1 || digits > 6) return false;
            if (sep == '/') return false;
            if (tee && digits != 3) return false;
            int frac = 0;
            for (int i = 20; i < fracEnd; i++)
            {
                if (!IsAsciiDigit(s[i])) return false;
                frac = frac * 10 + (s[i] - '0');
            }
            for (int k = digits; k < 7; k++) frac *= 10;
            fractionTicks = frac;
            hasTime = true;
            return TryMakeUtc(year, month, day, hour, minute, second, fractionTicks, out dt);
        }

        private void NoteDate(DateTime dt, bool hasTimePart)
        {
            if (hasTimePart)
            {
                DateTimeCount++;
                HasTime = true;
            }
            else
            {
                DateOnlyCount++;
            }
            if (!MinDate.HasValue || dt < MinDate.Value) MinDate = dt;
            if (!MaxDate.HasValue || dt > MaxDate.Value) MaxDate = dt;

            int dayKey = dt.Year * 10000 + dt.Month * 100 + dt.Day;
            string ymd;
            string ym;
            if (_cachedDayKey == dayKey)
            {
                ymd = _cachedDayText;
                ym = _cachedMonthText;
            }
            else
            {
                ymd = FormatDay(dt.Year, dt.Month, dt.Day);
                ym = ymd.Substring(0, 7);
                _cachedDayKey = dayKey;
                _cachedDayText = ymd;
                _cachedMonthText = ym;
            }
            long mc;
            MonthlyCounts[ym] = MonthlyCounts.TryGetValue(ym, out mc) ? mc + 1 : 1;
            long dc;
            DailyCounts[ymd] = DailyCounts.TryGetValue(ymd, out dc) ? dc + 1 : 1;
            int dow = ((int)dt.DayOfWeek + 6) % 7;
            DayOfWeekCounts[dow]++;
            if (hasTimePart) HourlyCounts[dt.Hour]++;
        }

        private static string FormatDay(int year, int month, int day)
        {
            char[] chars = new char[10];
            chars[0] = (char)('0' + year / 1000);
            chars[1] = (char)('0' + (year / 100) % 10);
            chars[2] = (char)('0' + (year / 10) % 10);
            chars[3] = (char)('0' + year % 10);
            chars[4] = '-';
            chars[5] = (char)('0' + month / 10);
            chars[6] = (char)('0' + month % 10);
            chars[7] = '-';
            chars[8] = (char)('0' + day / 10);
            chars[9] = (char)('0' + day % 10);
            return new string(chars);
        }

        private void AddNumeric(double v)
        {
            NumCount++;
            if (v < NumMin) NumMin = v;
            if (v > NumMax) NumMax = v;
            if (v == 0.0) ZeroCount++;
            if (v < 0.0) NegativeCount++;

            // Welford online mean & variance + Neumaier compensated sum
            double sumT = NumSum + v;
            if (Math.Abs(NumSum) >= Math.Abs(v)) NumSumComp += (NumSum - sumT) + v;
            else NumSumComp += (v - sumT) + NumSum;
            NumSum = sumT;

            double delta = v - NumMean;
            NumMean += delta / NumCount;
            double delta2 = v - NumMean;
            NumM2 += delta * delta2;

            if (NumSample.Count < MAX_RESERVOIR)
            {
                NumSample.Add(v);
            }
            else
            {
                int slot = _rng.Next((int)Math.Min(int.MaxValue, NumCount));
                if (slot < MAX_RESERVOIR) NumSample[slot] = v;
            }
        }
    }

    // Kept only for the current ExportAndAnalyze run. A resumed process reads skipped tables from disk.
    public sealed class CsvExportAnalysis
    {
        internal readonly Dictionary<string, CsvTableProfile> Profiles =
            new Dictionary<string, CsvTableProfile>(StringComparer.OrdinalIgnoreCase);
    }

    internal sealed class CsvTableProfile
    {
        internal readonly TableAnalysis Table;
        internal readonly Dictionary<string, Dictionary<string, long>> KeyMaps =
            new Dictionary<string, Dictionary<string, long>>(StringComparer.OrdinalIgnoreCase);
        private List<ColumnAccumulator> _columns;
        private RowChurnSketchBuilder _sketchBuilder;
        private ulong[] _h1;
        private ulong[] _h2;
        private bool[] _hashed;
        private ulong _rowHashSum1, _rowHashSum2, _rowHashXor1, _rowHashXor2;
        private long _rows, _raggedRows, _sumRowBytes, _minRowBytes = long.MaxValue, _maxRowBytes;

        internal CsvTableProfile(string path, string encodingName, IList<string> headers)
        {
            Table = new TableAnalysis {
                FileName = Path.GetFileName(path), FilePath = path,
                TableName = Path.GetFileNameWithoutExtension(path), EncodingName = encodingName
            };
            _columns = CsvAnalysisEngine.CreateColumnAccumulators(headers);
            List<string> normHeaders = new List<string>(_columns.Count);
            for (int i = 0; i < _columns.Count; i++) normHeaders.Add(_columns[i].ColumnName);
            string keyLabel;
            int[] keyCols = ContentHashHelper.SelectSketchKeyColumns(normHeaders, null, out keyLabel);
            _sketchBuilder = new RowChurnSketchBuilder(keyCols, keyLabel, false, RowChurnSketchBuilder.DefaultCap);
            _h1 = new ulong[_columns.Count];
            _h2 = new ulong[_columns.Count];
            _hashed = new bool[_columns.Count];
        }

        internal void AddRow(IList<string> row)
        {
            if (row.Count == 1 && string.IsNullOrEmpty(row[0]) && _columns.Count > 1) return;
            if (row.Count != _columns.Count) _raggedRows++;
            CsvAnalysisEngine.RowHashAccum rh = CsvAnalysisEngine.AccumulateRow(row, _columns, _h1, _h2, _hashed, _sketchBuilder);
            unchecked
            {
                _rowHashSum1 += rh.H1;
                _rowHashSum2 += rh.H2;
                _rowHashXor1 ^= rh.H1;
                _rowHashXor2 ^= rh.H2;
            }
            _rows++;
            _sumRowBytes += rh.Utf8Bytes;
            if (rh.Utf8Bytes < _minRowBytes) _minRowBytes = rh.Utf8Bytes;
            if (rh.Utf8Bytes > _maxRowBytes) _maxRowBytes = rh.Utf8Bytes;
        }

        internal void Complete(long size)
        {
            Table.FileSizeBytes = size;
            long effectiveMin = _rows > 0 && _minRowBytes != long.MaxValue ? _minRowBytes : 0L;
            CsvAnalysisEngine.FinalizeTableAnalysis(Table, _columns, _rows, _raggedRows, _sumRowBytes, effectiveMin, _maxRowBytes, KeyMaps);
            unchecked
            {
                ulong rcMix = ContentHashHelper.Mix64((ulong)_rows ^ 0x517CC1B727220A95UL);
                Table.RowMultisetHash = ContentHashHelper.Format256(
                    _rowHashSum1 ^ rcMix,
                    _rowHashSum2 + rcMix,
                    _rowHashXor1 ^ ContentHashHelper.Mix64((ulong)Table.ColumnCount),
                    _rowHashXor2);
            }
            if (_sketchBuilder != null)
            {
                Table.RowSketch = _sketchBuilder.Build(Table);
                _sketchBuilder = null;
            }
            _columns = null; // Only finalized statistics and FK key maps survive between tables.
        }
    }

    public static class CsvAnalysisEngine
    {
        internal struct RowHashAccum
        {
            public ulong H1;
            public ulong H2;
            public long Utf8Bytes;
        }

        internal static RowHashAccum AccumulateRow(
            IList<string> row,
            List<ColumnAccumulator> columns,
            ulong[] hash1,
            ulong[] hash2,
            bool[] hashed,
            RowChurnSketchBuilder sketch)
        {
            RowHashAccum result = new RowHashAccum();
            int colCount = columns.Count;
            int vCount = row != null ? row.Count : 0;
            ulong a, b;
            ContentHashHelper.BeginRowHash(colCount, out a, out b);
            for (int i = 0; i < colCount; i++)
            {
                string value = i < vCount ? row[i] : null;
                int fieldBytes;
                ulong ch1, ch2;
                bool colHashed;
                columns[i].Add(value, out fieldBytes, out hash1[i], out hash2[i], out colHashed);
                hashed[i] = colHashed;
                result.Utf8Bytes += fieldBytes;
                if (colHashed)
                {
                    ch1 = hash1[i];
                    ch2 = hash2[i];
                }
                else
                {
                    ContentHashHelper.NullColumnHash(i, out ch1, out ch2);
                }
                ContentHashHelper.MixRowColumn(ref a, ref b, i, ch1, ch2);
            }
            result.H1 = a;
            result.H2 = b;
            if (sketch != null) sketch.AddFromHashes(hash1, hash2, hashed, a);
            return result;
        }

        public static FolderAnalysisResult AnalyzeFolder(string folderPath)
        {
            return AnalyzeFolder(folderPath, null, null);
        }

        public static FolderAnalysisResult AnalyzeFolder(string folderPath, OperationCancelState cancelState, Action<string> progressCallback)
        {
            return AnalyzeFiles(folderPath, null, null, cancelState, progressCallback);
        }

        internal static FolderAnalysisResult AnalyzeFiles(string folderPath, string[] selectedFiles, CsvExportAnalysis exportAnalysis,
            OperationCancelState cancelState, Action<string> progressCallback)
        {
            if (cancelState != null) cancelState.ThrowIfCancelled();
            if (string.IsNullOrEmpty(folderPath) || !Directory.Exists(folderPath))
            {
                throw new DirectoryNotFoundException("Input CSV folder not found: " + folderPath);
            }

            IEnumerable<string> inputs = selectedFiles ?? Directory.GetFiles(folderPath, "*.csv", SearchOption.TopDirectoryOnly)
                .Where(f => Path.GetExtension(f).Equals(".csv", StringComparison.OrdinalIgnoreCase) &&
                    !Path.GetFileName(f).StartsWith("~$") && !Path.GetFileName(f).StartsWith("_analysis_")).ToArray();
            string[] files = inputs
                .OrderBy(f => Path.GetFileName(f), StringComparer.OrdinalIgnoreCase)
                .ToArray();

            if (files.Length == 0)
            {
                throw new FileNotFoundException("No CSV files found in folder: " + folderPath);
            }

            FolderAnalysisResult result = new FolderAnalysisResult();
            result.InputFolder = folderPath;
            result.AnalyzedAt = DateTime.Now;

            // Keep key maps for cross-table FK & Fan-Out analysis
            Dictionary<string, Dictionary<string, Dictionary<string, long>>> keyMaps =
                new Dictionary<string, Dictionary<string, Dictionary<string, long>>>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < files.Length; i++)
            {
                if (cancelState != null) cancelState.ThrowIfCancelled();
                string file = files[i];
                CsvTableProfile profile = null;
                if (exportAnalysis != null) exportAnalysis.Profiles.TryGetValue(file, out profile);
                if (progressCallback != null)
                {
                    progressCallback(string.Format("  {0} [{1}/{2}]: {3}", profile == null ? "CSV解析中" : "出力時の集計を使用",
                        i + 1, files.Length, Path.GetFileName(file)));
                }
                if (profile == null) profile = AnalyzeCsvFile(file, cancelState);
                result.Tables.Add(profile.Table);
                keyMaps[profile.Table.TableName] = profile.KeyMaps;
            }

            if (cancelState != null) cancelState.ThrowIfCancelled();
            FinalizeFolderAnalysis(result, keyMaps, cancelState);
            return result;
        }

        internal static void FinalizeFolderAnalysis(
            FolderAnalysisResult result,
            Dictionary<string, Dictionary<string, Dictionary<string, long>>> keyMaps, OperationCancelState cancelState)
        {
            result.TotalRows = 0;
            result.TotalColumns = 0;
            result.TotalFileBytes = 0;
            result.TotalEstInnoDbBytes = 0;
            result.Relationships.Clear();
            result.Warnings.Clear();
            result.JMeterCandidates.Clear();
            result.AuroraInsights.Clear();
            result.GlobalMonthlyCounts.Clear();
            Array.Clear(result.GlobalDayOfWeekCounts, 0, result.GlobalDayOfWeekCounts.Length);
            Array.Clear(result.GlobalHourlyCounts, 0, result.GlobalHourlyCounts.Length);
            result.PrimaryTimeSeriesSource = "";

            foreach (TableAnalysis ta in result.Tables)
            {
                result.TotalRows += ta.RowCount;
                result.TotalColumns += ta.ColumnCount;
                result.TotalFileBytes += ta.FileSizeBytes;
                result.TotalEstInnoDbBytes += ta.EstInnoDbBytes;
            }

            result.TableCount = result.Tables.Count;

            // Compute cross-table shares
            foreach (TableAnalysis ta in result.Tables)
            {
                ta.RowShare = result.TotalRows > 0 ? ((double)ta.RowCount / result.TotalRows) : 0.0;
                ta.SizeShare = result.TotalFileBytes > 0 ? ((double)ta.FileSizeBytes / result.TotalFileBytes) : 0.0;
            }

            // Sort tables by RowCount desc (tie-break by FileSize desc, TableName asc)
            result.Tables = result.Tables
                .OrderByDescending(t => t.RowCount)
                .ThenByDescending(t => t.FileSizeBytes)
                .ThenBy(t => t.TableName, StringComparer.OrdinalIgnoreCase)
                .ToList();

            // Cross-table FK & 1:N Fan-Out analysis
            AnalyzeRelationships(result, keyMaps, cancelState);

            // Global primary time-series aggregation (pick richest datetime column from largest transaction table)
            BuildGlobalTimeSeries(result);

            // Consolidate Warnings, JMeter Candidates, and Aurora Serverless v2 Insights
            BuildWarningsAndRecommendations(result);
            if (cancelState != null) cancelState.ThrowIfCancelled();
        }

        private static Encoding DetectEncoding(string filePath, out string encodingName)
        {
            byte[] buf = new byte[65536];
            int read = 0;
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 65536, FileOptions.SequentialScan))
            {
                read = fs.Read(buf, 0, buf.Length);
            }
            if (read >= 3 && buf[0] == 0xEF && buf[1] == 0xBB && buf[2] == 0xBF)
            {
                encodingName = "UTF-8 (BOM)";
                return new UTF8Encoding(true);
            }
            if (read >= 2 && buf[0] == 0xFF && buf[1] == 0xFE)
            {
                encodingName = "UTF-16 LE";
                return Encoding.Unicode;
            }
            if (read >= 2 && buf[0] == 0xFE && buf[1] == 0xFF)
            {
                encodingName = "UTF-16 BE";
                return Encoding.BigEndianUnicode;
            }

            // Validate strict UTF-8
            bool hasNonAscii = false;
            for (int i = 0; i < read; i++)
            {
                if (buf[i] >= 0x80) { hasNonAscii = true; break; }
            }
            if (!hasNonAscii)
            {
                encodingName = "UTF-8 (ASCII)";
                return new UTF8Encoding(false);
            }

            try
            {
                // Trim potentially split multi-byte sequence at end of 64KB sample
                int validLen = read;
                while (validLen > 0 && validLen > read - 4 && (buf[validLen - 1] & 0xC0) == 0x80)
                {
                    validLen--;
                }
                if (validLen > 0 && validLen < read && buf[validLen - 1] >= 0xC0) validLen--;

                UTF8Encoding strictUtf8 = new UTF8Encoding(false, true);
                strictUtf8.GetString(buf, 0, validLen);
                encodingName = "UTF-8";
                return new UTF8Encoding(false);
            }
            catch
            {
                encodingName = "Shift-JIS (CP932)";
                return Encoding.GetEncoding(932);
            }
        }

        private sealed class CsvRecordReader
        {
            private readonly TextReader _reader;
            private readonly char[] _buf = new char[65536];
            private readonly StringBuilder _sb = new StringBuilder(64);
            private readonly List<string> _fields = new List<string>();
            private int _pos;
            private int _len;

            internal List<string> Fields { get { return _fields; } }

            internal CsvRecordReader(TextReader reader)
            {
                _reader = reader;
            }

            internal bool Read()
            {
                if (!Ensure()) return false;
                _fields.Clear();
                _sb.Length = 0;
                bool inQuotes = false;
                bool fieldStart = true;
                while (true)
                {
                    if (_pos >= _len && !Fill())
                    {
                        _fields.Add(FinishField());
                        return true;
                    }
                    char ch = _buf[_pos];
                    if (inQuotes)
                    {
                        if (ch == '"')
                        {
                            _pos++;
                            if (!Ensure())
                            {
                                inQuotes = false;
                                continue;
                            }
                            if (_buf[_pos] == '"')
                            {
                                _pos++;
                                _sb.Append('"');
                            }
                            else
                            {
                                inQuotes = false;
                            }
                        }
                        else
                        {
                            int start = _pos;
                            while (_pos < _len && _buf[_pos] != '"') _pos++;
                            _sb.Append(_buf, start, _pos - start);
                        }
                    }
                    else if (fieldStart && ch == '"')
                    {
                        _pos++;
                        inQuotes = true;
                        fieldStart = false;
                    }
                    else if (ch == ',')
                    {
                        _pos++;
                        _fields.Add(FinishField());
                        fieldStart = true;
                    }
                    else if (ch == '\r')
                    {
                        _pos++;
                        if (Ensure() && _buf[_pos] == '\n') _pos++;
                        _fields.Add(FinishField());
                        return true;
                    }
                    else if (ch == '\n')
                    {
                        _pos++;
                        _fields.Add(FinishField());
                        return true;
                    }
                    else
                    {
                        int start = _pos;
                        while (_pos < _len)
                        {
                            char c2 = _buf[_pos];
                            if (c2 == ',' || c2 == '\r' || c2 == '\n') break;
                            _pos++;
                        }
                        _sb.Append(_buf, start, _pos - start);
                        fieldStart = false;
                    }
                }
            }

            private string FinishField()
            {
                string value = _sb.ToString();
                _sb.Length = 0;
                return value;
            }

            private bool Ensure()
            {
                if (_pos < _len) return true;
                return Fill();
            }

            private bool Fill()
            {
                _len = _reader.Read(_buf, 0, _buf.Length);
                _pos = 0;
                return _len > 0;
            }
        }

        private static CsvTableProfile AnalyzeCsvFile(string filePath, OperationCancelState cancelState)
        {
            string encName;
            Encoding enc = DetectEncoding(filePath, out encName);
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 262144, FileOptions.SequentialScan))
            using (StreamReader sr = new StreamReader(fs, enc, true, 262144))
            {
                CsvRecordReader csv = new CsvRecordReader(sr);
                IList<string> header = csv.Read() ? csv.Fields : null;
                CsvTableProfile profile = new CsvTableProfile(filePath, encName, header);
                long rowsRead = 0;
                while (csv.Read())
                {
                    if (cancelState != null && (rowsRead++ & 2047) == 0) cancelState.ThrowIfCancelled();
                    profile.AddRow(csv.Fields);
                }
                if (cancelState != null) cancelState.ThrowIfCancelled();
                profile.Complete(fs.Length);
                return profile;
            }
        }

        internal static List<ColumnAccumulator> CreateColumnAccumulators(IList<string> rawHeaders)
        {
            List<ColumnAccumulator> accs = new List<ColumnAccumulator>();
            if (rawHeaders == null || rawHeaders.Count == 0) return accs;

            // Deduplicate empty/duplicate column names safely
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < rawHeaders.Count; i++)
            {
                string h = (rawHeaders[i] ?? "").Trim();
                if (i == 0 && h.Length > 0 && h[0] == '\uFEFF') h = h.Substring(1).Trim();
                if (string.IsNullOrEmpty(h)) h = "col_" + (i + 1).ToString(CultureInfo.InvariantCulture);
                string uniqueName = h;
                int suffix = 2;
                while (seen.Contains(uniqueName))
                {
                    uniqueName = h + "_" + suffix.ToString(CultureInfo.InvariantCulture);
                    suffix++;
                }
                seen.Add(uniqueName);
                accs.Add(new ColumnAccumulator(uniqueName, i));
            }
            return accs;
        }

        internal static void FinalizeTableAnalysis(
            TableAnalysis ta,
            List<ColumnAccumulator> accs,
            long rowCount,
            long raggedRows,
            long sumRowBytes,
            long minRowBytes,
            long maxRowBytes,
            Dictionary<string, Dictionary<string, long>> tableKeyMaps)
        {
            ta.RowCount = rowCount;
            ta.ColumnCount = accs.Count;
            ta.RaggedRowCount = raggedRows;
            ta.MinRowBytes = rowCount > 0 ? minRowBytes : 0L;
            ta.AvgRowBytes = rowCount > 0 ? ((double)sumRowBytes / rowCount) : 0.0;
            ta.MaxRowBytes = maxRowBytes;
            if (rowCount > 0)
            {
                double rawInno = (sumRowBytes + rowCount * 27.0) * 1.35;
                long pages = (long)Math.Ceiling(rawInno / 16384.0);
                if (pages < 1) pages = 1;
                ta.EstInnoDbBytes = pages * 16384L;
                ta.EstInnoDbPages = pages;
                ta.RowsPerInnoDbPage = (double)rowCount / pages;
            }
            else
            {
                ta.EstInnoDbBytes = 16384L; // 1 empty InnoDB page
                ta.EstInnoDbPages = 1L;
                ta.RowsPerInnoDbPage = 0.0;
            }

            List<string> pkCols = new List<string>();
            List<string> dtCols = new List<string>();
            List<string> catCols = new List<string>();
            Dictionary<string, int> typeCounts = new Dictionary<string, int>(StringComparer.Ordinal);
            double nullRateSum = 0.0;
            double cardRateSum = 0.0;
            int allNullCnt = 0;
            int constCnt = 0;
            int highSkewCnt = 0;
            int largeColCnt = 0;
            ColumnAnalysis bestDtCol = null;
            long bestDtScore = -1;

            foreach (ColumnAccumulator acc in accs)
            {
                ColumnAnalysis ca = FinalizeColumn(ta.TableName, acc);
                ta.Columns.Add(ca);
                nullRateSum += ca.NullRate;
                cardRateSum += ca.CardinalityRate;

                string shortType = ShortTypeName(ca.InferredType);
                int tc;
                typeCounts[shortType] = typeCounts.TryGetValue(shortType, out tc) ? (tc + 1) : 1;

                if (ca.TotalRows > 0 && ca.NullRate >= 0.999) allNullCnt++;
                if (ca.SelectivityRating.StartsWith("定数")) constCnt++;
                if (ca.SkewRisk.StartsWith("高リスク")) highSkewCnt++;
                if (ca.MaxByteLen >= 4000) largeColCnt++;

                if (ca.IsCandidatePk) pkCols.Add(ca.ColumnName);
                if (ca.IsDateTime)
                {
                    dtCols.Add(ca.ColumnName);
                    if (ca.MinDate.HasValue && ca.MaxDate.HasValue)
                    {
                        long dScore = ca.NonNullCount * (ca.HasTimeComponent ? 2L : 1L);
                        if (dScore > bestDtScore)
                        {
                            bestDtScore = dScore;
                            bestDtCol = ca;
                        }
                    }
                }
                if (ca.IsCategorical) catCols.Add(ca.ColumnName);

                // Retain bounded key map for FK/Fan-Out detection
                bool looksLikeKeyCol = ca.IsCandidatePk ||
                    ca.ColumnName.Equals("id", StringComparison.OrdinalIgnoreCase) ||
                    ca.ColumnName.EndsWith("_id", StringComparison.OrdinalIgnoreCase) ||
                    ca.ColumnName.EndsWith("Id", StringComparison.Ordinal) ||
                    ca.ColumnName.EndsWith("_cd", StringComparison.OrdinalIgnoreCase) ||
                    ca.ColumnName.EndsWith("_code", StringComparison.OrdinalIgnoreCase) ||
                    ca.ColumnName.EndsWith("_no", StringComparison.OrdinalIgnoreCase) ||
                    ca.ColumnName.EndsWith("_key", StringComparison.OrdinalIgnoreCase);

                if (looksLikeKeyCol && !acc.DistinctCapped && acc.Freq.Count > 0 && acc.Freq.Count <= 100000)
                {
                    tableKeyMaps[ca.ColumnName] = acc.Freq;
                }
            }

            ta.AvgNullRate = ta.ColumnCount > 0 ? (nullRateSum / ta.ColumnCount) : 0.0;
            ta.AvgCardinalityRate = ta.ColumnCount > 0 ? (cardRateSum / ta.ColumnCount) : 0.0;
            ta.CandidatePks = string.Join(", ", pkCols.ToArray());
            ta.DateColumns = string.Join(", ", dtCols.ToArray());
            ta.CategoricalColumns = string.Join(", ", catCols.ToArray());
            ta.AllNullColumnCount = allNullCnt;
            ta.ConstantColumnCount = constCnt;
            ta.HighSkewColumnCount = highSkewCnt;
            ta.LargeColumnCount = largeColCnt;

            if (typeCounts.Count > 0)
            {
                ta.ColumnTypeBreakdown = string.Join(" / ", typeCounts
                    .OrderByDescending(kv => kv.Value)
                    .ThenBy(kv => kv.Key, StringComparer.Ordinal)
                    .Select(kv => kv.Key + ":" + kv.Value.ToString(CultureInfo.InvariantCulture))
                    .ToArray());
            }
            else
            {
                ta.ColumnTypeBreakdown = "-";
            }

            List<string> qRisks = new List<string>();
            if (allNullCnt > 0) qRisks.Add("全NULL:" + allNullCnt.ToString(CultureInfo.InvariantCulture));
            if (constCnt > 0) qRisks.Add("定数:" + constCnt.ToString(CultureInfo.InvariantCulture));
            if (highSkewCnt > 0) qRisks.Add("高偏り:" + highSkewCnt.ToString(CultureInfo.InvariantCulture));
            if (largeColCnt > 0) qRisks.Add("大容量:" + largeColCnt.ToString(CultureInfo.InvariantCulture));
            ta.QualityRiskBreakdown = qRisks.Count > 0 ? string.Join(" / ", qRisks.ToArray()) : "問題なし";

            if (bestDtCol != null && bestDtCol.MinDate.HasValue && bestDtCol.MaxDate.HasValue)
            {
                ta.PrimaryDateRange = string.Format(CultureInfo.InvariantCulture, "{0}〜{1} ({2:0}日)",
                    bestDtCol.MinDate.Value.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture),
                    bestDtCol.MaxDate.Value.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture),
                    Math.Round(bestDtCol.DateSpanDays));
            }
            else
            {
                ta.PrimaryDateRange = "-";
            }

            // Infer Table Role for Performance Testing
            int fkLikeCount = ta.Columns.Count(c =>
                c.ColumnName.EndsWith("_id", StringComparison.OrdinalIgnoreCase) &&
                !c.ColumnName.Equals("id", StringComparison.OrdinalIgnoreCase));

            if (ta.RowCount == 0)
            {
                ta.InferredRole = "空テーブル (データ無)";
            }
            else if (ta.TableName.IndexOf("log", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     ta.TableName.IndexOf("history", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     ta.TableName.IndexOf("audit", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     ta.TableName.IndexOf("event", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                ta.InferredRole = "履歴・ログ系";
            }
            else if (ta.TableName.StartsWith("m_", StringComparison.OrdinalIgnoreCase) ||
                     ta.TableName.EndsWith("_mst", StringComparison.OrdinalIgnoreCase) ||
                     ta.TableName.IndexOf("master", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                ta.InferredRole = "マスタ系";
            }
            else if (fkLikeCount >= 2 && ta.ColumnCount <= 6)
            {
                ta.InferredRole = "中間・関連テーブル";
            }
            else if (ta.RowCount <= 500 && dtCols.Count == 0)
            {
                ta.InferredRole = "マスタ系";
            }
            else if (dtCols.Count > 0 && ta.RowCount >= 100)
            {
                ta.InferredRole = "トランザクション系";
            }
            else
            {
                ta.InferredRole = "業務エンティティ";
            }
        }

        private static string ShortTypeName(string inferredType)
        {
            if (string.IsNullOrEmpty(inferredType)) return "その他";
            if (inferredType.StartsWith("整数")) return "整数";
            if (inferredType.StartsWith("小数")) return "小数";
            if (inferredType.StartsWith("文字列")) return "文字";
            if (inferredType.StartsWith("長文")) return "長文";
            if (inferredType.StartsWith("日時")) return "日時";
            if (inferredType.StartsWith("日付")) return "日付";
            if (inferredType.StartsWith("フラグ")) return "フラグ";
            if (inferredType.StartsWith("メール")) return "メール";
            if (inferredType.StartsWith("空")) return "空";
            return inferredType;
        }

        private static string SanitizeCompactLabel(string raw, int maxLen)
        {
            if (string.IsNullOrEmpty(raw)) return "";
            string s = raw.Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ").Trim();
            if (s.Length > maxLen && maxLen > 3) s = s.Substring(0, maxLen - 3) + "...";
            return s;
        }

        private static ColumnAnalysis FinalizeColumn(string tableName, ColumnAccumulator acc)
        {
            ColumnAnalysis ca = new ColumnAnalysis();
            ca.TableName = tableName;
            ca.ColumnName = acc.ColumnName;
            ca.ColumnIndex = acc.ColumnIndex;
            ca.TotalRows = acc.TotalRows;
            ca.NonNullCount = acc.NonNullCount;
            ca.NullCount = acc.NullCount;
            ca.NullRate = acc.TotalRows > 0 ? ((double)acc.NullCount / acc.TotalRows) : 0.0;
            ca.DistinctCount = acc.Freq.Count;
            ca.DistinctCapped = acc.DistinctCapped;
            ca.CardinalityRate = acc.NonNullCount > 0 ? ((double)acc.Freq.Count / acc.NonNullCount) : 0.0;
            ca.WhitespaceOnlyCount = acc.WhitespaceOnlyCount;
            ca.MultilineCount = acc.MultilineCount;
            ca.NonAsciiCount = acc.NonAsciiCount;
            ca.NonAsciiRate = acc.NonNullCount > 0 ? ((double)acc.NonAsciiCount / acc.NonNullCount) : 0.0;

            unchecked
            {
                ulong nullMix = ContentHashHelper.Mix64((ulong)acc.NullCount ^ 0xA0761D6478BD642FUL);
                ulong nnMix = ContentHashHelper.Mix64((ulong)acc.NonNullCount ^ 0xE7037ED1A0B428DBUL);
                ca.ContentHash = ContentHashHelper.Format256(
                    acc.HashSum1 ^ nullMix,
                    acc.HashSum2 + nnMix,
                    acc.HashXor1 ^ nnMix,
                    acc.HashXor2 ^ nullMix);
            }
            if (!acc.DistinctCapped)
            {
                if (acc.Freq.Count == 0)
                {
                    ca.DistinctSetHash = "00000000000000000000000000000000";
                }
                else
                {
                    ulong ds1 = 0UL, ds2 = 0UL;
                    foreach (string fk in acc.Freq.Keys)
                    {
                        ulong kh1, kh2;
                        ContentHashHelper.HashString128(fk, out kh1, out kh2);
                        unchecked { ds1 += kh1; ds2 ^= kh2; }
                    }
                    ca.DistinctSetHash = ContentHashHelper.Format128(ds1, ds2);
                }
            }

            if (acc.NonNullCount == 0)
            {
                ca.InferredType = "空 (全NULL)";
                ca.KeyRole = "データなし";
                ca.SelectivityRating = "データなし";
                ca.SkewRisk = "要注意 (全NULL)";
                ca.DetailStatsSummary = "有効データなし (全行NULL)";
                ca.AlertNote = "全行NULLまたは空文字";
                return ca;
            }

            ca.MinCharLen = acc.MinCharLen == int.MaxValue ? 0 : acc.MinCharLen;
            ca.MaxCharLen = acc.MaxCharLen;
            ca.AvgCharLen = (double)acc.SumCharLen / acc.NonNullCount;
            ca.MinByteLen = acc.MinByteLen == int.MaxValue ? 0 : acc.MinByteLen;
            ca.MaxByteLen = acc.MaxByteLen;
            ca.AvgByteLen = (double)acc.SumByteLen / acc.NonNullCount;
            if (acc.ByteLenSample.Count > 0)
            {
                acc.ByteLenSample.Sort();
                int p95Idx = (int)Math.Min(acc.ByteLenSample.Count - 1, Math.Floor(0.95 * (acc.ByteLenSample.Count - 1)));
                ca.P95ByteLen = acc.ByteLenSample[p95Idx];
            }

            // Type inference
            double nn = (double)acc.NonNullCount;
            double numRatio = (acc.IntCount + acc.FloatCount) / nn;
            double dtRatio = (acc.DateOnlyCount + acc.DateTimeCount) / nn;
            double boolRatio = acc.BoolCount / nn;
            double uuidRatio = acc.UuidCount / nn;
            double jsonRatio = acc.JsonCount / nn;
            double emailRatio = acc.EmailCount / nn;

            bool nameLooksFlag = acc.ColumnName.StartsWith("is_", StringComparison.OrdinalIgnoreCase) ||
                                 acc.ColumnName.StartsWith("has_", StringComparison.OrdinalIgnoreCase) ||
                                 acc.ColumnName.EndsWith("_flg", StringComparison.OrdinalIgnoreCase) ||
                                 acc.ColumnName.EndsWith("_flag", StringComparison.OrdinalIgnoreCase);

            if (boolRatio >= 0.99 && acc.Freq.Count <= 2 && (nameLooksFlag || acc.IntCount == 0))
            {
                ca.InferredType = "フラグ (Boolean)";
            }
            else if (dtRatio >= 0.95)
            {
                ca.InferredType = acc.HasTime ? "日時 (DateTime)" : "日付 (Date)";
                ca.IsDateTime = true;
                ca.HasTimeComponent = acc.HasTime;
            }
            else if (numRatio >= 0.95)
            {
                ca.InferredType = (acc.FloatCount > 0) ? "小数 (Decimal/Float)" : "整数 (Integer)";
                ca.IsNumeric = true;
            }
            else if (uuidRatio >= 0.95)
            {
                ca.InferredType = "UUID";
            }
            else if (jsonRatio >= 0.90)
            {
                ca.InferredType = "JSON";
            }
            else if (emailRatio >= 0.90)
            {
                ca.InferredType = "メールアドレス";
            }
            else if (ca.MaxByteLen >= 1000)
            {
                ca.InferredType = "長文テキスト (TEXT)";
            }
            else
            {
                ca.InferredType = "文字列 (VARCHAR)";
            }

            // Candidate PK check
            bool nameLooksPk = acc.ColumnIndex == 0 ||
                               acc.ColumnName.Equals("id", StringComparison.OrdinalIgnoreCase) ||
                               acc.ColumnName.EndsWith("_id", StringComparison.OrdinalIgnoreCase) ||
                               acc.ColumnName.EndsWith("_cd", StringComparison.OrdinalIgnoreCase) ||
                               acc.ColumnName.EndsWith("_code", StringComparison.OrdinalIgnoreCase) ||
                               acc.ColumnName.EndsWith("_no", StringComparison.OrdinalIgnoreCase) ||
                               acc.ColumnName.EndsWith("_uuid", StringComparison.OrdinalIgnoreCase);

            ca.IsCandidatePk = (!acc.DistinctCapped && acc.NullCount == 0 && acc.TotalRows >= 1 &&
                                acc.Freq.Count == acc.TotalRows && nameLooksPk);

            // Selectivity rating
            if (acc.Freq.Count == 1 && !acc.DistinctCapped)
            {
                ca.SelectivityRating = "定数 (1種類)";
            }
            else if (ca.IsCandidatePk)
            {
                ca.SelectivityRating = "一意 (PK/UK候補)";
            }
            else if (ca.DistinctCapped || ca.CardinalityRate >= 0.50)
            {
                ca.SelectivityRating = "高選択性 (>=50%)";
            }
            else if (ca.CardinalityRate >= 0.05)
            {
                ca.SelectivityRating = "中選択性 (5-50%)";
            }
            else
            {
                ca.SelectivityRating = "低選択性 (区分値系)";
            }

            // Frequency distribution, Top-N shares, Entropy & Gini
            List<KeyValuePair<string, long>> sorted = acc.Freq
                .OrderByDescending(kv => kv.Value)
                .ThenBy(kv => kv.Key, StringComparer.Ordinal)
                .ToList();

            double cum = 0.0;
            for (int i = 0; i < sorted.Count && i < 20; i++)
            {
                double share = (double)sorted[i].Value / acc.NonNullCount;
                cum += share;
                ValueShareItem vsi = new ValueShareItem();
                vsi.Value = sorted[i].Key;
                vsi.Count = sorted[i].Value;
                vsi.Share = share;
                vsi.CumulativeShare = cum;
                ca.TopValues.Add(vsi);
            }

            if (ca.TopValues.Count > 0)
            {
                ca.Top1Value = ca.TopValues[0].Value;
                ca.Top1Count = ca.TopValues[0].Count;
                ca.Top1Share = ca.TopValues[0].Share;
                ca.Top3Share = ca.TopValues[Math.Min(2, ca.TopValues.Count - 1)].CumulativeShare;
                ca.Top5Share = ca.TopValues[Math.Min(4, ca.TopValues.Count - 1)].CumulativeShare;
                ca.Top10Share = ca.TopValues[Math.Min(9, ca.TopValues.Count - 1)].CumulativeShare;
            }

            // Shannon Entropy & Gini Coefficient
            if (sorted.Count > 1 && !acc.DistinctCapped)
            {
                double entropy = 0.0;
                double totalF = (double)acc.NonNullCount;
                for (int i = 0; i < sorted.Count; i++)
                {
                    double p = sorted[i].Value / totalF;
                    if (p > 0) entropy -= p * Math.Log(p, 2.0);
                }
                double maxEntropy = Math.Log(sorted.Count, 2.0);
                ca.NormalizedEntropy = maxEntropy > 0 ? (entropy / maxEntropy) : 0.0;

                // Gini on ascending frequencies
                double n = sorted.Count;
                double sumWeighted = 0.0;
                double sumVal = 0.0;
                for (int i = 0; i < sorted.Count; i++)
                {
                    long valAsc = sorted[sorted.Count - 1 - i].Value;
                    sumWeighted += (i + 1) * valAsc;
                    sumVal += valAsc;
                }
                if (sumVal > 0 && n > 1)
                {
                    ca.GiniCoefficient = Math.Max(0.0, Math.Min(1.0, (2.0 * sumWeighted) / (n * sumVal) - (n + 1.0) / n));
                }
            }
            else if (sorted.Count == 1)
            {
                ca.NormalizedEntropy = 0.0;
                ca.GiniCoefficient = 0.0;
            }
            else
            {
                ca.NormalizedEntropy = 1.0;
                ca.GiniCoefficient = 0.0;
            }

            // Categorical detection (ideal for JMeter scenario weights)
            bool isFlagOrCategory = !acc.DistinctCapped && !ca.IsDateTime && !ca.IsCandidatePk &&
                                    acc.NonNullCount >= 2 &&
                                    acc.Freq.Count >= 2 &&
                                    ((acc.Freq.Count <= 25 && ca.CardinalityRate <= 0.50) ||
                                     (acc.Freq.Count <= 15 && acc.NonNullCount <= 50));
            ca.IsCategorical = isFlagOrCategory;

            // Numeric stats & 10-bin Histogram
            if (ca.IsNumeric && acc.NumCount > 0)
            {
                ca.NumMin = acc.NumMin;
                ca.NumMax = acc.NumMax;
                ca.NumMean = acc.NumMean;
                ca.NumSum = acc.NumSum + acc.NumSumComp;
                ca.NumStdDev = acc.NumCount > 1 ? Math.Sqrt(acc.NumM2 / (acc.NumCount - 1)) : 0.0;
                ca.ZeroCount = acc.ZeroCount;
                ca.ZeroRate = (double)acc.ZeroCount / acc.NumCount;
                ca.NegativeCount = acc.NegativeCount;
                ca.NegativeRate = (double)acc.NegativeCount / acc.NumCount;

                if (acc.NumSample.Count > 0)
                {
                    acc.NumSample.Sort();
                    ca.NumP05 = Percentile(acc.NumSample, 0.05);
                    ca.NumP10 = Percentile(acc.NumSample, 0.10);
                    ca.NumP25 = Percentile(acc.NumSample, 0.25);
                    ca.NumMedian = Percentile(acc.NumSample, 0.50);
                    ca.NumP75 = Percentile(acc.NumSample, 0.75);
                    ca.NumP90 = Percentile(acc.NumSample, 0.90);
                    ca.NumP95 = Percentile(acc.NumSample, 0.95);
                    ca.NumP99 = Percentile(acc.NumSample, 0.99);

                    double iqr = ca.NumP75 - ca.NumP25;
                    if (iqr > 0)
                    {
                        double lowFence = ca.NumP25 - 1.5 * iqr;
                        double highFence = ca.NumP75 + 1.5 * iqr;
                        long outCnt = 0;
                        for (int i = 0; i < acc.NumSample.Count; i++)
                        {
                            double v = acc.NumSample[i];
                            if (v < lowFence || v > highFence) outCnt++;
                        }
                        ca.OutlierRate = (double)outCnt / acc.NumSample.Count;
                        ca.OutlierCount = (long)Math.Round(ca.OutlierRate * acc.NumCount);
                    }

                    // Build 10-bin histogram if range > 0 and not a pure unique PK ID
                    if (ca.NumMax > ca.NumMin && !ca.IsCandidatePk && acc.Freq.Count >= 3)
                    {
                        int bins = Math.Min(10, Math.Max(4, acc.Freq.Count));
                        double step = (ca.NumMax - ca.NumMin) / bins;
                        long[] counts = new long[bins];
                        for (int i = 0; i < acc.NumSample.Count; i++)
                        {
                            int b = (int)Math.Floor((acc.NumSample[i] - ca.NumMin) / step);
                            if (b < 0) b = 0;
                            if (b >= bins) b = bins - 1;
                            counts[b]++;
                        }
                        double hCum = 0.0;
                        double scale = (double)acc.NumCount / acc.NumSample.Count;
                        for (int b = 0; b < bins; b++)
                        {
                            double lo = ca.NumMin + b * step;
                            double hi = (b == bins - 1) ? ca.NumMax : (ca.NumMin + (b + 1) * step);
                            long estCount = (long)Math.Round(counts[b] * scale);
                            double share = (double)counts[b] / acc.NumSample.Count;
                            hCum += share;
                            HistogramBin hb = new HistogramBin();
                            hb.BinIndex = b + 1;
                            hb.LowerBound = lo;
                            hb.UpperBound = hi;
                            hb.Label = string.Format(CultureInfo.InvariantCulture, "{0:#,##0.##} - {1:#,##0.##}", lo, hi);
                            hb.Count = estCount;
                            hb.Share = share;
                            hb.CumulativeShare = Math.Min(1.0, hCum);
                            ca.Histogram.Add(hb);
                        }
                    }
                }
            }

            // Date/Time stats
            if (ca.IsDateTime && acc.MinDate.HasValue && acc.MaxDate.HasValue)
            {
                ca.MinDate = acc.MinDate;
                ca.MaxDate = acc.MaxDate;
                ca.DateSpanDays = Math.Max(1.0, (acc.MaxDate.Value - acc.MinDate.Value).TotalDays);
                ca.MonthlyCounts = acc.MonthlyCounts;
                ca.DayOfWeekCounts = acc.DayOfWeekCounts;
                ca.HourlyCounts = acc.HourlyCounts;

                if (acc.DailyCounts.Count > 0)
                {
                    ca.AvgPerDay = (double)acc.NonNullCount / acc.DailyCounts.Count;
                    string peakD = "";
                    long peakC = 0;
                    foreach (KeyValuePair<string, long> kv in acc.DailyCounts)
                    {
                        if (kv.Value > peakC)
                        {
                            peakC = kv.Value;
                            peakD = kv.Key;
                        }
                    }
                    ca.PeakDay = peakD;
                    ca.PeakDayCount = peakC;
                    ca.PeakBurstRatio = ca.AvgPerDay > 0 ? (peakC / ca.AvgPerDay) : 1.0;

                    List<string> sortedDays = new List<string>(acc.DailyCounts.Keys);
                    sortedDays.Sort(StringComparer.Ordinal);
                    int startIdx = Math.Max(0, sortedDays.Count - 31);
                    for (int di = startIdx; di < sortedDays.Count; di++)
                    {
                        string dk = sortedDays[di];
                        ca.RecentDailyCounts[dk] = acc.DailyCounts[dk];
                    }
                    if (!string.IsNullOrEmpty(peakD) && !ca.RecentDailyCounts.ContainsKey(peakD))
                    {
                        ca.RecentDailyCounts[peakD] = peakC;
                    }
                }
            }

            // Skew Risk & Alerts
            List<string> alerts = new List<string>();
            if (ca.NullRate >= 0.50)
            {
                alerts.Add(string.Format("高NULL率({0:P1})", ca.NullRate));
            }
            if (acc.Freq.Count == 1 && !acc.DistinctCapped && acc.NonNullCount >= 10)
            {
                alerts.Add("単一定数値");
            }
            if (ca.MaxByteLen >= 4000)
            {
                alerts.Add(string.Format("大容量ペイロード(最大{0:#,##0}B)", ca.MaxByteLen));
            }

            bool isBinaryFlag = (acc.Freq.Count <= 2 && (ca.InferredType.IndexOf("フラグ") >= 0 || nameLooksFlag));
            double concentrationRatio = (acc.Freq.Count > 0) ? (ca.Top1Share * acc.Freq.Count) : 1.0;
            ca.ConcentrationRatio = concentrationRatio;
            bool isWhaleHotKey = (acc.Freq.Count >= 5 && ca.Top1Share >= 0.20 && concentrationRatio >= 10.0);
            if (acc.Freq.Count >= 3 && !acc.DistinctCapped && acc.NonNullCount >= 20 &&
                (ca.Top1Share >= 0.50 || ca.GiniCoefficient >= 0.70 || isWhaleHotKey))
            {
                ca.SkewRisk = "高リスク (ホットキー偏り)";
                alerts.Add(string.Format("偏り大(Top1={0:P1}, Gini={1:0.00}, 集中={2:0.0}倍)", ca.Top1Share, ca.GiniCoefficient, concentrationRatio));
            }
            else if (!isBinaryFlag && acc.Freq.Count >= 2 && !acc.DistinctCapped && acc.NonNullCount >= 20 &&
                     (ca.Top1Share >= 0.35 || ca.GiniCoefficient >= 0.50))
            {
                ca.SkewRisk = "要注意 (中程度の偏り)";
            }
            else if (acc.Freq.Count == 1 && !acc.DistinctCapped)
            {
                ca.SkewRisk = "要注意 (定数)";
            }
            else
            {
                ca.SkewRisk = "正常";
            }

            if (ca.IsNumeric && ca.OutlierRate >= 0.05 && !ca.IsCandidatePk)
            {
                alerts.Add(string.Format("外れ値多({0:P1})", ca.OutlierRate));
            }

            // Key / Role classification
            if (ca.IsCandidatePk) ca.KeyRole = "PK/UK候補";
            else if (ca.IsDateTime) ca.KeyRole = "日時軸";
            else if (ca.IsCategorical) ca.KeyRole = "カテゴリ区分";
            else ca.KeyRole = "一般";

            // Min / Max display values
            if (ca.IsNumeric && acc.NumCount > 0)
            {
                ca.MinValueText = ca.NumMin.ToString("#,##0.##", CultureInfo.InvariantCulture);
                ca.MaxValueText = ca.NumMax.ToString("#,##0.##", CultureInfo.InvariantCulture);
            }
            else if (ca.IsDateTime && ca.MinDate.HasValue && ca.MaxDate.HasValue)
            {
                string dfmt = ca.HasTimeComponent ? "yyyy/MM/dd HH:mm:ss" : "yyyy/MM/dd";
                ca.MinValueText = ca.MinDate.Value.ToString(dfmt, CultureInfo.InvariantCulture);
                ca.MaxValueText = ca.MaxDate.Value.ToString(dfmt, CultureInfo.InvariantCulture);
            }

            // Top5 value distribution summary
            if (ca.TopValues.Count > 0)
            {
                if (ca.IsCandidatePk && ca.DistinctCount >= 3)
                {
                    List<string> exVals = new List<string>();
                    for (int i = 0; i < Math.Min(3, ca.TopValues.Count); i++)
                    {
                        exVals.Add(SanitizeCompactLabel(ca.TopValues[i].Value, 20));
                    }
                    ca.TopValuesSummary = "全件一意 (例: " + string.Join(", ", exVals.ToArray()) + ")";
                }
                else
                {
                    List<string> topParts = new List<string>();
                    int takeN = Math.Min(5, ca.TopValues.Count);
                    for (int i = 0; i < takeN; i++)
                    {
                        ValueShareItem tv = ca.TopValues[i];
                        topParts.Add(string.Format(CultureInfo.InvariantCulture, "{0}: {1:#,##0} ({2:0.0%})",
                            SanitizeCompactLabel(tv.Value, 24), tv.Count, tv.Share));
                    }
                    ca.TopValuesSummary = string.Join(" / ", topParts.ToArray());
                }
            }

            // Type-specific detailed statistical profile
            if (ca.IsNumeric && acc.NumCount > 0)
            {
                ca.DetailStatsSummary = string.Format(CultureInfo.InvariantCulture,
                    "平均:{0:#,##0.##} / 中央値:{1:#,##0.##} / P25:{2:#,##0.##} / P75:{3:#,##0.##} / P95:{4:#,##0.##} / P99:{5:#,##0.##} / 標準偏差:{6:#,##0.##} / ゼロ率:{7:0.0%} / 負値率:{8:0.0%} / 外れ値率:{9:0.0%}",
                    ca.NumMean, ca.NumMedian, ca.NumP25, ca.NumP75, ca.NumP95, ca.NumP99, ca.NumStdDev, ca.ZeroRate, ca.NegativeRate, ca.OutlierRate);
            }
            else if (ca.IsDateTime && ca.MinDate.HasValue && ca.MaxDate.HasValue)
            {
                StringBuilder ds = new StringBuilder();
                ds.AppendFormat(CultureInfo.InvariantCulture, "期間:{0:0.0}日 / 日平均:{1:#,##0.0}件 / ピーク日:{2} ({3:#,##0}件, {4:0.0}倍)",
                    ca.DateSpanDays, ca.AvgPerDay, ca.PeakDay, ca.PeakDayCount, ca.PeakBurstRatio);
                if (ca.HasTimeComponent)
                {
                    int pkHr = 0;
                    long pkHrCnt = 0;
                    long sumHr = 0;
                    for (int h = 0; h < 24; h++)
                    {
                        sumHr += ca.HourlyCounts[h];
                        if (ca.HourlyCounts[h] > pkHrCnt) { pkHrCnt = ca.HourlyCounts[h]; pkHr = h; }
                    }
                    if (sumHr > 0)
                    {
                        ds.AppendFormat(CultureInfo.InvariantCulture, " / ピーク時間帯:{0:00}:00台 ({1:0.0%})", pkHr, (double)pkHrCnt / sumHr);
                    }
                }
                ca.DetailStatsSummary = ds.ToString();
            }
            else
            {
                StringBuilder ss = new StringBuilder();
                ss.AppendFormat(CultureInfo.InvariantCulture, "文字長:{0}〜平均{1:0.0}〜最大{2}文字 / 非ASCII率:{3:0.0%}",
                    ca.MinCharLen, ca.AvgCharLen, ca.MaxCharLen, ca.NonAsciiRate);
                if (ca.MultilineCount > 0) ss.AppendFormat(CultureInfo.InvariantCulture, " / 改行含:{0:#,##0}件", ca.MultilineCount);
                if (ca.WhitespaceOnlyCount > 0) ss.AppendFormat(CultureInfo.InvariantCulture, " / 空白のみ:{0:#,##0}件", ca.WhitespaceOnlyCount);
                ca.DetailStatsSummary = ss.ToString();
            }

            ca.AlertNote = string.Join(" / ", alerts.ToArray());
            return ca;
        }

        private static double Percentile(List<double> sorted, double p)
        {
            if (sorted.Count == 0) return 0.0;
            if (sorted.Count == 1) return sorted[0];
            double pos = p * (sorted.Count - 1);
            int lo = (int)Math.Floor(pos);
            int hi = (int)Math.Ceiling(pos);
            if (lo == hi) return sorted[lo];
            double frac = pos - lo;
            return sorted[lo] * (1.0 - frac) + sorted[hi] * frac;
        }

        private static string Singularize(string name)
        {
            string s = name.ToLowerInvariant();
            if (s.StartsWith("t_") || s.StartsWith("m_")) s = s.Substring(2);
            if (s.EndsWith("ies") && s.Length > 3) return s.Substring(0, s.Length - 3) + "y";
            if (s.EndsWith("ses") && s.Length > 3) return s.Substring(0, s.Length - 2);
            if (s.EndsWith("s") && !s.EndsWith("ss") && s.Length > 2) return s.Substring(0, s.Length - 1);
            return s;
        }

        private static void AnalyzeRelationships(
            FolderAnalysisResult result,
            Dictionary<string, Dictionary<string, Dictionary<string, long>>> keyMaps, OperationCancelState cancelState)
        {
            // Collect parent candidate PKs
            List<Tuple<TableAnalysis, ColumnAnalysis, Dictionary<string, long>>> parents =
                new List<Tuple<TableAnalysis, ColumnAnalysis, Dictionary<string, long>>>();

            foreach (TableAnalysis ta in result.Tables)
            {
                Dictionary<string, Dictionary<string, long>> tMap;
                if (!keyMaps.TryGetValue(ta.TableName, out tMap)) continue;
                foreach (ColumnAnalysis ca in ta.Columns)
                {
                    Dictionary<string, long> freq;
                    if (ca.IsCandidatePk && tMap.TryGetValue(ca.ColumnName, out freq) && freq.Count > 0)
                    {
                        parents.Add(Tuple.Create(ta, ca, freq));
                    }
                }
            }

            HashSet<string> emitted = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (TableAnalysis child in result.Tables)
            {
                if (cancelState != null) cancelState.ThrowIfCancelled();
                Dictionary<string, Dictionary<string, long>> cMap;
                if (!keyMaps.TryGetValue(child.TableName, out cMap)) continue;

                foreach (ColumnAnalysis childCol in child.Columns)
                {
                    Dictionary<string, long> childFreq;
                    if (!cMap.TryGetValue(childCol.ColumnName, out childFreq) || childFreq.Count == 0) continue;

                    string cColLower = childCol.ColumnName.ToLowerInvariant();
                    if (cColLower == "id") continue;

                    Tuple<TableAnalysis, ColumnAnalysis, Dictionary<string, long>> bestParent = null;
                    double bestScore = 0.0;
                    long bestMatched = 0;

                    foreach (Tuple<TableAnalysis, ColumnAnalysis, Dictionary<string, long>> p in parents)
                    {
                        if (cancelState != null) cancelState.ThrowIfCancelled();
                        TableAnalysis parent = p.Item1;
                        ColumnAnalysis parentCol = p.Item2;
                        Dictionary<string, long> parentFreq = p.Item3;

                        if (string.Equals(parent.TableName, child.TableName, StringComparison.OrdinalIgnoreCase)) continue;

                        string pTableLower = parent.TableName.ToLowerInvariant();
                        string pSing = Singularize(parent.TableName);
                        string pColLower = parentCol.ColumnName.ToLowerInvariant();

                        bool nameMatch =
                            cColLower == pSing + "_id" ||
                            cColLower == pTableLower + "_id" ||
                            cColLower == pSing + "_cd" ||
                            cColLower == pSing + "_code" ||
                            cColLower == pSing + "_no" ||
                            (pColLower != "id" && cColLower == pColLower);

                        if (!nameMatch) continue;

                        long matchedRows = 0;
                        long matchedKeys = 0;
                        foreach (KeyValuePair<string, long> kv in childFreq)
                        {
                            if (parentFreq.ContainsKey(kv.Key))
                            {
                                matchedRows += kv.Value;
                                matchedKeys++;
                            }
                        }

                        if (matchedKeys == 0) continue;
                        double rate = (double)matchedRows / childCol.NonNullCount;
                        if (rate >= 0.30 && rate > bestScore)
                        {
                            bestScore = rate;
                            bestMatched = matchedRows;
                            bestParent = p;
                        }
                    }

                    if (bestParent != null)
                    {
                        string sig = bestParent.Item1.TableName + "." + bestParent.Item2.ColumnName + "->" + child.TableName + "." + childCol.ColumnName;
                        if (emitted.Contains(sig)) continue;
                        emitted.Add(sig);

                        List<double> fanOuts = new List<double>(childFreq.Count);
                        long maxFo = 0;
                        long minFo = long.MaxValue;
                        string maxKey = "";
                        long sumFo = 0;

                        foreach (KeyValuePair<string, long> kv in childFreq)
                        {
                            if (bestParent.Item3.ContainsKey(kv.Key))
                            {
                                fanOuts.Add((double)kv.Value);
                                sumFo += kv.Value;
                                if (kv.Value < minFo) minFo = kv.Value;
                                if (kv.Value > maxFo) { maxFo = kv.Value; maxKey = kv.Key; }
                            }
                        }

                        fanOuts.Sort();
                        RelationshipAnalysis rel = new RelationshipAnalysis();
                        rel.ParentTable = bestParent.Item1.TableName;
                        rel.ParentColumn = bestParent.Item2.ColumnName;
                        rel.ChildTable = child.TableName;
                        rel.ChildColumn = childCol.ColumnName;
                        rel.ParentDistinctCount = bestParent.Item3.Count;
                        rel.ChildNonNullRows = childCol.NonNullCount;
                        rel.ChildDistinctKeys = childFreq.Count;
                        rel.MatchedRows = bestMatched;
                        rel.OrphanRows = childCol.NonNullCount - bestMatched;
                        rel.MatchRate = childCol.NonNullCount > 0 ? ((double)bestMatched / childCol.NonNullCount) : 0.0;
                        rel.FanOutMin = minFo == long.MaxValue ? 0 : minFo;
                        rel.FanOutAvg = fanOuts.Count > 0 ? ((double)sumFo / fanOuts.Count) : 0.0;
                        rel.FanOutP50 = Percentile(fanOuts, 0.50);
                        rel.FanOutP95 = Percentile(fanOuts, 0.95);
                        rel.FanOutP99 = Percentile(fanOuts, 0.99);
                        rel.FanOutMax = maxFo;
                        rel.MaxFanOutKey = maxKey;

                        List<string> notes = new List<string>();
                        if (rel.OrphanRows > 0)
                        {
                            notes.Add(string.Format("孤立子レコード {0:#,##0}件", rel.OrphanRows));
                        }
                        if (rel.FanOutAvg > 0 && rel.FanOutMax >= Math.Max(50, rel.FanOutAvg * 10.0))
                        {
                            rel.RiskLevel = "高リスク (スーパーユーザー偏り)";
                            notes.Add(string.Format("最大ファンアウト {0:#,##0}件 (平均の{1:0.0}倍, Key={2})", rel.FanOutMax, rel.FanOutMax / rel.FanOutAvg, maxKey));
                        }
                        else if (rel.OrphanRows > 0)
                        {
                            rel.RiskLevel = "要注意 (参照整合性)";
                        }
                        else
                        {
                            rel.RiskLevel = "正常";
                            notes.Add("均衡な1:N分布");
                        }
                        rel.Note = string.Join(" / ", notes.ToArray());
                        result.Relationships.Add(rel);
                    }
                }
            }

            result.Relationships = result.Relationships
                .OrderByDescending(r => r.FanOutMax)
                .ThenByDescending(r => r.OrphanRows)
                .ToList();

            // Attach concise relation summaries to tables and columns
            Dictionary<string, TableAnalysis> tblMap = new Dictionary<string, TableAnalysis>(StringComparer.OrdinalIgnoreCase);
            foreach (TableAnalysis ta in result.Tables) tblMap[ta.TableName] = ta;

            Dictionary<string, List<string>> tblRelNotes = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, List<string>> parentColRefs = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

            foreach (RelationshipAnalysis rel in result.Relationships)
            {
                string orphanTxt = rel.OrphanRows > 0 ? string.Format(CultureInfo.InvariantCulture, ", 孤立{0:#,##0}", rel.OrphanRows) : "";
                string childTblTxt = string.Format(CultureInfo.InvariantCulture, "親:{0}(最大{1:#,##0}倍{2})", rel.ParentTable, rel.FanOutMax, orphanTxt);
                string parentTblTxt = string.Format(CultureInfo.InvariantCulture, "子:{0}(最大{1:#,##0}倍)", rel.ChildTable, rel.FanOutMax);

                List<string> cList;
                if (!tblRelNotes.TryGetValue(rel.ChildTable, out cList)) { cList = new List<string>(); tblRelNotes[rel.ChildTable] = cList; }
                cList.Add(childTblTxt);

                List<string> pList;
                if (!tblRelNotes.TryGetValue(rel.ParentTable, out pList)) { pList = new List<string>(); tblRelNotes[rel.ParentTable] = pList; }
                pList.Add(parentTblTxt);

                TableAnalysis cTa;
                if (tblMap.TryGetValue(rel.ChildTable, out cTa))
                {
                    ColumnAnalysis cCol = cTa.Columns.FirstOrDefault(x => string.Equals(x.ColumnName, rel.ChildColumn, StringComparison.OrdinalIgnoreCase));
                    if (cCol != null)
                    {
                        if (!cCol.IsCandidatePk) cCol.KeyRole = "FK候補";
                        cCol.RelationSummary = string.Format(CultureInfo.InvariantCulture,
                            "参照先: {0}.{1} (一致{2:0.0%}, 平均{3:0.0}/最大{4:#,##0}件/親{5})",
                            rel.ParentTable, rel.ParentColumn, rel.MatchRate, rel.FanOutAvg, rel.FanOutMax, orphanTxt);
                    }
                }

                string pKey = rel.ParentTable + "." + rel.ParentColumn;
                List<string> refList;
                if (!parentColRefs.TryGetValue(pKey, out refList)) { refList = new List<string>(); parentColRefs[pKey] = refList; }
                refList.Add(string.Format(CultureInfo.InvariantCulture, "{0}.{1}(最大{2:#,##0}倍)", rel.ChildTable, rel.ChildColumn, rel.FanOutMax));
            }

            foreach (TableAnalysis ta in result.Tables)
            {
                List<string> rn;
                if (tblRelNotes.TryGetValue(ta.TableName, out rn) && rn.Count > 0)
                {
                    ta.RelationSummary = string.Join(" / ", rn.Distinct().ToArray());
                }
                else
                {
                    ta.RelationSummary = "-";
                }

                foreach (ColumnAnalysis ca in ta.Columns)
                {
                    string pKey = ta.TableName + "." + ca.ColumnName;
                    List<string> refs;
                    if (parentColRefs.TryGetValue(pKey, out refs) && refs.Count > 0)
                    {
                        string pNote = "被参照: " + string.Join(", ", refs.ToArray());
                        ca.RelationSummary = (ca.RelationSummary == "-" || string.IsNullOrEmpty(ca.RelationSummary))
                            ? pNote
                            : (ca.RelationSummary + " / " + pNote);
                    }
                }
            }
        }

        private static void BuildGlobalTimeSeries(FolderAnalysisResult result)
        {
            ColumnAnalysis bestDt = null;
            long bestScore = -1;
            foreach (TableAnalysis ta in result.Tables)
            {
                foreach (ColumnAnalysis ca in ta.Columns)
                {
                    if (!ca.IsDateTime || ca.NonNullCount == 0) continue;
                    long score = ca.NonNullCount;
                    if (ca.HasTimeComponent) score *= 2;
                    if (ca.ColumnName.IndexOf("created", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        ca.ColumnName.IndexOf("order", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        ca.ColumnName.IndexOf("event", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        score = (long)(score * 1.5);
                    }
                    if (score > bestScore)
                    {
                        bestScore = score;
                        bestDt = ca;
                    }
                }
            }

            if (bestDt != null)
            {
                result.PrimaryTimeSeriesSource = bestDt.TableName + "." + bestDt.ColumnName;
                foreach (KeyValuePair<string, long> kv in bestDt.MonthlyCounts)
                {
                    result.GlobalMonthlyCounts[kv.Key] = kv.Value;
                }
                for (int i = 0; i < 7; i++) result.GlobalDayOfWeekCounts[i] = bestDt.DayOfWeekCounts[i];
                for (int i = 0; i < 24; i++) result.GlobalHourlyCounts[i] = bestDt.HourlyCounts[i];
            }
        }

        private static void BuildWarningsAndRecommendations(FolderAnalysisResult result)
        {
            foreach (TableAnalysis ta in result.Tables)
            {
                List<string> tWarns = new List<string>();

                if (ta.RowCount == 0)
                {
                    AddWarning(result, "高", "データ不足", ta.TableName, "-", "0件",
                        "空テーブルでは性能試験で実運用相当のINDEX探索・JOIN負荷が再現不可",
                        "性能試験前に本番想定件数相当のダミーデータを投入してください");
                    tWarns.Add("空テーブル");
                }
                else if (ta.RowCount < 10 && ta.InferredRole != "マスタ系")
                {
                    AddWarning(result, "中", "データ規模過小", ta.TableName, "-", ta.RowCount + "件",
                        "極小データではInnoDB Buffer Pool全キャッシュヒットとなり実I/O負荷が計測不可",
                        "本番同等スケールまでデータ拡充を検討してください");
                    tWarns.Add("件数過小");
                }

                if (ta.RaggedRowCount > 0)
                {
                    AddWarning(result, "中", "CSV書式異常", ta.TableName, "-", ta.RaggedRowCount + "行",
                        "ヘッダーと列数が一致しない行が存在（改行・区切り文字のエスケープ崩れ可能性）",
                        "エクスポート設定のクォート・改行コード設定を確認してください");
                    tWarns.Add("列数不一致行あり");
                }

                foreach (ColumnAnalysis ca in ta.Columns)
                {
                    if (ta.RowCount == 0) continue;

                    if (ca.NullRate >= 0.999)
                    {
                        AddWarning(result, "中", "全NULLカラム", ta.TableName, ca.ColumnName, "NULL率 100.0%",
                            "全行NULLのため該当カラム条件のWHERE検索・結合テストが機能しない可能性",
                            "テストデータ生成時に本番相当の非NULL率を設定してください");
                        tWarns.Add(ca.ColumnName + ":全NULL");
                    }
                    else if (ca.NullRate >= 0.50)
                    {
                        AddWarning(result, "低", "高NULL率", ta.TableName, ca.ColumnName, string.Format("NULL率 {0:P1}", ca.NullRate),
                            "高NULL率カラムに対するIS NOT NULL検索やINDEX選択性の偏りに注意",
                            "本番データ分布と乖離がないか確認してください");
                    }

                    if (ca.SkewRisk.StartsWith("高リスク"))
                    {
                        AddWarning(result, "高", "ホットキー・偏り", ta.TableName, ca.ColumnName,
                            string.Format("Top1={0:P1} / Gini={1:0.00} (値:{2})", ca.Top1Share, ca.GiniCoefficient, ca.Top1Value),
                            "Aurora InnoDBの特定行・ページへのロック競合(Lock Wait)やホットスポット起因のレイテンシ悪化リスク",
                            "JMeterパラメータ分布に実データの重み付けを反映し、同時更新競合を必ず検証してください");
                        tWarns.Add(ca.ColumnName + ":偏り大");
                    }

                    if (ca.MaxByteLen >= 4000)
                    {
                        AddWarning(result, "中", "大容量ペイロード", ta.TableName, ca.ColumnName,
                            string.Format("最大 {0:#,##0} Bytes (P95: {1:#,##0} B)", ca.MaxByteLen, ca.P95ByteLen),
                            "InnoDB Off-Page格納による追加I/O、Auroraネットワーク帯域圧迫、ACUスケール要因",
                            "SELECT * を避け必要列のみ取得するクエリ設計・ペイロード上限を確認してください");
                        tWarns.Add(ca.ColumnName + ":大容量");
                    }

                    if (ca.IsDateTime && ca.PeakBurstRatio >= 3.0 && ca.PeakDayCount >= 20)
                    {
                        AddWarning(result, "中", "時系列バースト", ta.TableName, ca.ColumnName,
                            string.Format("ピーク倍率 {0:0.0}倍 ({1}: {2:#,##0}件)", ca.PeakBurstRatio, ca.PeakDay, ca.PeakDayCount),
                            "通常日の数倍以上の突発トラフィック発生履歴あり。Aurora Serverless v2のACUスケール追従遅延リスク",
                            "バースト負荷シナリオを追加し、Min ACU・ウォームアップ時間を検証してください");
                    }

                    // JMeter CSV Data Set candidate detection
                    if (ca.IsCandidatePk && ca.DistinctCount >= 10)
                    {
                        JMeterCandidateItem jc = new JMeterCandidateItem();
                        jc.TableName = ta.TableName;
                        jc.KeyColumn = ca.ColumnName;
                        jc.UniqueCount = ca.DistinctCount;
                        jc.TotalRows = ta.RowCount;
                        jc.DataType = ca.InferredType;
                        if (ca.DistinctCount >= 1000)
                        {
                            jc.Suitability = "最適 (高並列対応)";
                            jc.RecommendedUsage = "JMeter CSV Data Set (Sharing: All threads / Unique) で衝突なし高負荷試験可能";
                        }
                        else if (ca.DistinctCount >= 100)
                        {
                            jc.Suitability = "良好 (中規模並列)";
                            jc.RecommendedUsage = "中規模負荷試験に適用可能。大規模試験時はデータ拡充を推奨";
                        }
                        else
                        {
                            jc.Suitability = "要注意 (キー数少)";
                            jc.RecommendedUsage = "キー数が少なくキャッシュ・ロック競合に偏る恐れ。データ増量またはRecycle設定要検討";
                        }
                        result.JMeterCandidates.Add(jc);
                    }
                }

                ta.WarningCount = tWarns.Count;
                ta.WarningSummary = string.Join(", ", tWarns.Take(4).ToArray());
            }

            // Relationship warnings
            foreach (RelationshipAnalysis rel in result.Relationships)
            {
                if (rel.OrphanRows > 0)
                {
                    AddWarning(result, "高", "参照整合性(孤立行)", rel.ChildTable, rel.ChildColumn,
                        string.Format("孤立 {0:#,##0}件 (一致率 {1:P1})", rel.OrphanRows, rel.MatchRate),
                        "親テーブル(" + rel.ParentTable + ")に存在しないFK値があり、INNER JOINで欠落・業務エラーの原因",
                        "テストデータの外部キー整合性を修復するか、異常系データとして意図的か確認してください");
                }
                if (rel.RiskLevel.StartsWith("高リスク"))
                {
                    AddWarning(result, "高", "極端な1:Nファンアウト", rel.ChildTable, rel.ChildColumn,
                        string.Format("最大 {0:#,##0}件/親 (平均 {1:0.0}, P95 {2:0.0})", rel.FanOutMax, rel.FanOutAvg, rel.FanOutP95),
                        "特定親キー(" + rel.ParentTable + "." + rel.ParentColumn + "=" + rel.MaxFanOutKey + ")取得時のレスポンス劣化・N+1/OOM/ACUスパイク要因",
                        "性能試験にスーパーユーザー(極大件数親キー)検索シナリオとページング検証を組み込んでください");
                }
            }

            result.Warnings = result.Warnings
                .OrderBy(w => w.Severity == "高" ? 0 : (w.Severity == "中" ? 1 : 2))
                .ThenBy(w => w.TableName, StringComparer.OrdinalIgnoreCase)
                .ToList();

            result.TotalWarningCount = result.Warnings.Count;
            result.HighWarningCount = result.Warnings.Count(w => w.Severity == "高");
            result.MediumWarningCount = result.Warnings.Count(w => w.Severity == "中");
            result.LowWarningCount = result.Warnings.Count(w => w.Severity == "低");
            result.JMeterCandidates = result.JMeterCandidates.OrderByDescending(j => j.UniqueCount).ToList();

            // Reconcile per-table warning counts (including relationship warnings) and JMeter readiness
            Dictionary<string, List<WarningItem>> warnByTable = new Dictionary<string, List<WarningItem>>(StringComparer.OrdinalIgnoreCase);
            foreach (WarningItem w in result.Warnings)
            {
                List<WarningItem> wl;
                if (!warnByTable.TryGetValue(w.TableName, out wl)) { wl = new List<WarningItem>(); warnByTable[w.TableName] = wl; }
                wl.Add(w);
            }
            Dictionary<string, JMeterCandidateItem> bestJmByTable = new Dictionary<string, JMeterCandidateItem>(StringComparer.OrdinalIgnoreCase);
            foreach (JMeterCandidateItem jc in result.JMeterCandidates)
            {
                if (!bestJmByTable.ContainsKey(jc.TableName)) bestJmByTable[jc.TableName] = jc;
            }

            long totalCells = 0L;
            long totalNullCells = 0L;
            double weightedRowBytesSum = 0.0;
            long overallMaxRowB = 0L;
            long totalInnoPages = 0L;
            int emptyTblCnt = 0;
            long raggedTotal = 0L;
            int pkColTotal = 0;
            int allNullTotal = 0;
            int highSkewTotal = 0;
            int largeColTotal = 0;

            foreach (TableAnalysis ta in result.Tables)
            {
                totalInnoPages += ta.EstInnoDbPages;
                if (ta.RowCount == 0) emptyTblCnt++;
                raggedTotal += ta.RaggedRowCount;
                weightedRowBytesSum += ta.AvgRowBytes * ta.RowCount;
                if (ta.MaxRowBytes > overallMaxRowB) overallMaxRowB = ta.MaxRowBytes;
                allNullTotal += ta.AllNullColumnCount;
                highSkewTotal += ta.HighSkewColumnCount;
                largeColTotal += ta.LargeColumnCount;

                foreach (ColumnAnalysis ca in ta.Columns)
                {
                    totalCells += ca.TotalRows;
                    totalNullCells += ca.NullCount;
                    if (ca.IsCandidatePk) pkColTotal++;
                }

                List<WarningItem> twList;
                if (warnByTable.TryGetValue(ta.TableName, out twList))
                {
                    ta.WarningCount = twList.Count;
                    ta.HighWarningCount = twList.Count(x => x.Severity == "高");
                    List<string> extraRelWarns = new List<string>();
                    foreach (WarningItem rw in twList)
                    {
                        if (rw.Category == "参照整合性(孤立行)") extraRelWarns.Add(rw.ColumnName + ":孤立FK");
                        else if (rw.Category == "極端な1:Nファンアウト") extraRelWarns.Add(rw.ColumnName + ":極端ファンアウト");
                    }
                    if (extraRelWarns.Count > 0)
                    {
                        string extraStr = string.Join(", ", extraRelWarns.Distinct().ToArray());
                        ta.WarningSummary = string.IsNullOrEmpty(ta.WarningSummary) ? extraStr : (ta.WarningSummary + ", " + extraStr);
                    }
                }
                else
                {
                    ta.WarningCount = 0;
                    ta.HighWarningCount = 0;
                }

                JMeterCandidateItem bestJc;
                if (ta.RowCount == 0)
                {
                    ta.JMeterReadiness = "対象外 (空テーブル)";
                }
                else if (bestJmByTable.TryGetValue(ta.TableName, out bestJc))
                {
                    ta.JMeterReadiness = string.Format(CultureInfo.InvariantCulture, "{0} ({1}: {2:#,##0}件)",
                        bestJc.Suitability, bestJc.KeyColumn, bestJc.UniqueCount);
                }
                else if (!string.IsNullOrEmpty(ta.CandidatePks))
                {
                    ta.JMeterReadiness = "要注意 (キー数少 <10件)";
                }
                else
                {
                    ta.JMeterReadiness = "PK候補なし";
                }
            }

            result.TotalEstInnoDbPages = totalInnoPages;
            result.EmptyTableCount = emptyTblCnt;
            result.RaggedRowTotal = raggedTotal;
            result.CandidatePkColumnCount = pkColTotal;
            result.AllNullColumnCount = allNullTotal;
            result.HighSkewColumnCount = highSkewTotal;
            result.LargeColumnCount = largeColTotal;
            result.TotalOrphanRows = result.Relationships.Sum(r => r.OrphanRows);
            result.OverallAvgNullRate = totalCells > 0 ? ((double)totalNullCells / totalCells) : 0.0;
            result.OverallAvgRowBytes = result.TotalRows > 0 ? (weightedRowBytesSum / result.TotalRows) : 0.0;
            result.OverallMaxRowBytes = overallMaxRowB;

            double estInnoGb = result.TotalEstInnoDbBytes / (1024.0 * 1024.0 * 1024.0);
            result.EstMinAcuCurrent = Math.Max(0.5, Math.Ceiling((estInnoGb / 1.5) * 2.0) / 2.0);
            result.EstMinAcu10x = Math.Max(0.5, Math.Ceiling(((estInnoGb * 10.0) / 1.5) * 2.0) / 2.0);

            // Synthesize detailed Aurora Serverless v2 (MySQL 8.4.7) Insights
            double estInnoMb = result.TotalEstInnoDbBytes / (1024.0 * 1024.0);
            string sizeMetricLabel = result.IsDirectDb ? "DB直接取得(CSV換算)" : "CSV合計";
            result.AuroraInsights.Add(string.Format(CultureInfo.InvariantCulture,
                "【データ規模・Buffer Pool・ACU設計】全 {0} テーブル / 総レコード {1:#,##0} 件 / {2} {3:#,##0.00} MB / 推定InnoDB容量 {4:#,##0.00} MB ({5:#,##0} ページ, 全体平均行長 {6:#,##0.0} B / 最大 {7:#,##0} B)。1 ACU≒Buffer Pool約1.5GB換算での最低ACU目安は現状 {8:0.0} ACU・10倍データ時 {9:0.0} ACUです。本番想定倍率と同時接続ワークメモリを加味してMin/Max ACUとウォームアップ時間を設計してください。",
                result.TableCount, result.TotalRows, sizeMetricLabel, result.TotalFileBytes / (1024.0 * 1024.0), estInnoMb,
                result.TotalEstInnoDbPages, result.OverallAvgRowBytes, result.OverallMaxRowBytes, result.EstMinAcuCurrent, result.EstMinAcu10x));

            if (result.Tables.Count > 0)
            {
                TableAnalysis topT = result.Tables[0];
                int topN = Math.Min(3, result.Tables.Count);
                double top3RowShare = result.Tables.Take(topN).Sum(t => t.RowShare);
                double top3SizeShare = result.Tables.Take(topN).Sum(t => t.SizeShare);
                string top3Names = string.Join(", ", result.Tables.Take(topN).Select(t => "`" + t.TableName + "`").ToArray());
                result.AuroraInsights.Add(string.Format(CultureInfo.InvariantCulture,
                    "【支配的テーブル・負荷集中】最大テーブルは `{0}` ({1:#,##0} 件 / 全体レコードの {2:P1}, 推定InnoDB {3:#,##0.0} KB)。上位 {4} テーブル ({5}) で全レコードの {6:P1}・全容量の {7:P1} を占有しており、INDEX設計(カバリングインデックス・パーティション検討)および負荷シナリオの最優先対象です。",
                    topT.TableName, topT.RowCount, topT.RowShare, topT.EstInnoDbBytes / 1024.0, topN, top3Names, top3RowShare, top3SizeShare));
            }

            result.AuroraInsights.Add(string.Format(CultureInfo.InvariantCulture,
                "【スキーマ構成・データ品質】全 {0:#,##0} カラム (PK/UK候補 {1:#,##0} 列)、全体セルNULL率 {2:P1}。全NULLカラム {3} 列・大容量ペイロード(>=4KB) {4} 列・CSV列数不一致行 {5:#,##0} 行を検出。{6}",
                result.TotalColumns, result.CandidatePkColumnCount, result.OverallAvgNullRate,
                result.AllNullColumnCount, result.LargeColumnCount, result.RaggedRowTotal,
                (result.LargeColumnCount > 0 ? "大容量列を含むテーブルは SELECT * を避け必要列のみ取得するクエリ設計を推奨します。" : "極端な大容量ペイロードによるOff-Page I/O要因は検出されませんでした。")));

            List<WarningItem> skewWarns = result.Warnings.Where(w => w.Category == "ホットキー・偏り").ToList();
            if (skewWarns.Count > 0)
            {
                string skewEx = string.Join(", ", skewWarns.Take(4).Select(w => "`" + w.TableName + "." + w.ColumnName + "`").ToArray());
                result.AuroraInsights.Add(string.Format(CultureInfo.InvariantCulture,
                    "【ロック競合・ホットスポット】偏り高リスクカラムが {0} 件検出されました ({1}{2})。MySQL 8.4 InnoDB の行ロック・ギャップロック競合を再現するため、JMeterでは均一ランダムではなく実データのTop-N重み付け分布を使用してください。",
                    skewWarns.Count, skewEx, skewWarns.Count > 4 ? " 他" : ""));
            }
            else
            {
                result.AuroraInsights.Add("【データ偏り状況】極端なホットキー偏りは検出されませんでした。標準的な分布ベースおよび境界値シナリオで負荷試験を構成できます。");
            }

            if (result.Relationships.Count > 0)
            {
                RelationshipAnalysis maxRel = result.Relationships[0];
                result.AuroraInsights.Add(string.Format(CultureInfo.InvariantCulture,
                    "【リレーション・1:Nファンアウト・参照整合性】テーブル間FKリレーションを {0} 件推定。孤立子レコード合計 {1:#,##0} 行、最大ファンアウトは `{2}.{3}` → `{4}.{5}` の {6:#,##0} 件/親 (平均 {7:0.0} / P95 {8:0.0}, Key={9})。スーパーユーザー検索・JOIN・ページング時のレスポンス劣化とN+1発生有無を検証してください。",
                    result.Relationships.Count, result.TotalOrphanRows, maxRel.ParentTable, maxRel.ParentColumn,
                    maxRel.ChildTable, maxRel.ChildColumn, maxRel.FanOutMax, maxRel.FanOutAvg, maxRel.FanOutP95, maxRel.MaxFanOutKey));
            }

            if (!string.IsNullOrEmpty(result.PrimaryTimeSeriesSource))
            {
                int peakHour = 0;
                long peakHourCnt = 0;
                long totalHourly = 0;
                for (int h = 0; h < 24; h++)
                {
                    totalHourly += result.GlobalHourlyCounts[h];
                    if (result.GlobalHourlyCounts[h] > peakHourCnt)
                    {
                        peakHourCnt = result.GlobalHourlyCounts[h];
                        peakHour = h;
                    }
                }
                string[] dowLabels = new string[] { "月曜", "火曜", "水曜", "木曜", "金曜", "土曜", "日曜" };
                int peakDow = 0;
                long peakDowCnt = 0;
                long totalDow = 0;
                for (int d = 0; d < 7; d++)
                {
                    totalDow += result.GlobalDayOfWeekCounts[d];
                    if (result.GlobalDayOfWeekCounts[d] > peakDowCnt) { peakDowCnt = result.GlobalDayOfWeekCounts[d]; peakDow = d; }
                }
                if (totalHourly > 0)
                {
                    string dowNote = totalDow > 0 ? string.Format(CultureInfo.InvariantCulture, "、曜日ピークは{0} ({1:P1})", dowLabels[peakDow], (double)peakDowCnt / totalDow) : "";
                    result.AuroraInsights.Add(string.Format(CultureInfo.InvariantCulture,
                        "【ピーク時間帯・時系列モデル】代表日時軸 `{0}` では {1:00}:00台 が最大ピーク ({2:#,##0}件 / 全体の{3:P1}){4}。JMeterのStepping/Concurrency Thread Groupのランプアップ・ピーク維持設計の基準として活用できます。",
                        result.PrimaryTimeSeriesSource, peakHour, peakHourCnt, (double)peakHourCnt / totalHourly, dowNote));
                }
            }

            int optimalJm = result.JMeterCandidates.Count(j => j.Suitability.StartsWith("最適"));
            int goodJm = result.JMeterCandidates.Count(j => j.Suitability.StartsWith("良好"));
            if (result.EmptyTableCount > 0)
            {
                result.AuroraInsights.Add(string.Format(CultureInfo.InvariantCulture,
                    "【テストデータ充足性・JMeterパラメータ化】空テーブルが {0} 個存在します。実運用のJOIN・集計・INDEX統計情報(ANALYZE TABLE)を正確に再現するため、試験前データ投入を実施してください。なお現在、JMeter高並列適性キー(>=1,000件)は {1} 列、中規模並列キー(>=100件)は {2} 列検出されています。",
                    result.EmptyTableCount, optimalJm, goodJm));
            }
            else
            {
                result.AuroraInsights.Add(string.Format(CultureInfo.InvariantCulture,
                    "【JMeterパラメータ化適性】空テーブルは0件です。JMeter CSV Data Set向けに高並列適性キー(>=1,000件) {0} 列・中規模並列キー(>=100件) {1} 列を検出しました。同時更新系シナリオではスレッド間で一意キーを分割(Sharing: All threads)してください。",
                    optimalJm, goodJm));
            }
        }

        private static void AddWarning(FolderAnalysisResult res, string sev, string cat, string tbl, string col, string metric, string impact, string rec)
        {
            WarningItem w = new WarningItem();
            w.Severity = sev;
            w.Category = cat;
            w.TableName = tbl;
            w.ColumnName = col;
            w.MetricValue = metric;
            w.PerfImpact = impact;
            w.Recommendation = rec;
            res.Warnings.Add(w);
        }
    }

    // =========================================================================
    //  Direct DB Streaming Analysis Session (Zero CSV Disk Output)
    // =========================================================================
    public class DirectAnalysisSession
    {
        private readonly FolderAnalysisResult _result;
        private readonly Dictionary<string, Dictionary<string, Dictionary<string, long>>> _keyMaps;
        private Dictionary<string, DbTableMetadata> _tableMetaMap;
        private Dictionary<string, Dictionary<string, DbColumnMetadata>> _colMetaMap;

        public DirectAnalysisSession(string sourceDescription)
        {
            _result = new FolderAnalysisResult();
            _result.InputFolder = sourceDescription ?? "";
            _result.IsDirectDb = true;
            _result.SourceMode = "DirectDB";
            _result.AnalyzedAt = DateTime.Now;
            _keyMaps = new Dictionary<string, Dictionary<string, Dictionary<string, long>>>(StringComparer.OrdinalIgnoreCase);
            _tableMetaMap = new Dictionary<string, DbTableMetadata>(StringComparer.OrdinalIgnoreCase);
            _colMetaMap = new Dictionary<string, Dictionary<string, DbColumnMetadata>>(StringComparer.OrdinalIgnoreCase);
        }

        public void AttachDbMetadata(
            DbServerMetadata serverMeta,
            Dictionary<string, DbTableMetadata> tableMetaMap,
            Dictionary<string, Dictionary<string, DbColumnMetadata>> colMetaMap)
        {
            _result.DbServer = serverMeta;
            if (tableMetaMap != null) _tableMetaMap = tableMetaMap;
            if (colMetaMap != null) _colMetaMap = colMetaMap;
        }

        public TableAnalysis AnalyzeMySqlTable(
            MySql84Client client,
            string tableName,
            string databaseName,
            string sql,
            int maxRows)
        {
            return AnalyzeMySqlTable(client, tableName, databaseName, sql, maxRows, null, null);
        }

        public TableAnalysis AnalyzeMySqlTable(
            MySql84Client client,
            string tableName,
            string databaseName,
            string sql,
            int maxRows,
            OperationCancelState cancelState,
            Action<long> onRowProgress)
        {
            if (client == null) throw new ArgumentNullException("client");
            if (string.IsNullOrEmpty(tableName)) throw new ArgumentException("Table name is required.", "tableName");
            if (cancelState != null) cancelState.ThrowIfCancelled();

            Dictionary<string, Dictionary<string, long>> tableKeyMaps =
                new Dictionary<string, Dictionary<string, long>>(StringComparer.OrdinalIgnoreCase);

            TableAnalysis ta = new TableAnalysis();
            ta.TableName = tableName;
            ta.FileName = string.IsNullOrEmpty(databaseName) ? tableName : (databaseName + "." + tableName);
            ta.FilePath = "";
            ta.EncodingName = "UTF-8 (DB直接)";

            Dictionary<string, DbColumnMetadata> tblColMeta = null;
            _colMetaMap.TryGetValue(tableName, out tblColMeta);

            List<string> explicitPkCols = null;
            if (tblColMeta != null && tblColMeta.Count > 0)
            {
                List<DbColumnMetadata> priCols = tblColMeta.Values
                    .Where(cm => string.Equals(cm.ColumnKey, "PRI", StringComparison.OrdinalIgnoreCase))
                    .OrderBy(cm => cm.OrdinalPosition)
                    .ToList();
                if (priCols.Count > 0)
                {
                    explicitPkCols = new List<string>(priCols.Count);
                    for (int pi = 0; pi < priCols.Count; pi++) explicitPkCols.Add(priCols[pi].ColumnName);
                }
            }

            List<ColumnAccumulator> accs = new List<ColumnAccumulator>();
            RowChurnSketchBuilder sketchBuilder = null;
            ulong[] hash1 = null;
            ulong[] hash2 = null;
            bool[] hashed = null;
            ulong rowHashSum1 = 0UL, rowHashSum2 = 0UL, rowHashXor1 = 0UL, rowHashXor2 = 0UL;
            long rowCount = 0;
            long sumRowBytes = 0;
            long minRowBytes = long.MaxValue;
            long maxRowBytes = 0;
            long csvEquivBytes = 3L; // UTF-8 BOM equivalent (matches ExportQueryToCsv writeUtf8Bom=true)
            bool receivedColumns = false;
            int colCount = 0;

            client.StreamQuery(
                sql,
                maxRows,
                delegate(List<MySqlColumnInfo> cols)
                {
                    receivedColumns = true;
                    List<string> headers = new List<string>(cols.Count);
                    List<string> wirePriCols = new List<string>();
                    for (int i = 0; i < cols.Count; i++)
                    {
                        string cn = cols[i].Name ?? "";
                        headers.Add(cn);
                        csvEquivBytes += MySql84Client.CalcCsvEncodedUtf8Length(cn);
                        if ((cols[i].Flags & 0x0002) != 0) wirePriCols.Add(cn);
                    }
                    if (cols.Count > 1) csvEquivBytes += (cols.Count - 1);
                    csvEquivBytes += 2L; // CRLF
                    accs = CsvAnalysisEngine.CreateColumnAccumulators(headers);
                    colCount = accs.Count;

                    List<string> normHeaders = new List<string>(colCount);
                    for (int i = 0; i < colCount; i++) normHeaders.Add(accs[i].ColumnName);
                    List<string> pkSrc = (explicitPkCols != null && explicitPkCols.Count > 0)
                        ? explicitPkCols
                        : (wirePriCols.Count > 0 ? wirePriCols : null);
                    string keyLabel;
                    int[] keyCols = ContentHashHelper.SelectSketchKeyColumns(normHeaders, pkSrc, out keyLabel);
                    sketchBuilder = new RowChurnSketchBuilder(keyCols, keyLabel, pkSrc != null && pkSrc.Count > 0, RowChurnSketchBuilder.DefaultCap);
                    hash1 = new ulong[colCount];
                    hash2 = new ulong[colCount];
                    hashed = new bool[colCount];
                },
                delegate(string[] rowVals, long rawRowBytes, long csvLineBytes)
                {
                    CsvAnalysisEngine.RowHashAccum rh = CsvAnalysisEngine.AccumulateRow(rowVals, accs, hash1, hash2, hashed, sketchBuilder);
                    unchecked
                    {
                        rowHashSum1 += rh.H1;
                        rowHashSum2 += rh.H2;
                        rowHashXor1 ^= rh.H1;
                        rowHashXor2 ^= rh.H2;
                    }
                    sumRowBytes += rawRowBytes;
                    if (rawRowBytes < minRowBytes) minRowBytes = rawRowBytes;
                    if (rawRowBytes > maxRowBytes) maxRowBytes = rawRowBytes;
                    csvEquivBytes += csvLineBytes;
                    rowCount++;
                },
                cancelState,
                onRowProgress);

            if (cancelState != null) cancelState.ThrowIfCancelled();
            ta.FileSizeBytes = receivedColumns ? csvEquivBytes : 0L;
            long effectiveMinRowBytes = rowCount > 0 && minRowBytes != long.MaxValue ? minRowBytes : 0L;
            CsvAnalysisEngine.FinalizeTableAnalysis(ta, accs, rowCount, 0L, sumRowBytes, effectiveMinRowBytes, maxRowBytes, tableKeyMaps);

            unchecked
            {
                ulong rcMix = ContentHashHelper.Mix64((ulong)rowCount ^ 0x517CC1B727220A95UL);
                ta.RowMultisetHash = ContentHashHelper.Format256(
                    rowHashSum1 ^ rcMix,
                    rowHashSum2 + rcMix,
                    rowHashXor1 ^ ContentHashHelper.Mix64((ulong)ta.ColumnCount),
                    rowHashXor2);
            }
            if (sketchBuilder != null)
            {
                ta.RowSketch = sketchBuilder.Build(ta);
            }

            DbTableMetadata tm;
            if (_tableMetaMap.TryGetValue(tableName, out tm))
            {
                ta.DbMeta = tm;
            }
            if (tblColMeta != null)
            {
                for (int ci = 0; ci < ta.Columns.Count; ci++)
                {
                    DbColumnMetadata cm;
                    if (tblColMeta.TryGetValue(ta.Columns[ci].ColumnName, out cm))
                    {
                        ta.Columns[ci].DbMeta = cm;
                    }
                }
            }

            _result.Tables.Add(ta);
            _keyMaps[ta.TableName] = tableKeyMaps;
            return ta;
        }

        public FolderAnalysisResult Complete()
        {
            return Complete(null);
        }

        internal FolderAnalysisResult Complete(OperationCancelState cancelState)
        {
            if (_result.Tables.Count == 0)
            {
                throw new InvalidOperationException("No tables were analyzed from the database.");
            }
            _result.AnalyzedAt = DateTime.Now;
            CsvAnalysisEngine.FinalizeFolderAnalysis(_result, _keyMaps, cancelState);
            return _result;
        }
    }

    // =========================================================================
    //  Before / After Analysis Workbook Reader & Table Diff Comparison Engine
    // =========================================================================
    public class SnapshotCategoryItem
    {
        public string Value;
        public long Count;
        public double Share;
        public double CumulativeShare;
    }

    public class SnapshotNumericStats
    {
        public string ColumnName;
        public long NonNullCount;
        public double Min;
        public bool HasP05;
        public double P05;
        public bool HasP10;
        public double P10;
        public double P25;
        public double Median;
        public double Mean;
        public double P75;
        public bool HasP90;
        public double P90;
        public double P95;
        public double P99;
        public double Max;
        public double StdDev;
        public bool HasSum;
        public double Sum;
        public bool HasZeroRate;
        public double ZeroRate;
        public bool HasNegativeRate;
        public double NegativeRate;
        public double OutlierRate;
        public List<HistogramBin> Histogram;

        public SnapshotNumericStats()
        {
            Histogram = new List<HistogramBin>();
        }
    }

    public class SnapshotDateTimeStats
    {
        public string ColumnName;
        public long NonNullCount;
        public bool HasTimeComponent;
        public string MinDate;
        public string MaxDate;
        public double DateSpanDays;
        public double AvgPerDay;
        public string PeakDay;
        public long PeakDayCount;
        public double PeakBurstRatio;
        public SortedDictionary<string, long> MonthlyCounts;
        public SortedDictionary<string, long> RecentDailyCounts;
        public long[] DayOfWeekCounts;
        public long[] HourlyCounts;
        public bool HasTimeSeriesDist;

        public SnapshotDateTimeStats()
        {
            MinDate = "-";
            MaxDate = "-";
            PeakDay = "";
            MonthlyCounts = new SortedDictionary<string, long>(StringComparer.Ordinal);
            RecentDailyCounts = new SortedDictionary<string, long>(StringComparer.Ordinal);
            DayOfWeekCounts = new long[7];
            HourlyCounts = new long[24];
        }
    }

    public class SnapshotColumn
    {
        public string TableName;
        public string ColumnName;
        public int ColumnIndex;
        public string InferredType;
        public string KeyRole;
        public long TotalRows;
        public long NonNullCount;
        public long NullCount;
        public double NullRate;
        public long DistinctCount;
        public bool DistinctCapped;
        public double CardinalityRate;
        public string SelectivityRating;
        public bool IsCandidatePk;
        public int MinCharLen;
        public double AvgCharLen;
        public int MaxCharLen;
        public long MinByteLen;
        public double AvgByteLen;
        public long P95ByteLen;
        public long MaxByteLen;
        public string Top1Value;
        public long Top1Count;
        public double Top1Share;
        public double Top3Share;
        public double Top5Share;
        public double Top10Share;
        public double GiniCoefficient;
        public double NormalizedEntropy;
        public bool HasNormalizedEntropy;
        public string SkewRisk;
        public bool IsCategorical;
        public long WhitespaceOnlyCount;
        public long MultilineCount;
        public long NonAsciiCount;
        public double NonAsciiRate;
        public bool HasCharQualityCounts;
        public string AlertNote;
        public string MinValueText;
        public string MaxValueText;
        public string DetailStatsSummary;
        public string RelationSummary;
        public string ContentHash;
        public string DistinctSetHash;
        public DbColumnMetadata DbMeta;
        public Dictionary<string, long> TopDistCounts;

        public SnapshotColumn()
        {
            InferredType = "";
            KeyRole = "";
            SelectivityRating = "";
            Top1Value = "";
            SkewRisk = "";
            AlertNote = "";
            MinValueText = "";
            MaxValueText = "";
            DetailStatsSummary = "";
            RelationSummary = "";
            ContentHash = "";
            DistinctSetHash = "";
            TopDistCounts = new Dictionary<string, long>(StringComparer.Ordinal);
        }
    }

    public class SnapshotTable
    {
        public string TableName;
        public string FileName;
        public string EncodingName;
        public string InferredRole;
        public long RowCount;
        public double RowShare;
        public int ColumnCount;
        public double FileSizeKB;
        public double SizeShare;
        public long MinRowBytes;
        public double AvgRowBytes;
        public long MaxRowBytes;
        public double EstInnoDbKB;
        public long EstInnoDbPages;
        public double RowsPerInnoDbPage;
        public double AvgNullRate;
        public double AvgCardinalityRate;
        public string ColumnTypeBreakdown;
        public string QualityRiskBreakdown;
        public long RaggedRowCount;
        public string CandidatePks;
        public string CategoricalColumns;
        public string DateColumns;
        public string PrimaryDateRange;
        public string RelationSummary;
        public string JMeterReadiness;
        public int WarningCount;
        public int HighWarningCount;
        public string WarningSummary;
        public string RowMultisetHash;
        public RowChurnSketch RowSketch;
        public DbTableMetadata DbMeta;

        public Dictionary<string, SnapshotColumn> Columns;
        public List<SnapshotColumn> ColumnList;
        public Dictionary<string, Dictionary<string, SnapshotCategoryItem>> Distributions;
        public Dictionary<string, SnapshotNumericStats> NumericStats;
        public Dictionary<string, SnapshotDateTimeStats> DateTimeStats;

        public SnapshotTable()
        {
            FileName = "";
            EncodingName = "";
            InferredRole = "";
            ColumnTypeBreakdown = "";
            QualityRiskBreakdown = "";
            CandidatePks = "";
            CategoricalColumns = "";
            DateColumns = "";
            PrimaryDateRange = "";
            RelationSummary = "";
            JMeterReadiness = "";
            WarningSummary = "";
            RowMultisetHash = "";
            Columns = new Dictionary<string, SnapshotColumn>(StringComparer.OrdinalIgnoreCase);
            ColumnList = new List<SnapshotColumn>();
            Distributions = new Dictionary<string, Dictionary<string, SnapshotCategoryItem>>(StringComparer.OrdinalIgnoreCase);
            NumericStats = new Dictionary<string, SnapshotNumericStats>(StringComparer.OrdinalIgnoreCase);
            DateTimeStats = new Dictionary<string, SnapshotDateTimeStats>(StringComparer.OrdinalIgnoreCase);
        }
    }

    public class SnapshotRelationship
    {
        public string ParentTable;
        public string ParentColumn;
        public string ChildTable;
        public string ChildColumn;
        public long ParentDistinctCount;
        public long ChildNonNullRows;
        public long ChildDistinctKeys;
        public long MatchedRows;
        public long OrphanRows;
        public double MatchRate;
        public long FanOutMin;
        public double FanOutAvg;
        public double FanOutP50;
        public double FanOutP95;
        public double FanOutP99;
        public long FanOutMax;
        public string MaxFanOutKey;
        public string RiskLevel;
        public string Note;

        public SnapshotRelationship()
        {
            ParentTable = "";
            ParentColumn = "";
            ChildTable = "";
            ChildColumn = "";
            MaxFanOutKey = "";
            RiskLevel = "";
            Note = "";
        }
    }

    public class SnapshotWarning
    {
        public string Severity;
        public string Category;
        public string TableName;
        public string ColumnName;
        public string MetricValue;
        public string PerfImpact;
        public string Recommendation;

        public SnapshotWarning()
        {
            Severity = "";
            Category = "";
            TableName = "";
            ColumnName = "";
            MetricValue = "";
            PerfImpact = "";
            Recommendation = "";
        }
    }

    public class SnapshotJMeterCandidate
    {
        public string TableName;
        public string KeyColumn;
        public long UniqueCount;
        public long TotalRows;
        public string DataType;
        public string Suitability;
        public string RecommendedUsage;

        public SnapshotJMeterCandidate()
        {
            TableName = "";
            KeyColumn = "";
            DataType = "";
            Suitability = "";
            RecommendedUsage = "";
        }
    }

    public class AnalysisWorkbookSnapshot
    {
        public string FilePath;
        public string FileName;
        public string Subtitle;
        public string AnalyzedAtText;
        public DateTime? AnalyzedAt;
        public string InputFolder;
        public string SourceMode;
        public bool IsDirectDb;
        public string DbLabel;
        public bool HasEmbeddedSnapshot;
        public int SnapshotVersion;
        public double EstMinAcuCurrent;
        public double EstMinAcu10x;
        public bool HasAcuEstimate;
        public long TotalEstInnoDbPages;
        public long TotalOrphanRows;
        public string PrimaryTimeSeriesSource;
        public DbServerMetadata DbServer;

        public Dictionary<string, SnapshotTable> Tables;
        public List<SnapshotTable> TableList;
        public List<SnapshotRelationship> Relationships;
        public List<SnapshotWarning> Warnings;
        public List<SnapshotJMeterCandidate> JMeterCandidates;

        public AnalysisWorkbookSnapshot()
        {
            Subtitle = "";
            AnalyzedAtText = "";
            InputFolder = "";
            SourceMode = "";
            DbLabel = "";
            PrimaryTimeSeriesSource = "";
            Tables = new Dictionary<string, SnapshotTable>(StringComparer.OrdinalIgnoreCase);
            TableList = new List<SnapshotTable>();
            Relationships = new List<SnapshotRelationship>();
            Warnings = new List<SnapshotWarning>();
            JMeterCandidates = new List<SnapshotJMeterCandidate>();
        }
    }

    // =========================================================================
    //  Lossless OOXML CustomXML Snapshot Codec (Standard customXml/item1.xml)
    // =========================================================================
    public static class AnalysisSnapshotCodec
    {
        public const string SnapshotNamespace = "urn:aurora-csv-analyzer:snapshot:v2";
        private const int MaxTotalSketchEntries = 32768;

        private static string CleanXml(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            StringBuilder sb = null;
            for (int i = 0; i < s.Length; i++)
            {
                char ch = s[i];
                if (ch == '\r')
                {
                    if (sb == null) { sb = new StringBuilder(s.Length); sb.Append(s, 0, i); }
                    if (i + 1 < s.Length && s[i + 1] == '\n') i++;
                    sb.Append('\n');
                    continue;
                }
                bool valid = (ch == 0x9 || ch == 0xA || (ch >= 0x20 && ch <= 0xD7FF) || (ch >= 0xE000 && ch <= 0xFFFD));
                if (!valid && char.IsHighSurrogate(ch) && i + 1 < s.Length && char.IsLowSurrogate(s[i + 1]))
                {
                    if (sb != null) { sb.Append(ch); sb.Append(s[i + 1]); }
                    i++;
                    continue;
                }
                if (!valid)
                {
                    if (sb == null) { sb = new StringBuilder(s.Length); sb.Append(s, 0, i); }
                    continue;
                }
                if (sb != null) sb.Append(ch);
            }
            return sb != null ? sb.ToString() : s;
        }

        private static string FmtD(double v)
        {
            if (double.IsNaN(v) || double.IsInfinity(v)) return "0";
            return v.ToString("R", CultureInfo.InvariantCulture);
        }

        private static string FmtL(long v)
        {
            return v.ToString(CultureInfo.InvariantCulture);
        }

        private static string FmtUL(ulong v)
        {
            return v.ToString(CultureInfo.InvariantCulture);
        }

        private static void WriteAttrIf(XmlWriter xw, string name, string val)
        {
            if (!string.IsNullOrEmpty(val)) xw.WriteAttributeString(name, CleanXml(val));
        }

        public static string SerializeToXml(FolderAnalysisResult analysis, string dbLabel)
        {
            if (analysis == null) return "";

            // Adaptively budget sketch entries across all tables so embedded XML remains lightweight
            int totalSketchEntries = 0;
            for (int i = 0; i < analysis.Tables.Count; i++)
            {
                RowChurnSketch sk = analysis.Tables[i].RowSketch;
                if (sk != null && sk.IsValid && sk.Entries != null) totalSketchEntries += sk.Entries.Count;
            }
            int perTableCap = RowChurnSketchBuilder.DefaultCap;
            if (totalSketchEntries > MaxTotalSketchEntries && analysis.Tables.Count > 0)
            {
                perTableCap = Math.Max(64, MaxTotalSketchEntries / analysis.Tables.Count);
            }

            XmlWriterSettings ws = new XmlWriterSettings();
            ws.Encoding = new UTF8Encoding(false);
            ws.Indent = false;
            ws.OmitXmlDeclaration = false;
            ws.CheckCharacters = false;

            using (MemoryStream ms = new MemoryStream(65536))
            {
                using (XmlWriter xw = XmlWriter.Create(ms, ws))
                {
                    xw.WriteStartDocument();
                    xw.WriteStartElement("AuroraAnalysisSnapshot", SnapshotNamespace);
                    xw.WriteAttributeString("version", "2");
                    xw.WriteAttributeString("sourceMode", CleanXml(analysis.SourceMode ?? "CSV"));
                    xw.WriteAttributeString("isDirectDb", analysis.IsDirectDb ? "1" : "0");
                    xw.WriteAttributeString("analyzedAt", analysis.AnalyzedAt.ToString("yyyy-MM-ddTHH:mm:ss.fff", CultureInfo.InvariantCulture));
                    xw.WriteAttributeString("analyzedAtDisplay", analysis.AnalyzedAt.ToString("yyyy/MM/dd HH:mm:ss", CultureInfo.InvariantCulture));
                    WriteAttrIf(xw, "inputFolder", analysis.InputFolder);
                    WriteAttrIf(xw, "dbLabel", dbLabel);
                    xw.WriteAttributeString("tableCount", FmtL(analysis.TableCount));
                    xw.WriteAttributeString("totalRows", FmtL(analysis.TotalRows));
                    xw.WriteAttributeString("totalColumns", FmtL(analysis.TotalColumns));
                    xw.WriteAttributeString("totalFileBytes", FmtL(analysis.TotalFileBytes));
                    xw.WriteAttributeString("totalEstInnoDbBytes", FmtL(analysis.TotalEstInnoDbBytes));
                    xw.WriteAttributeString("totalEstInnoDbPages", FmtL(analysis.TotalEstInnoDbPages));
                    xw.WriteAttributeString("totalOrphanRows", FmtL(analysis.TotalOrphanRows));
                    xw.WriteAttributeString("estMinAcuCurrent", FmtD(analysis.EstMinAcuCurrent));
                    xw.WriteAttributeString("estMinAcu10x", FmtD(analysis.EstMinAcu10x));
                    WriteAttrIf(xw, "primaryTsSource", analysis.PrimaryTimeSeriesSource);

                    if (analysis.DbServer != null && (analysis.DbServer.HasVariables || analysis.DbServer.HasStatus))
                    {
                        xw.WriteStartElement("DbServer", SnapshotNamespace);
                        if (analysis.DbServer.HasVariables)
                        {
                            xw.WriteStartElement("Variables", SnapshotNamespace);
                            foreach (KeyValuePair<string, string> kv in analysis.DbServer.Variables)
                            {
                                xw.WriteStartElement("V", SnapshotNamespace);
                                xw.WriteAttributeString("k", CleanXml(kv.Key));
                                xw.WriteAttributeString("v", CleanXml(kv.Value));
                                xw.WriteEndElement();
                            }
                            xw.WriteEndElement();
                        }
                        if (analysis.DbServer.HasStatus)
                        {
                            xw.WriteStartElement("Status", SnapshotNamespace);
                            foreach (KeyValuePair<string, long> kv in analysis.DbServer.StatusCounters)
                            {
                                xw.WriteStartElement("S", SnapshotNamespace);
                                xw.WriteAttributeString("k", CleanXml(kv.Key));
                                xw.WriteAttributeString("v", FmtL(kv.Value));
                                xw.WriteEndElement();
                            }
                            xw.WriteEndElement();
                        }
                        xw.WriteEndElement();
                    }

                    xw.WriteStartElement("Tables", SnapshotNamespace);
                    for (int ti = 0; ti < analysis.Tables.Count; ti++)
                    {
                        TableAnalysis t = analysis.Tables[ti];
                        xw.WriteStartElement("Table", SnapshotNamespace);
                        xw.WriteAttributeString("name", CleanXml(t.TableName));
                        WriteAttrIf(xw, "file", t.FileName);
                        WriteAttrIf(xw, "enc", t.EncodingName);
                        WriteAttrIf(xw, "role", t.InferredRole);
                        xw.WriteAttributeString("rows", FmtL(t.RowCount));
                        xw.WriteAttributeString("rowShare", FmtD(t.RowShare));
                        xw.WriteAttributeString("cols", FmtL(t.ColumnCount));
                        xw.WriteAttributeString("sizeKB", FmtD(Math.Round(t.FileSizeBytes / 1024.0, 4)));
                        xw.WriteAttributeString("sizeShare", FmtD(t.SizeShare));
                        xw.WriteAttributeString("minRowB", FmtL(t.MinRowBytes));
                        xw.WriteAttributeString("avgRowB", FmtD(t.AvgRowBytes));
                        xw.WriteAttributeString("maxRowB", FmtL(t.MaxRowBytes));
                        xw.WriteAttributeString("innoKB", FmtD(Math.Round(t.EstInnoDbBytes / 1024.0, 2)));
                        xw.WriteAttributeString("innoPages", FmtL(t.EstInnoDbPages));
                        xw.WriteAttributeString("rowsPerPg", FmtD(t.RowsPerInnoDbPage));
                        xw.WriteAttributeString("avgNull", FmtD(t.AvgNullRate));
                        xw.WriteAttributeString("avgCard", FmtD(t.AvgCardinalityRate));
                        WriteAttrIf(xw, "typeBrk", t.ColumnTypeBreakdown);
                        WriteAttrIf(xw, "qRisk", t.QualityRiskBreakdown);
                        xw.WriteAttributeString("ragged", FmtL(t.RaggedRowCount));
                        WriteAttrIf(xw, "pks", t.CandidatePks);
                        WriteAttrIf(xw, "catCols", t.CategoricalColumns);
                        WriteAttrIf(xw, "dtCols", t.DateColumns);
                        WriteAttrIf(xw, "dtRange", t.PrimaryDateRange);
                        WriteAttrIf(xw, "relSum", t.RelationSummary);
                        WriteAttrIf(xw, "jmReady", t.JMeterReadiness);
                        xw.WriteAttributeString("warnCnt", FmtL(t.WarningCount));
                        xw.WriteAttributeString("hiWarnCnt", FmtL(t.HighWarningCount));
                        WriteAttrIf(xw, "warnSum", t.WarningSummary);
                        WriteAttrIf(xw, "rowHash", t.RowMultisetHash);

                        if (t.DbMeta != null && (t.DbMeta.HasMetadata || t.DbMeta.Indexes.Count > 0 || t.DbMeta.DeclaredFks.Count > 0))
                        {
                            xw.WriteStartElement("DbMeta", SnapshotNamespace);
                            xw.WriteAttributeString("hasMeta", t.DbMeta.HasMetadata ? "1" : "0");
                            WriteAttrIf(xw, "engine", t.DbMeta.Engine);
                            WriteAttrIf(xw, "rowFmt", t.DbMeta.RowFormat);
                            xw.WriteAttributeString("estRows", FmtL(t.DbMeta.EstTableRows));
                            xw.WriteAttributeString("avgLen", FmtL(t.DbMeta.AvgRowLength));
                            xw.WriteAttributeString("dataLen", FmtL(t.DbMeta.DataLength));
                            xw.WriteAttributeString("idxLen", FmtL(t.DbMeta.IndexLength));
                            xw.WriteAttributeString("dataFree", FmtL(t.DbMeta.DataFree));
                            if (t.DbMeta.HasAutoIncrement)
                            {
                                xw.WriteAttributeString("hasAi", "1");
                                xw.WriteAttributeString("ai", FmtL(t.DbMeta.AutoIncrement));
                            }
                            WriteAttrIf(xw, "ctime", t.DbMeta.CreateTime);
                            WriteAttrIf(xw, "utime", t.DbMeta.UpdateTime);
                            WriteAttrIf(xw, "collation", t.DbMeta.TableCollation);
                            WriteAttrIf(xw, "comment", t.DbMeta.TableComment);
                            if (t.DbMeta.Indexes.Count > 0)
                            {
                                xw.WriteStartElement("Indexes", SnapshotNamespace);
                                for (int ii = 0; ii < t.DbMeta.Indexes.Count; ii++)
                                {
                                    DbIndexMetadata im = t.DbMeta.Indexes[ii];
                                    xw.WriteStartElement("Idx", SnapshotNamespace);
                                    xw.WriteAttributeString("name", CleanXml(im.IndexName));
                                    xw.WriteAttributeString("nonUniq", im.NonUnique ? "1" : "0");
                                    WriteAttrIf(xw, "type", im.IndexType);
                                    WriteAttrIf(xw, "cols", im.Columns);
                                    xw.WriteAttributeString("card", FmtL(im.Cardinality));
                                    xw.WriteEndElement();
                                }
                                xw.WriteEndElement();
                            }
                            if (t.DbMeta.DeclaredFks.Count > 0)
                            {
                                xw.WriteStartElement("Fks", SnapshotNamespace);
                                for (int fi = 0; fi < t.DbMeta.DeclaredFks.Count; fi++)
                                {
                                    DbForeignKeyMetadata fk = t.DbMeta.DeclaredFks[fi];
                                    xw.WriteStartElement("Fk", SnapshotNamespace);
                                    xw.WriteAttributeString("name", CleanXml(fk.ConstraintName));
                                    xw.WriteAttributeString("col", CleanXml(fk.ColumnName));
                                    xw.WriteAttributeString("refT", CleanXml(fk.ReferencedTableName));
                                    xw.WriteAttributeString("refC", CleanXml(fk.ReferencedColumnName));
                                    xw.WriteEndElement();
                                }
                                xw.WriteEndElement();
                            }
                            xw.WriteEndElement();
                        }

                        if (t.RowSketch != null && t.RowSketch.IsValid && t.RowSketch.Entries != null && t.RowSketch.Entries.Count > 0)
                        {
                            int takeN = Math.Min(t.RowSketch.Entries.Count, perTableCap);
                            bool effExact = t.RowSketch.IsExact && (takeN == t.RowSketch.Entries.Count);
                            ulong effMaxPk = effExact ? ulong.MaxValue : t.RowSketch.Entries[takeN - 1].PkHash;
                            xw.WriteStartElement("RowSketch", SnapshotNamespace);
                            xw.WriteAttributeString("keys", CleanXml(t.RowSketch.KeyColumns));
                            xw.WriteAttributeString("exact", effExact ? "1" : "0");
                            xw.WriteAttributeString("total", FmtL(t.RowSketch.TotalRows));
                            xw.WriteAttributeString("maxPk", FmtUL(effMaxPk));
                            xw.WriteString(ContentHashHelper.EncodeSketchEntries(t.RowSketch.Entries, takeN));
                            xw.WriteEndElement();
                        }

                        xw.WriteStartElement("Columns", SnapshotNamespace);
                        for (int ci = 0; ci < t.Columns.Count; ci++)
                        {
                            ColumnAnalysis c = t.Columns[ci];
                            xw.WriteStartElement("Col", SnapshotNamespace);
                            xw.WriteAttributeString("idx", FmtL(c.ColumnIndex));
                            xw.WriteAttributeString("name", CleanXml(c.ColumnName));
                            xw.WriteAttributeString("type", CleanXml(c.InferredType));
                            WriteAttrIf(xw, "role", c.KeyRole);
                            xw.WriteAttributeString("rows", FmtL(c.TotalRows));
                            xw.WriteAttributeString("nn", FmtL(c.NonNullCount));
                            xw.WriteAttributeString("nulls", FmtL(c.NullCount));
                            xw.WriteAttributeString("nullR", FmtD(c.NullRate));
                            xw.WriteAttributeString("dist", FmtL(c.DistinctCount));
                            xw.WriteAttributeString("distCap", c.DistinctCapped ? "1" : "0");
                            xw.WriteAttributeString("cardR", FmtD(c.CardinalityRate));
                            WriteAttrIf(xw, "sel", c.SelectivityRating);
                            xw.WriteAttributeString("isPk", c.IsCandidatePk ? "1" : "0");
                            xw.WriteAttributeString("minC", FmtL(c.MinCharLen));
                            xw.WriteAttributeString("avgC", FmtD(c.AvgCharLen));
                            xw.WriteAttributeString("maxC", FmtL(c.MaxCharLen));
                            xw.WriteAttributeString("minB", FmtL(c.MinByteLen));
                            xw.WriteAttributeString("avgB", FmtD(c.AvgByteLen));
                            xw.WriteAttributeString("p95B", FmtL(c.P95ByteLen));
                            xw.WriteAttributeString("maxB", FmtL(c.MaxByteLen));
                            WriteAttrIf(xw, "top1", c.Top1Value);
                            xw.WriteAttributeString("top1C", FmtL(c.Top1Count));
                            xw.WriteAttributeString("top1S", FmtD(c.Top1Share));
                            xw.WriteAttributeString("top3S", FmtD(c.Top3Share));
                            xw.WriteAttributeString("top5S", FmtD(c.Top5Share));
                            xw.WriteAttributeString("top10S", FmtD(c.Top10Share));
                            xw.WriteAttributeString("gini", FmtD(c.GiniCoefficient));
                            xw.WriteAttributeString("ent", FmtD(c.NormalizedEntropy));
                            WriteAttrIf(xw, "skew", c.SkewRisk);
                            xw.WriteAttributeString("isCat", c.IsCategorical ? "1" : "0");
                            xw.WriteAttributeString("ws", FmtL(c.WhitespaceOnlyCount));
                            xw.WriteAttributeString("ml", FmtL(c.MultilineCount));
                            xw.WriteAttributeString("na", FmtL(c.NonAsciiCount));
                            xw.WriteAttributeString("naR", FmtD(c.NonAsciiRate));
                            WriteAttrIf(xw, "minV", c.MinValueText == "-" ? "" : c.MinValueText);
                            WriteAttrIf(xw, "maxV", c.MaxValueText == "-" ? "" : c.MaxValueText);
                            WriteAttrIf(xw, "detSum", c.DetailStatsSummary == "-" ? "" : c.DetailStatsSummary);
                            WriteAttrIf(xw, "relSum", c.RelationSummary == "-" ? "" : c.RelationSummary);
                            WriteAttrIf(xw, "alert", c.AlertNote);
                            WriteAttrIf(xw, "cHash", c.ContentHash);
                            WriteAttrIf(xw, "dHash", c.DistinctSetHash);

                            if (c.DbMeta != null)
                            {
                                xw.WriteStartElement("DbCol", SnapshotNamespace);
                                xw.WriteAttributeString("pos", FmtL(c.DbMeta.OrdinalPosition));
                                WriteAttrIf(xw, "ddl", c.DbMeta.ColumnType);
                                WriteAttrIf(xw, "null", c.DbMeta.IsNullable);
                                WriteAttrIf(xw, "key", c.DbMeta.ColumnKey);
                                WriteAttrIf(xw, "def", c.DbMeta.ColumnDefault);
                                WriteAttrIf(xw, "extra", c.DbMeta.Extra);
                                WriteAttrIf(xw, "cs", c.DbMeta.CharacterSetName);
                                WriteAttrIf(xw, "coll", c.DbMeta.CollationName);
                                WriteAttrIf(xw, "cmt", c.DbMeta.ColumnComment);
                                xw.WriteEndElement();
                            }

                            if (c.TopValues != null && c.TopValues.Count > 0 && !c.IsCandidatePk)
                            {
                                xw.WriteStartElement("Top", SnapshotNamespace);
                                for (int vi = 0; vi < c.TopValues.Count; vi++)
                                {
                                    ValueShareItem tv = c.TopValues[vi];
                                    xw.WriteStartElement("I", SnapshotNamespace);
                                    xw.WriteAttributeString("v", CleanXml(tv.Value));
                                    xw.WriteAttributeString("c", FmtL(tv.Count));
                                    xw.WriteAttributeString("s", FmtD(tv.Share));
                                    xw.WriteAttributeString("cs", FmtD(tv.CumulativeShare));
                                    xw.WriteEndElement();
                                }
                                xw.WriteEndElement();
                            }

                            if (c.IsNumeric && c.NonNullCount > 0)
                            {
                                xw.WriteStartElement("Num", SnapshotNamespace);
                                xw.WriteAttributeString("nn", FmtL(c.NonNullCount));
                                xw.WriteAttributeString("min", FmtD(c.NumMin));
                                xw.WriteAttributeString("p05", FmtD(c.NumP05));
                                xw.WriteAttributeString("p10", FmtD(c.NumP10));
                                xw.WriteAttributeString("p25", FmtD(c.NumP25));
                                xw.WriteAttributeString("med", FmtD(c.NumMedian));
                                xw.WriteAttributeString("mean", FmtD(c.NumMean));
                                xw.WriteAttributeString("p75", FmtD(c.NumP75));
                                xw.WriteAttributeString("p90", FmtD(c.NumP90));
                                xw.WriteAttributeString("p95", FmtD(c.NumP95));
                                xw.WriteAttributeString("p99", FmtD(c.NumP99));
                                xw.WriteAttributeString("max", FmtD(c.NumMax));
                                xw.WriteAttributeString("std", FmtD(c.NumStdDev));
                                xw.WriteAttributeString("sum", FmtD(c.NumSum));
                                xw.WriteAttributeString("zR", FmtD(c.ZeroRate));
                                xw.WriteAttributeString("nR", FmtD(c.NegativeRate));
                                xw.WriteAttributeString("oR", FmtD(c.OutlierRate));
                                xw.WriteEndElement();
                            }

                            if (c.IsDateTime && c.NonNullCount > 0 && c.MinDate.HasValue && c.MaxDate.HasValue)
                            {
                                xw.WriteStartElement("Dt", SnapshotNamespace);
                                xw.WriteAttributeString("nn", FmtL(c.NonNullCount));
                                xw.WriteAttributeString("hasT", c.HasTimeComponent ? "1" : "0");
                                xw.WriteAttributeString("min", c.MinDate.Value.ToString("yyyy/MM/dd HH:mm:ss", CultureInfo.InvariantCulture));
                                xw.WriteAttributeString("max", c.MaxDate.Value.ToString("yyyy/MM/dd HH:mm:ss", CultureInfo.InvariantCulture));
                                xw.WriteAttributeString("span", FmtD(c.DateSpanDays));
                                xw.WriteAttributeString("avg", FmtD(c.AvgPerDay));
                                WriteAttrIf(xw, "pkD", c.PeakDay);
                                xw.WriteAttributeString("pkC", FmtL(c.PeakDayCount));
                                xw.WriteAttributeString("pkR", FmtD(c.PeakBurstRatio));
                                if (c.MonthlyCounts != null && c.MonthlyCounts.Count > 0)
                                {
                                    xw.WriteStartElement("M", SnapshotNamespace);
                                    foreach (KeyValuePair<string, long> mk in c.MonthlyCounts)
                                    {
                                        xw.WriteStartElement("E", SnapshotNamespace);
                                        xw.WriteAttributeString("k", CleanXml(mk.Key));
                                        xw.WriteAttributeString("v", FmtL(mk.Value));
                                        xw.WriteEndElement();
                                    }
                                    xw.WriteEndElement();
                                }
                                if (c.RecentDailyCounts != null && c.RecentDailyCounts.Count > 0)
                                {
                                    xw.WriteStartElement("D", SnapshotNamespace);
                                    foreach (KeyValuePair<string, long> dk in c.RecentDailyCounts)
                                    {
                                        xw.WriteStartElement("E", SnapshotNamespace);
                                        xw.WriteAttributeString("k", CleanXml(dk.Key));
                                        xw.WriteAttributeString("v", FmtL(dk.Value));
                                        xw.WriteEndElement();
                                    }
                                    xw.WriteEndElement();
                                }
                                if (c.DayOfWeekCounts != null && c.DayOfWeekCounts.Length == 7)
                                {
                                    xw.WriteElementString("Dow", SnapshotNamespace, string.Join(",", c.DayOfWeekCounts.Select(x => FmtL(x)).ToArray()));
                                }
                                if (c.HourlyCounts != null && c.HourlyCounts.Length == 24)
                                {
                                    xw.WriteElementString("Hr", SnapshotNamespace, string.Join(",", c.HourlyCounts.Select(x => FmtL(x)).ToArray()));
                                }
                                xw.WriteEndElement();
                            }

                            xw.WriteEndElement(); // Col
                        }
                        xw.WriteEndElement(); // Columns
                        xw.WriteEndElement(); // Table
                    }
                    xw.WriteEndElement(); // Tables

                    if (analysis.Relationships != null && analysis.Relationships.Count > 0)
                    {
                        xw.WriteStartElement("Relationships", SnapshotNamespace);
                        for (int ri = 0; ri < analysis.Relationships.Count; ri++)
                        {
                            RelationshipAnalysis rl = analysis.Relationships[ri];
                            xw.WriteStartElement("Rel", SnapshotNamespace);
                            xw.WriteAttributeString("pT", CleanXml(rl.ParentTable));
                            xw.WriteAttributeString("pC", CleanXml(rl.ParentColumn));
                            xw.WriteAttributeString("cT", CleanXml(rl.ChildTable));
                            xw.WriteAttributeString("cC", CleanXml(rl.ChildColumn));
                            xw.WriteAttributeString("pD", FmtL(rl.ParentDistinctCount));
                            xw.WriteAttributeString("cN", FmtL(rl.ChildNonNullRows));
                            xw.WriteAttributeString("cD", FmtL(rl.ChildDistinctKeys));
                            xw.WriteAttributeString("mR", FmtL(rl.MatchedRows));
                            xw.WriteAttributeString("oR", FmtL(rl.OrphanRows));
                            xw.WriteAttributeString("rate", FmtD(rl.MatchRate));
                            xw.WriteAttributeString("fMin", FmtL(rl.FanOutMin));
                            xw.WriteAttributeString("fAvg", FmtD(rl.FanOutAvg));
                            xw.WriteAttributeString("fP50", FmtD(rl.FanOutP50));
                            xw.WriteAttributeString("fP95", FmtD(rl.FanOutP95));
                            xw.WriteAttributeString("fP99", FmtD(rl.FanOutP99));
                            xw.WriteAttributeString("fMax", FmtL(rl.FanOutMax));
                            WriteAttrIf(xw, "maxK", rl.MaxFanOutKey);
                            WriteAttrIf(xw, "risk", rl.RiskLevel);
                            WriteAttrIf(xw, "note", rl.Note);
                            xw.WriteEndElement();
                        }
                        xw.WriteEndElement();
                    }

                    if (analysis.Warnings != null && analysis.Warnings.Count > 0)
                    {
                        xw.WriteStartElement("Warnings", SnapshotNamespace);
                        for (int wi = 0; wi < analysis.Warnings.Count; wi++)
                        {
                            WarningItem w = analysis.Warnings[wi];
                            xw.WriteStartElement("W", SnapshotNamespace);
                            xw.WriteAttributeString("sev", CleanXml(w.Severity));
                            xw.WriteAttributeString("cat", CleanXml(w.Category));
                            xw.WriteAttributeString("tbl", CleanXml(w.TableName));
                            xw.WriteAttributeString("col", CleanXml(w.ColumnName));
                            WriteAttrIf(xw, "met", w.MetricValue);
                            WriteAttrIf(xw, "imp", w.PerfImpact);
                            WriteAttrIf(xw, "rec", w.Recommendation);
                            xw.WriteEndElement();
                        }
                        xw.WriteEndElement();
                    }

                    if (analysis.JMeterCandidates != null && analysis.JMeterCandidates.Count > 0)
                    {
                        xw.WriteStartElement("JMeter", SnapshotNamespace);
                        for (int ji = 0; ji < analysis.JMeterCandidates.Count; ji++)
                        {
                            JMeterCandidateItem jc = analysis.JMeterCandidates[ji];
                            xw.WriteStartElement("J", SnapshotNamespace);
                            xw.WriteAttributeString("tbl", CleanXml(jc.TableName));
                            xw.WriteAttributeString("col", CleanXml(jc.KeyColumn));
                            xw.WriteAttributeString("uniq", FmtL(jc.UniqueCount));
                            xw.WriteAttributeString("tot", FmtL(jc.TotalRows));
                            WriteAttrIf(xw, "type", jc.DataType);
                            WriteAttrIf(xw, "suit", jc.Suitability);
                            WriteAttrIf(xw, "use", jc.RecommendedUsage);
                            xw.WriteEndElement();
                        }
                        xw.WriteEndElement();
                    }

                    xw.WriteEndElement(); // AuroraAnalysisSnapshot
                    xw.WriteEndDocument();
                }
                return Encoding.UTF8.GetString(ms.ToArray());
            }
        }

        private static string Attr(XmlNode node, string name)
        {
            if (node == null || node.Attributes == null) return "";
            XmlAttribute a = node.Attributes[name];
            if (a == null || string.IsNullOrEmpty(a.Value)) return "";
            string v = a.Value;
            if (v.IndexOf('\r') >= 0) v = v.Replace("\r\n", "\n").Replace("\r", "\n");
            return v;
        }

        private static long AttrL(XmlNode node, string name)
        {
            string s = Attr(node, name);
            long v;
            return long.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out v) ? v : 0L;
        }

        private static ulong AttrUL(XmlNode node, string name)
        {
            string s = Attr(node, name);
            ulong v;
            return ulong.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out v) ? v : 0UL;
        }

        private static double AttrD(XmlNode node, string name)
        {
            string s = Attr(node, name);
            double v;
            return double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out v) ? v : 0.0;
        }

        private static long[] ParseCsvLongs(string csv, int expectedLen)
        {
            long[] arr = new long[expectedLen];
            if (string.IsNullOrWhiteSpace(csv)) return arr;
            string[] parts = csv.Split(',');
            for (int i = 0; i < parts.Length && i < expectedLen; i++)
            {
                long v;
                if (long.TryParse(parts[i].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out v)) arr[i] = v;
            }
            return arr;
        }

        public static bool TryLoadFromZip(ZipArchive zip, AnalysisWorkbookSnapshot snap)
        {
            if (zip == null || snap == null) return false;
            foreach (ZipArchiveEntry entry in zip.Entries)
            {
                string norm = entry.FullName.Replace('\\', '/').TrimStart('/');
                if (!norm.StartsWith("customXml/item", StringComparison.OrdinalIgnoreCase) ||
                    !norm.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) ||
                    norm.IndexOf("itemProps", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    continue;
                }
                try
                {
                    XmlDocument doc = new XmlDocument();
                    doc.XmlResolver = null;
                    XmlReaderSettings rs = new XmlReaderSettings();
                    rs.DtdProcessing = DtdProcessing.Prohibit;
                    rs.CheckCharacters = false;
                    using (Stream s = entry.Open())
                    using (XmlReader xr = XmlReader.Create(s, rs))
                    {
                        doc.Load(xr);
                    }
                    XmlElement root = doc.DocumentElement;
                    if (root == null || root.LocalName != "AuroraAnalysisSnapshot") continue;
                    PopulateSnapshotFromXml(root, snap);
                    return true;
                }
                catch
                {
                    // Skip non-matching customXml parts
                }
            }
            return false;
        }

        private static void PopulateSnapshotFromXml(XmlElement root, AnalysisWorkbookSnapshot snap)
        {
            snap.HasEmbeddedSnapshot = true;
            snap.SnapshotVersion = (int)AttrL(root, "version");
            snap.SourceMode = Attr(root, "sourceMode");
            snap.IsDirectDb = Attr(root, "isDirectDb") == "1";
            snap.DbLabel = Attr(root, "dbLabel");
            string inp = Attr(root, "inputFolder");
            if (!string.IsNullOrEmpty(inp)) snap.InputFolder = inp;
            string dispAt = Attr(root, "analyzedAtDisplay");
            if (!string.IsNullOrEmpty(dispAt)) snap.AnalyzedAtText = dispAt;
            string isoAt = Attr(root, "analyzedAt");
            DateTime parsedDt;
            if (!string.IsNullOrEmpty(isoAt) && DateTime.TryParseExact(isoAt, new string[] { "yyyy-MM-ddTHH:mm:ss.fff", "yyyy-MM-ddTHH:mm:ss" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDt))
            {
                snap.AnalyzedAt = parsedDt;
            }
            snap.TotalEstInnoDbPages = AttrL(root, "totalEstInnoDbPages");
            snap.TotalOrphanRows = AttrL(root, "totalOrphanRows");
            snap.EstMinAcuCurrent = AttrD(root, "estMinAcuCurrent");
            snap.EstMinAcu10x = AttrD(root, "estMinAcu10x");
            snap.HasAcuEstimate = snap.EstMinAcuCurrent > 0.0;
            snap.PrimaryTimeSeriesSource = Attr(root, "primaryTsSource");

            foreach (XmlNode child in root.ChildNodes)
            {
                if (child.NodeType != XmlNodeType.Element) continue;
                if (child.LocalName == "DbServer")
                {
                    DbServerMetadata srv = new DbServerMetadata();
                    foreach (XmlNode sub in child.ChildNodes)
                    {
                        if (sub.NodeType != XmlNodeType.Element) continue;
                        if (sub.LocalName == "Variables")
                        {
                            foreach (XmlNode vn in sub.ChildNodes)
                            {
                                if (vn.NodeType != XmlNodeType.Element) continue;
                                string k = Attr(vn, "k");
                                if (!string.IsNullOrEmpty(k)) srv.Variables[k] = Attr(vn, "v");
                            }
                            if (srv.Variables.Count > 0) srv.HasVariables = true;
                        }
                        else if (sub.LocalName == "Status")
                        {
                            foreach (XmlNode sn in sub.ChildNodes)
                            {
                                if (sn.NodeType != XmlNodeType.Element) continue;
                                string k = Attr(sn, "k");
                                if (!string.IsNullOrEmpty(k)) srv.StatusCounters[k] = AttrL(sn, "v");
                            }
                            if (srv.StatusCounters.Count > 0) srv.HasStatus = true;
                        }
                    }
                    snap.DbServer = srv;
                }
                else if (child.LocalName == "Tables")
                {
                    foreach (XmlNode tn in child.ChildNodes)
                    {
                        if (tn.NodeType != XmlNodeType.Element || tn.LocalName != "Table") continue;
                        SnapshotTable st = new SnapshotTable();
                        st.TableName = Attr(tn, "name");
                        st.FileName = Attr(tn, "file");
                        st.EncodingName = Attr(tn, "enc");
                        st.InferredRole = Attr(tn, "role");
                        st.RowCount = AttrL(tn, "rows");
                        st.RowShare = AttrD(tn, "rowShare");
                        st.ColumnCount = (int)AttrL(tn, "cols");
                        st.FileSizeKB = AttrD(tn, "sizeKB");
                        st.SizeShare = AttrD(tn, "sizeShare");
                        st.MinRowBytes = AttrL(tn, "minRowB");
                        st.AvgRowBytes = AttrD(tn, "avgRowB");
                        st.MaxRowBytes = AttrL(tn, "maxRowB");
                        st.EstInnoDbKB = AttrD(tn, "innoKB");
                        st.EstInnoDbPages = AttrL(tn, "innoPages");
                        st.RowsPerInnoDbPage = AttrD(tn, "rowsPerPg");
                        st.AvgNullRate = AttrD(tn, "avgNull");
                        st.AvgCardinalityRate = AttrD(tn, "avgCard");
                        st.ColumnTypeBreakdown = Attr(tn, "typeBrk");
                        st.QualityRiskBreakdown = Attr(tn, "qRisk");
                        st.RaggedRowCount = AttrL(tn, "ragged");
                        st.CandidatePks = Attr(tn, "pks");
                        st.CategoricalColumns = Attr(tn, "catCols");
                        st.DateColumns = Attr(tn, "dtCols");
                        st.PrimaryDateRange = Attr(tn, "dtRange");
                        st.RelationSummary = Attr(tn, "relSum");
                        st.JMeterReadiness = Attr(tn, "jmReady");
                        st.WarningCount = (int)AttrL(tn, "warnCnt");
                        st.HighWarningCount = (int)AttrL(tn, "hiWarnCnt");
                        st.WarningSummary = Attr(tn, "warnSum");
                        st.RowMultisetHash = Attr(tn, "rowHash");

                        foreach (XmlNode tSub in tn.ChildNodes)
                        {
                            if (tSub.NodeType != XmlNodeType.Element) continue;
                            if (tSub.LocalName == "DbMeta")
                            {
                                DbTableMetadata dm = new DbTableMetadata();
                                dm.HasMetadata = Attr(tSub, "hasMeta") == "1";
                                dm.Engine = Attr(tSub, "engine");
                                dm.RowFormat = Attr(tSub, "rowFmt");
                                dm.EstTableRows = AttrL(tSub, "estRows");
                                dm.AvgRowLength = AttrL(tSub, "avgLen");
                                dm.DataLength = AttrL(tSub, "dataLen");
                                dm.IndexLength = AttrL(tSub, "idxLen");
                                dm.DataFree = AttrL(tSub, "dataFree");
                                dm.HasAutoIncrement = Attr(tSub, "hasAi") == "1";
                                dm.AutoIncrement = AttrL(tSub, "ai");
                                dm.CreateTime = Attr(tSub, "ctime");
                                dm.UpdateTime = Attr(tSub, "utime");
                                dm.TableCollation = Attr(tSub, "collation");
                                dm.TableComment = Attr(tSub, "comment");
                                foreach (XmlNode dSub in tSub.ChildNodes)
                                {
                                    if (dSub.NodeType != XmlNodeType.Element) continue;
                                    if (dSub.LocalName == "Indexes")
                                    {
                                        foreach (XmlNode ix in dSub.ChildNodes)
                                        {
                                            if (ix.NodeType != XmlNodeType.Element) continue;
                                            DbIndexMetadata im = new DbIndexMetadata();
                                            im.IndexName = Attr(ix, "name");
                                            im.NonUnique = Attr(ix, "nonUniq") == "1";
                                            im.IndexType = Attr(ix, "type");
                                            im.Columns = Attr(ix, "cols");
                                            im.Cardinality = AttrL(ix, "card");
                                            dm.Indexes.Add(im);
                                        }
                                    }
                                    else if (dSub.LocalName == "Fks")
                                    {
                                        foreach (XmlNode fn in dSub.ChildNodes)
                                        {
                                            if (fn.NodeType != XmlNodeType.Element) continue;
                                            DbForeignKeyMetadata fk = new DbForeignKeyMetadata();
                                            fk.ConstraintName = Attr(fn, "name");
                                            fk.ColumnName = Attr(fn, "col");
                                            fk.ReferencedTableName = Attr(fn, "refT");
                                            fk.ReferencedColumnName = Attr(fn, "refC");
                                            dm.DeclaredFks.Add(fk);
                                        }
                                    }
                                }
                                st.DbMeta = dm;
                            }
                            else if (tSub.LocalName == "RowSketch")
                            {
                                RowChurnSketch sk = new RowChurnSketch();
                                sk.KeyColumns = Attr(tSub, "keys");
                                sk.IsExact = Attr(tSub, "exact") == "1";
                                sk.TotalRows = AttrL(tSub, "total");
                                sk.MaxKeptPkHash = AttrUL(tSub, "maxPk");
                                sk.Entries = ContentHashHelper.DecodeSketchEntries(tSub.InnerText);
                                sk.IsValid = sk.Entries.Count > 0 && !string.IsNullOrEmpty(sk.KeyColumns);
                                st.RowSketch = sk;
                            }
                            else if (tSub.LocalName == "Columns")
                            {
                                foreach (XmlNode cn in tSub.ChildNodes)
                                {
                                    if (cn.NodeType != XmlNodeType.Element || cn.LocalName != "Col") continue;
                                    SnapshotColumn sc = new SnapshotColumn();
                                    sc.TableName = st.TableName;
                                    sc.ColumnIndex = (int)AttrL(cn, "idx");
                                    sc.ColumnName = Attr(cn, "name");
                                    sc.InferredType = Attr(cn, "type");
                                    sc.KeyRole = Attr(cn, "role");
                                    sc.TotalRows = AttrL(cn, "rows");
                                    sc.NonNullCount = AttrL(cn, "nn");
                                    sc.NullCount = AttrL(cn, "nulls");
                                    sc.NullRate = AttrD(cn, "nullR");
                                    sc.DistinctCount = AttrL(cn, "dist");
                                    sc.DistinctCapped = Attr(cn, "distCap") == "1";
                                    sc.CardinalityRate = AttrD(cn, "cardR");
                                    sc.SelectivityRating = Attr(cn, "sel");
                                    sc.IsCandidatePk = Attr(cn, "isPk") == "1";
                                    sc.MinCharLen = (int)AttrL(cn, "minC");
                                    sc.AvgCharLen = AttrD(cn, "avgC");
                                    sc.MaxCharLen = (int)AttrL(cn, "maxC");
                                    sc.MinByteLen = AttrL(cn, "minB");
                                    sc.AvgByteLen = AttrD(cn, "avgB");
                                    sc.P95ByteLen = AttrL(cn, "p95B");
                                    sc.MaxByteLen = AttrL(cn, "maxB");
                                    sc.Top1Value = Attr(cn, "top1");
                                    sc.Top1Count = AttrL(cn, "top1C");
                                    sc.Top1Share = AttrD(cn, "top1S");
                                    sc.Top3Share = AttrD(cn, "top3S");
                                    sc.Top5Share = AttrD(cn, "top5S");
                                    sc.Top10Share = AttrD(cn, "top10S");
                                    sc.GiniCoefficient = AttrD(cn, "gini");
                                    sc.NormalizedEntropy = AttrD(cn, "ent");
                                    sc.HasNormalizedEntropy = true;
                                    sc.SkewRisk = Attr(cn, "skew");
                                    sc.IsCategorical = Attr(cn, "isCat") == "1";
                                    sc.WhitespaceOnlyCount = AttrL(cn, "ws");
                                    sc.MultilineCount = AttrL(cn, "ml");
                                    sc.NonAsciiCount = AttrL(cn, "na");
                                    sc.NonAsciiRate = AttrD(cn, "naR");
                                    sc.HasCharQualityCounts = true;
                                    sc.MinValueText = Attr(cn, "minV");
                                    sc.MaxValueText = Attr(cn, "maxV");
                                    sc.DetailStatsSummary = Attr(cn, "detSum");
                                    sc.RelationSummary = Attr(cn, "relSum");
                                    sc.AlertNote = Attr(cn, "alert");
                                    sc.ContentHash = Attr(cn, "cHash");
                                    sc.DistinctSetHash = Attr(cn, "dHash");

                                    foreach (XmlNode cSub in cn.ChildNodes)
                                    {
                                        if (cSub.NodeType != XmlNodeType.Element) continue;
                                        if (cSub.LocalName == "DbCol")
                                        {
                                            DbColumnMetadata cm = new DbColumnMetadata();
                                            cm.ColumnName = sc.ColumnName;
                                            cm.OrdinalPosition = (int)AttrL(cSub, "pos");
                                            cm.ColumnType = Attr(cSub, "ddl");
                                            cm.IsNullable = Attr(cSub, "null");
                                            cm.ColumnKey = Attr(cSub, "key");
                                            cm.ColumnDefault = Attr(cSub, "def");
                                            cm.Extra = Attr(cSub, "extra");
                                            cm.CharacterSetName = Attr(cSub, "cs");
                                            cm.CollationName = Attr(cSub, "coll");
                                            cm.ColumnComment = Attr(cSub, "cmt");
                                            sc.DbMeta = cm;
                                        }
                                        else if (cSub.LocalName == "Top")
                                        {
                                            Dictionary<string, SnapshotCategoryItem> dist = new Dictionary<string, SnapshotCategoryItem>(StringComparer.Ordinal);
                                            int topIdx = 0;
                                            foreach (XmlNode itemN in cSub.ChildNodes)
                                            {
                                                if (itemN.NodeType != XmlNodeType.Element) continue;
                                                SnapshotCategoryItem ci = new SnapshotCategoryItem();
                                                ci.Value = Attr(itemN, "v");
                                                ci.Count = AttrL(itemN, "c");
                                                ci.Share = AttrD(itemN, "s");
                                                ci.CumulativeShare = AttrD(itemN, "cs");
                                                if (!dist.ContainsKey(ci.Value)) dist[ci.Value] = ci;
                                                if (topIdx < 5 && !(sc.NonNullCount > 0 && sc.DistinctCount == sc.NonNullCount && sc.DistinctCount > 10))
                                                {
                                                    string compactKey = (ci.Value ?? "").Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ").Trim();
                                                    if (compactKey.Length > 24) compactKey = compactKey.Substring(0, 21) + "...";
                                                    if (!string.IsNullOrEmpty(compactKey) && !sc.TopDistCounts.ContainsKey(compactKey))
                                                    {
                                                        sc.TopDistCounts[compactKey] = ci.Count;
                                                    }
                                                }
                                                topIdx++;
                                            }
                                            if (sc.IsCategorical && dist.Count > 0)
                                            {
                                                st.Distributions[sc.ColumnName] = dist;
                                            }
                                        }
                                        else if (cSub.LocalName == "Num")
                                        {
                                            SnapshotNumericStats ns = new SnapshotNumericStats();
                                            ns.ColumnName = sc.ColumnName;
                                            ns.NonNullCount = AttrL(cSub, "nn");
                                            ns.Min = AttrD(cSub, "min");
                                            ns.HasP05 = true; ns.P05 = AttrD(cSub, "p05");
                                            ns.HasP10 = true; ns.P10 = AttrD(cSub, "p10");
                                            ns.P25 = AttrD(cSub, "p25");
                                            ns.Median = AttrD(cSub, "med");
                                            ns.Mean = AttrD(cSub, "mean");
                                            ns.P75 = AttrD(cSub, "p75");
                                            ns.HasP90 = true; ns.P90 = AttrD(cSub, "p90");
                                            ns.P95 = AttrD(cSub, "p95");
                                            ns.P99 = AttrD(cSub, "p99");
                                            ns.Max = AttrD(cSub, "max");
                                            ns.StdDev = AttrD(cSub, "std");
                                            ns.HasSum = true; ns.Sum = AttrD(cSub, "sum");
                                            ns.HasZeroRate = true; ns.ZeroRate = AttrD(cSub, "zR");
                                            ns.HasNegativeRate = true; ns.NegativeRate = AttrD(cSub, "nR");
                                            ns.OutlierRate = AttrD(cSub, "oR");
                                            st.NumericStats[sc.ColumnName] = ns;
                                        }
                                        else if (cSub.LocalName == "Dt")
                                        {
                                            SnapshotDateTimeStats ds = new SnapshotDateTimeStats();
                                            ds.ColumnName = sc.ColumnName;
                                            ds.NonNullCount = AttrL(cSub, "nn");
                                            ds.HasTimeComponent = Attr(cSub, "hasT") == "1";
                                            ds.MinDate = Attr(cSub, "min");
                                            ds.MaxDate = Attr(cSub, "max");
                                            ds.DateSpanDays = AttrD(cSub, "span");
                                            ds.AvgPerDay = AttrD(cSub, "avg");
                                            ds.PeakDay = Attr(cSub, "pkD");
                                            ds.PeakDayCount = AttrL(cSub, "pkC");
                                            ds.PeakBurstRatio = AttrD(cSub, "pkR");
                                            foreach (XmlNode dtSub in cSub.ChildNodes)
                                            {
                                                if (dtSub.NodeType != XmlNodeType.Element) continue;
                                                if (dtSub.LocalName == "M")
                                                {
                                                    foreach (XmlNode en in dtSub.ChildNodes)
                                                    {
                                                        if (en.NodeType != XmlNodeType.Element) continue;
                                                        string k = Attr(en, "k");
                                                        if (!string.IsNullOrEmpty(k)) ds.MonthlyCounts[k] = AttrL(en, "v");
                                                    }
                                                    ds.HasTimeSeriesDist = true;
                                                }
                                                else if (dtSub.LocalName == "D")
                                                {
                                                    foreach (XmlNode en in dtSub.ChildNodes)
                                                    {
                                                        if (en.NodeType != XmlNodeType.Element) continue;
                                                        string k = Attr(en, "k");
                                                        if (!string.IsNullOrEmpty(k)) ds.RecentDailyCounts[k] = AttrL(en, "v");
                                                    }
                                                    ds.HasTimeSeriesDist = true;
                                                }
                                                else if (dtSub.LocalName == "Dow")
                                                {
                                                    ds.DayOfWeekCounts = ParseCsvLongs(dtSub.InnerText, 7);
                                                    ds.HasTimeSeriesDist = true;
                                                }
                                                else if (dtSub.LocalName == "Hr")
                                                {
                                                    ds.HourlyCounts = ParseCsvLongs(dtSub.InnerText, 24);
                                                    ds.HasTimeSeriesDist = true;
                                                }
                                            }
                                            st.DateTimeStats[sc.ColumnName] = ds;
                                        }
                                    }

                                    st.Columns[sc.ColumnName] = sc;
                                    st.ColumnList.Add(sc);
                                }
                            }
                        }

                        if (!string.IsNullOrEmpty(st.TableName) && !snap.Tables.ContainsKey(st.TableName))
                        {
                            snap.Tables[st.TableName] = st;
                            snap.TableList.Add(st);
                        }
                    }
                }
                else if (child.LocalName == "Relationships")
                {
                    foreach (XmlNode rn in child.ChildNodes)
                    {
                        if (rn.NodeType != XmlNodeType.Element) continue;
                        SnapshotRelationship sr = new SnapshotRelationship();
                        sr.ParentTable = Attr(rn, "pT");
                        sr.ParentColumn = Attr(rn, "pC");
                        sr.ChildTable = Attr(rn, "cT");
                        sr.ChildColumn = Attr(rn, "cC");
                        sr.ParentDistinctCount = AttrL(rn, "pD");
                        sr.ChildNonNullRows = AttrL(rn, "cN");
                        sr.ChildDistinctKeys = AttrL(rn, "cD");
                        sr.MatchedRows = AttrL(rn, "mR");
                        sr.OrphanRows = AttrL(rn, "oR");
                        sr.MatchRate = AttrD(rn, "rate");
                        sr.FanOutMin = AttrL(rn, "fMin");
                        sr.FanOutAvg = AttrD(rn, "fAvg");
                        sr.FanOutP50 = AttrD(rn, "fP50");
                        sr.FanOutP95 = AttrD(rn, "fP95");
                        sr.FanOutP99 = AttrD(rn, "fP99");
                        sr.FanOutMax = AttrL(rn, "fMax");
                        sr.MaxFanOutKey = Attr(rn, "maxK");
                        sr.RiskLevel = Attr(rn, "risk");
                        sr.Note = Attr(rn, "note");
                        snap.Relationships.Add(sr);
                    }
                }
                else if (child.LocalName == "Warnings")
                {
                    foreach (XmlNode wn in child.ChildNodes)
                    {
                        if (wn.NodeType != XmlNodeType.Element) continue;
                        SnapshotWarning sw = new SnapshotWarning();
                        sw.Severity = Attr(wn, "sev");
                        sw.Category = Attr(wn, "cat");
                        sw.TableName = Attr(wn, "tbl");
                        sw.ColumnName = Attr(wn, "col");
                        sw.MetricValue = Attr(wn, "met");
                        sw.PerfImpact = Attr(wn, "imp");
                        sw.Recommendation = Attr(wn, "rec");
                        snap.Warnings.Add(sw);
                    }
                }
                else if (child.LocalName == "JMeter")
                {
                    foreach (XmlNode jn in child.ChildNodes)
                    {
                        if (jn.NodeType != XmlNodeType.Element) continue;
                        SnapshotJMeterCandidate sj = new SnapshotJMeterCandidate();
                        sj.TableName = Attr(jn, "tbl");
                        sj.KeyColumn = Attr(jn, "col");
                        sj.UniqueCount = AttrL(jn, "uniq");
                        sj.TotalRows = AttrL(jn, "tot");
                        sj.DataType = Attr(jn, "type");
                        sj.Suitability = Attr(jn, "suit");
                        sj.RecommendedUsage = Attr(jn, "use");
                        snap.JMeterCandidates.Add(sj);
                    }
                }
            }
        }
    }

    public class ColumnDiffItem
    {
        public string TableName;
        public string ColumnName;
        public string Category;      // スキーマ追加 / スキーマ削除 / DDLスキーマ / インデックス / 外部キー / 物理DB指標 / テーブル指標 / 内容ハッシュ / カラム統計 / カテゴリ分布 / 数値統計 / 時系列・日時
        public string MetricName;
        public string BeforeValue;
        public string AfterValue;
        public string DeltaSummary;
        public string Severity;      // 通常 / 注意 / 警告
    }

    public class ColumnComparisonItem
    {
        public string TableName;
        public string ColumnName;
        public int ColumnIndex;
        public string TableStatus;
        public string ColumnStatus;  // 追加（新規テーブル） / 削除（テーブル消失） / 追加（新規カラム） / 削除（カラム消失） / 変更あり（スキーマ・データ） / 変更あり（スキーマ変更） / 変更あり（データ更新） / 変更あり（追加・統計変化） / 変更なし
        public bool IsChanged;
        public bool IsUnchanged;
        public string BeforeType;
        public string AfterType;
        public string TypeDisplay;
        public string DdlKeyDisplay;
        public long BeforeNonNull;
        public long AfterNonNull;
        public long NonNullDelta;
        public long BeforeNullCount;
        public long AfterNullCount;
        public long NullCountDelta;
        public string NullChangeDisplay;
        public long BeforeDistinct;
        public long AfterDistinct;
        public long DistinctDelta;
        public string DistinctChangeDisplay;
        public string ByteLenDisplay;
        public string MinValDisplay;
        public string MaxValDisplay;
        public string Top1Display;
        public string NumericSumDeltaDisplay;
        public string ContentHashStatus; // 一致 / 変化 / 追加 / 削除 / 未取得(旧形式)
        public int DiffCount;
        public string ChangeSummary;
        public string Severity;          // 警告 / 注意 / 通常 / 変化なし

        public ColumnComparisonItem()
        {
            TableName = "";
            ColumnName = "";
            TableStatus = "";
            ColumnStatus = "";
            BeforeType = "";
            AfterType = "";
            TypeDisplay = "";
            DdlKeyDisplay = "-";
            NullChangeDisplay = "-";
            DistinctChangeDisplay = "-";
            ByteLenDisplay = "-";
            MinValDisplay = "-";
            MaxValDisplay = "-";
            Top1Display = "-";
            NumericSumDeltaDisplay = "-";
            ContentHashStatus = "-";
            ChangeSummary = "";
            Severity = "変化なし";
        }
    }

    public class WarningComparisonItem
    {
        public string Status;       // 新規発生 / 悪化 / 継続 / 改善 / 解消
        public string Severity;
        public string Category;
        public string TableName;
        public string ColumnName;
        public string BeforeMetric;
        public string AfterMetric;
        public string PerfImpact;
        public string Recommendation;
        public bool IsAlert;

        public WarningComparisonItem()
        {
            Status = "";
            Severity = "";
            Category = "";
            TableName = "";
            ColumnName = "";
            BeforeMetric = "-";
            AfterMetric = "-";
            PerfImpact = "";
            Recommendation = "";
        }
    }

    public class RelationshipComparisonItem
    {
        public string Status;       // 新規検出 / 消失 / 悪化（要注意） / 改善 / 変化あり / 変化なし
        public string RelationKey;
        public string ParentTable;
        public string ParentColumn;
        public string ChildTable;
        public string ChildColumn;
        public double BeforeMatchRate;
        public double AfterMatchRate;
        public double MatchRateDelta;
        public long BeforeOrphanRows;
        public long AfterOrphanRows;
        public long OrphanDelta;
        public double BeforeFanOutAvg;
        public double AfterFanOutAvg;
        public double BeforeFanOutP95;
        public double AfterFanOutP95;
        public long BeforeFanOutMax;
        public long AfterFanOutMax;
        public long FanOutMaxDelta;
        public string BeforeRisk;
        public string AfterRisk;
        public string Note;
        public bool HasAlert;

        public RelationshipComparisonItem()
        {
            Status = "";
            RelationKey = "";
            ParentTable = "";
            ParentColumn = "";
            ChildTable = "";
            ChildColumn = "";
            BeforeRisk = "-";
            AfterRisk = "-";
            Note = "";
        }
    }

    public class JMeterComparisonItem
    {
        public string Status;       // 新規候補 / 消失 / 改善 / 悪化（要注意） / キー増加 / キー減少 / 変化なし
        public string TableName;
        public string KeyColumn;
        public string DataType;
        public long BeforeUnique;
        public long AfterUnique;
        public long UniqueDelta;
        public string BeforeSuitability;
        public string AfterSuitability;
        public string RecommendedUsage;
        public bool HasAlert;

        public JMeterComparisonItem()
        {
            Status = "";
            TableName = "";
            KeyColumn = "";
            DataType = "";
            BeforeSuitability = "-";
            AfterSuitability = "-";
            RecommendedUsage = "";
        }
    }

    public class TimeSeriesDeltaItem
    {
        public string Granularity;  // 月別 / 日別 / 時間帯別 / 曜日別
        public string TableColumn;
        public string PeriodLabel;
        public long BeforeCount;
        public long AfterCount;
        public long DeltaCount;
        public bool HasGrowthRate;
        public double GrowthRate;

        public TimeSeriesDeltaItem()
        {
            Granularity = "";
            TableColumn = "";
            PeriodLabel = "";
        }
    }

    public class ServerMetricComparisonItem
    {
        public string Category;
        public string MetricName;
        public string Unit;
        public long BeforeValue;
        public long AfterValue;
        public long DeltaValue;
        public string DisplayBefore;
        public string DisplayAfter;
        public string DisplayDelta;
        public string RatePerSecDisplay;
        public string Severity;
        public string Note;

        public ServerMetricComparisonItem()
        {
            Category = "";
            MetricName = "";
            Unit = "";
            DisplayBefore = "";
            DisplayAfter = "";
            DisplayDelta = "";
            RatePerSecDisplay = "-";
            Severity = "通常";
            Note = "";
        }
    }

    public class DbSchemaDiffItem
    {
        public string TableName;
        public string ObjectType;   // カラムDDL / インデックス / 外部キー制約 / テーブル設定
        public string ObjectName;
        public string ChangeType;   // 追加 / 削除 / 変更
        public string BeforeDef;
        public string AfterDef;
        public string Severity;
        public string Note;

        public DbSchemaDiffItem()
        {
            TableName = "";
            ObjectType = "";
            ObjectName = "";
            ChangeType = "";
            BeforeDef = "-";
            AfterDef = "-";
            Severity = "注意";
            Note = "";
        }
    }

    public class TableComparisonItem
    {
        public string TableName;
        public string PrimaryStatus; // 追加（新規テーブル） / 追加（レコード増加） / 追加・更新 / 更新（データ更新） / 更新（レコード減少） / 削除（テーブル消失） / 変更なし
        public string ChangeGroup;   // 追加 / 追加・更新 / 更新 / 削除 / 変更なし
        public string AddedType;     // 新規テーブル追加 / レコード追加 / レコード追加・更新 / ""
        public bool IsNewTable;
        public bool IsAdded;
        public bool IsUpdated;
        public bool IsDeleted;
        public bool IsUnchanged;
        public bool HasInPlaceMutation;

        public string InferredRole;
        public long BeforeRows;
        public long AfterRows;
        public long RowDelta;
        public bool HasRowGrowthRate;
        public double RowGrowthRate;
        public bool HasVelocity;
        public double RowsPerMinute;

        public int BeforeCols;
        public int AfterCols;
        public int ColDelta;

        public double BeforeSizeKB;
        public double AfterSizeKB;
        public double SizeDeltaKB;

        public double BeforeInnoDbKB;
        public double AfterInnoDbKB;
        public double InnoDbDeltaKB;

        public double BeforeAvgRowBytes;
        public double AfterAvgRowBytes;
        public double AvgRowBytesDelta;

        public long BeforeMaxRowBytes;
        public long AfterMaxRowBytes;
        public long MaxRowBytesDelta;

        public double BeforeAvgNullRate;
        public double AfterAvgNullRate;
        public double AvgNullRateDelta;

        public string CandidatePks;
        public string DateColumns;
        public string BeforeWarning;
        public string AfterWarning;
        public int BeforeWarningCount;
        public int AfterWarningCount;
        public int WarningCountDelta;
        public string BeforeJMeterReadiness;
        public string AfterJMeterReadiness;

        public string ContentHashStatus; // 一致 / 変化 / 追加 / 削除 / 未取得(旧形式)
        public bool ContentHashChanged;
        public bool HasRowChurn;
        public bool RowChurnExact;
        public long InsertedRows;
        public long DeletedRows;
        public long UpdatedRows;
        public long UnchangedRows;
        public string ChurnSummary;

        public bool HasDbPhysicalMeta;
        public double BeforeDataLengthKB;
        public double AfterDataLengthKB;
        public double DataLengthDeltaKB;
        public double BeforeIndexLengthKB;
        public double AfterIndexLengthKB;
        public double IndexLengthDeltaKB;
        public double BeforeDataFreeKB;
        public double AfterDataFreeKB;
        public double DataFreeDeltaKB;
        public bool HasAutoIncrement;
        public long BeforeAutoIncrement;
        public long AfterAutoIncrement;
        public long AutoIncrementDelta;
        public long AutoIncrementGap;
        public string BeforeUpdateTime;
        public string AfterUpdateTime;

        public int ChangedColumnCount;
        public int UnchangedColumnCount;
        public int AddedColumnCount;
        public int RemovedColumnCount;
        public int SchemaChangedColumnCount;
        public string ChangedColumnsSummary;
        public string UnchangedColumnsSummary;
        public string ChangeDetail;
        public string DiffSummary;
        public bool HasAlert;
        public string AlertNote;

        public List<ColumnDiffItem> ColumnDiffs;
        public List<ColumnComparisonItem> ColumnComparisons;

        public TableComparisonItem()
        {
            PrimaryStatus = "";
            ChangeGroup = "";
            AddedType = "";
            InferredRole = "";
            CandidatePks = "";
            DateColumns = "";
            BeforeWarning = "";
            AfterWarning = "";
            BeforeJMeterReadiness = "";
            AfterJMeterReadiness = "";
            ContentHashStatus = "-";
            ChurnSummary = "-";
            BeforeUpdateTime = "";
            AfterUpdateTime = "";
            ChangedColumnsSummary = "";
            UnchangedColumnsSummary = "-";
            ChangeDetail = "";
            DiffSummary = "";
            AlertNote = "";
            ColumnDiffs = new List<ColumnDiffItem>();
            ColumnComparisons = new List<ColumnComparisonItem>();
        }
    }

    public class AnalysisComparisonResult
    {
        public string BeforeFilePath;
        public string AfterFilePath;
        public string BeforeFileName;
        public string AfterFileName;
        public string BeforeSubtitle;
        public string AfterSubtitle;
        public string BeforeAnalyzedAt;
        public string AfterAnalyzedAt;
        public DateTime ComparedAt;

        public string BeforeSourceMode;
        public string AfterSourceMode;
        public bool BeforeIsDirectDb;
        public bool AfterIsDirectDb;
        public string BeforeDbLabel;
        public string AfterDbLabel;
        public bool BeforeHasEmbeddedSnapshot;
        public bool AfterHasEmbeddedSnapshot;
        public bool UsedFullSnapshot;
        public string PrecisionLabel;

        public bool HasElapsedTime;
        public double ElapsedSeconds;
        public string ElapsedDisplay;
        public bool ClockReversed;
        public double TotalRowsPerMinute;
        public double TotalInnoDbKBPerMinute;

        public bool HasAcuComparison;
        public double BeforeEstMinAcuCurrent;
        public double AfterEstMinAcuCurrent;
        public double BeforeEstMinAcu10x;
        public double AfterEstMinAcu10x;

        public bool HasDbPhysicalMeta;
        public double BeforeTotalDataLengthKB;
        public double AfterTotalDataLengthKB;
        public double TotalDataLengthDeltaKB;
        public double BeforeTotalIndexLengthKB;
        public double AfterTotalIndexLengthKB;
        public double TotalIndexLengthDeltaKB;
        public double BeforeTotalDataFreeKB;
        public double AfterTotalDataFreeKB;
        public double TotalDataFreeDeltaKB;

        public bool HasDbServerStatus;
        public bool DbRestartDetected;
        public long ServerUptimeDeltaSeconds;

        public bool HasAnyRowChurn;
        public int RowChurnCoveredTables;
        public bool RowChurnAllExact;
        public long TotalInsertedRows;
        public long TotalUpdatedRows;
        public long TotalDeletedRows;
        public long TotalUnchangedRows;

        public int BeforeTableCount;
        public int AfterTableCount;
        public int TotalUniqueTableCount;
        public int ChangedTableCount;

        public long BeforeTotalRows;
        public long AfterTotalRows;
        public long TotalRowDelta;
        public bool HasTotalRowGrowthRate;
        public double TotalRowGrowthRate;

        public double BeforeTotalSizeKB;
        public double AfterTotalSizeKB;
        public double TotalSizeDeltaKB;

        public double BeforeTotalInnoDbKB;
        public double AfterTotalInnoDbKB;
        public double TotalInnoDbDeltaKB;

        public int NewTableCount;
        public int RowAddedOnlyCount;
        public int AddedAndUpdatedCount;
        public int UpdatedOnlyCount;
        public int RowDecreasedCount;
        public int DeletedTableCount;
        public int UnchangedTableCount;

        public int AddedSheetCount;
        public int UpdatedSheetCount;
        public int AlertTableCount;

        public int TotalUniqueColumnCount;
        public int ChangedColumnTotal;
        public int UnchangedColumnTotal;
        public int AddedColumnTotal;
        public int RemovedColumnTotal;
        public int SchemaChangedColumnTotal;

        public int NewWarningCount;
        public int WorsenedWarningCount;
        public int ResolvedWarningCount;
        public int OngoingWarningCount;
        public long BeforeTotalOrphanRows;
        public long AfterTotalOrphanRows;
        public long OrphanRowDelta;

        public List<TableComparisonItem> AllTables;
        public List<TableComparisonItem> UpdatedTables;
        public List<TableComparisonItem> AddedTables;
        public List<TableComparisonItem> UnchangedTables;
        public List<TableComparisonItem> DeletedTables;
        public List<ColumnDiffItem> AllColumnDiffs;
        public List<ColumnComparisonItem> AllColumnComparisons;
        public List<WarningComparisonItem> WarningComparisons;
        public List<RelationshipComparisonItem> RelationshipComparisons;
        public List<JMeterComparisonItem> JMeterComparisons;
        public List<TimeSeriesDeltaItem> TimeSeriesDeltas;
        public List<ServerMetricComparisonItem> ServerMetricComparisons;
        public List<DbSchemaDiffItem> DbSchemaDiffs;
        public List<string> Insights;

        public AnalysisComparisonResult()
        {
            BeforeSubtitle = "";
            AfterSubtitle = "";
            BeforeAnalyzedAt = "";
            AfterAnalyzedAt = "";
            BeforeSourceMode = "";
            AfterSourceMode = "";
            BeforeDbLabel = "";
            AfterDbLabel = "";
            PrecisionLabel = "";
            ElapsedDisplay = "-";
            AllTables = new List<TableComparisonItem>();
            UpdatedTables = new List<TableComparisonItem>();
            AddedTables = new List<TableComparisonItem>();
            UnchangedTables = new List<TableComparisonItem>();
            DeletedTables = new List<TableComparisonItem>();
            AllColumnDiffs = new List<ColumnDiffItem>();
            AllColumnComparisons = new List<ColumnComparisonItem>();
            WarningComparisons = new List<WarningComparisonItem>();
            RelationshipComparisons = new List<RelationshipComparisonItem>();
            JMeterComparisons = new List<JMeterComparisonItem>();
            TimeSeriesDeltas = new List<TimeSeriesDeltaItem>();
            ServerMetricComparisons = new List<ServerMetricComparisonItem>();
            DbSchemaDiffs = new List<DbSchemaDiffItem>();
            Insights = new List<string>();
        }
    }

    public static class AnalysisComparisonEngine
    {
        private static readonly Regex AnomalyValueRegex = new Regex(
            @"(?i)(^|[\s_\-\./:])(err|error|fail|failed|failure|ng|reject|rejected|cancel|canceled|cancelled|timeout|abort|aborted|dead|deadlock|invalid|deny|denied|crash|exception|fault|異常|失敗|エラー|取消|却下|中断|タイムアウト)($|[\s_\-\./:])",
            RegexOptions.Compiled);

        public static AnalysisWorkbookSnapshot LoadAnalysisWorkbook(string xlsxPath)
        {
            if (string.IsNullOrWhiteSpace(xlsxPath) || !File.Exists(xlsxPath))
            {
                throw new FileNotFoundException("分析Excelファイルが見つかりません: " + xlsxPath);
            }
            string ext = Path.GetExtension(xlsxPath);
            if (!string.Equals(ext, ".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("分析ファイルは .xlsx 形式のExcelファイルを指定してください: " + xlsxPath);
            }

            AnalysisWorkbookSnapshot snap = new AnalysisWorkbookSnapshot();
            snap.FilePath = Path.GetFullPath(xlsxPath);
            snap.FileName = Path.GetFileName(xlsxPath);

            try
            {
                using (FileStream fs = new FileStream(xlsxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read))
                {
                    List<string> sharedStrings = ReadSharedStrings(zip);
                    Dictionary<string, string> sheetParts = ReadSheetPartMap(zip);

                    string tblPart;
                    if (!sheetParts.TryGetValue("02_テーブル一覧", out tblPart))
                    {
                        throw new InvalidDataException("指定されたExcelファイルに '02_テーブル一覧' シートが見つかりません。本ツールで出力した分析ファイル (.xlsx) を指定してください: " + xlsxPath);
                    }

                    // Extract metadata subtitle from 01_サマリー (keeps display metadata consistent)
                    string sumPart;
                    if (sheetParts.TryGetValue("01_サマリー", out sumPart))
                    {
                        List<Dictionary<int, string>> sumRows = ReadSheetRows(zip, sumPart, sharedStrings, 4);
                        if (sumRows.Count >= 2)
                        {
                            string sub = GetCell(sumRows[1], 1);
                            if (!string.IsNullOrWhiteSpace(sub)) ParseSubtitleMeta(snap, sub);
                        }
                    }

                    // Prefer lossless OOXML CustomXML snapshot if present (available in DirectDB & CSV, 3-sheet & 8-sheet)
                    if (AnalysisSnapshotCodec.TryLoadFromZip(zip, snap) && snap.TableList.Count > 0)
                    {
                        return snap;
                    }

                    // Fallback: parse visible worksheets (legacy .xlsx compatibility)
                    List<Dictionary<int, string>> tblRows = ReadSheetRows(zip, tblPart, sharedStrings, 0);
                    if (string.IsNullOrWhiteSpace(snap.Subtitle) && tblRows.Count >= 2)
                    {
                        string sub = GetCell(tblRows[1], 1);
                        if (!string.IsNullOrWhiteSpace(sub)) ParseSubtitleMeta(snap, sub);
                    }
                    ParseTablesSheet(snap, tblRows);

                    string colPart;
                    if (sheetParts.TryGetValue("03_カラム分析", out colPart))
                    {
                        List<Dictionary<int, string>> colRows = ReadSheetRows(zip, colPart, sharedStrings, 0);
                        ParseColumnsSheet(snap, colRows);
                    }

                    string distPart;
                    if (sheetParts.TryGetValue("04_分布・割合分析", out distPart))
                    {
                        List<Dictionary<int, string>> distRows = ReadSheetRows(zip, distPart, sharedStrings, 0);
                        ParseDistributionsSheet(snap, distRows);
                    }

                    string numPart;
                    if (sheetParts.TryGetValue("05_数値統計・ヒストグラム", out numPart))
                    {
                        List<Dictionary<int, string>> numRows = ReadSheetRows(zip, numPart, sharedStrings, 0);
                        ParseNumericSheet(snap, numRows);
                    }

                    string dtPart;
                    if (sheetParts.TryGetValue("06_時系列分析", out dtPart))
                    {
                        List<Dictionary<int, string>> dtRows = ReadSheetRows(zip, dtPart, sharedStrings, 0);
                        ParseDateTimeSheet(snap, dtRows);
                    }

                    string relPart;
                    if (sheetParts.TryGetValue("07_リレーション・ファンアウト", out relPart))
                    {
                        List<Dictionary<int, string>> relRows = ReadSheetRows(zip, relPart, sharedStrings, 0);
                        ParseRelationshipsSheet(snap, relRows);
                    }

                    string warnPart;
                    if (sheetParts.TryGetValue("08_性能試験ガイド・警告", out warnPart))
                    {
                        List<Dictionary<int, string>> warnRows = ReadSheetRows(zip, warnPart, sharedStrings, 0);
                        ParseWarningsSheet(snap, warnRows);
                    }
                }
            }
            catch (InvalidDataException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new InvalidDataException("分析Excelファイルの読込に失敗しました (" + Path.GetFileName(xlsxPath) + "): " + ex.Message, ex);
            }

            return snap;
        }

        public static AnalysisComparisonResult CompareAnalysisWorkbooks(string beforePath, string afterPath)
        {
            return CompareAnalysisWorkbooks(beforePath, afterPath, null);
        }

        public static AnalysisComparisonResult CompareAnalysisWorkbooks(string beforePath, string afterPath, OperationCancelState cancelState)
        {
            if (cancelState != null) cancelState.ThrowIfCancelled();
            AnalysisWorkbookSnapshot before = LoadAnalysisWorkbook(beforePath);
            if (cancelState != null) cancelState.ThrowIfCancelled();
            AnalysisWorkbookSnapshot after = LoadAnalysisWorkbook(afterPath);
            if (cancelState != null) cancelState.ThrowIfCancelled();
            return CompareSnapshots(before, after);
        }

        private static string ModeDisplay(AnalysisWorkbookSnapshot s)
        {
            if (s.IsDirectDb || string.Equals(s.SourceMode, "DirectDB", StringComparison.OrdinalIgnoreCase)) return "DB直接";
            return "CSV";
        }

        private static string FormatElapsed(double seconds)
        {
            if (seconds < 0) seconds = 0;
            TimeSpan ts = TimeSpan.FromSeconds(Math.Round(seconds));
            if (ts.TotalDays >= 1.0) return string.Format(CultureInfo.InvariantCulture, "{0}日{1}時間{2}分", (int)ts.TotalDays, ts.Hours, ts.Minutes);
            if (ts.TotalHours >= 1.0) return string.Format(CultureInfo.InvariantCulture, "{0}時間{1}分{2}秒", (int)ts.TotalHours, ts.Minutes, ts.Seconds);
            if (ts.TotalMinutes >= 1.0) return string.Format(CultureInfo.InvariantCulture, "{0}分{1}秒", (int)ts.TotalMinutes, ts.Seconds);
            return string.Format(CultureInfo.InvariantCulture, "{0:0.0}秒", seconds);
        }

        public static AnalysisComparisonResult CompareSnapshots(AnalysisWorkbookSnapshot before, AnalysisWorkbookSnapshot after)
        {
            AnalysisComparisonResult res = new AnalysisComparisonResult();
            res.BeforeFilePath = before.FilePath;
            res.AfterFilePath = after.FilePath;
            res.BeforeFileName = before.FileName;
            res.AfterFileName = after.FileName;
            res.BeforeSubtitle = before.Subtitle;
            res.AfterSubtitle = after.Subtitle;
            res.BeforeAnalyzedAt = before.AnalyzedAtText;
            res.AfterAnalyzedAt = after.AnalyzedAtText;
            res.BeforeSourceMode = ModeDisplay(before);
            res.AfterSourceMode = ModeDisplay(after);
            res.BeforeIsDirectDb = before.IsDirectDb;
            res.AfterIsDirectDb = after.IsDirectDb;
            res.BeforeDbLabel = before.DbLabel ?? "";
            res.AfterDbLabel = after.DbLabel ?? "";
            res.BeforeHasEmbeddedSnapshot = before.HasEmbeddedSnapshot;
            res.AfterHasEmbeddedSnapshot = after.HasEmbeddedSnapshot;
            res.UsedFullSnapshot = before.HasEmbeddedSnapshot && after.HasEmbeddedSnapshot;
            res.ComparedAt = DateTime.Now;
            res.BeforeTableCount = before.TableList.Count;
            res.AfterTableCount = after.TableList.Count;

            // Elapsed time calculation
            DateTime bDt, aDt;
            bool hasBdt = before.AnalyzedAt.HasValue;
            bDt = hasBdt ? before.AnalyzedAt.Value : DateTime.MinValue;
            if (!hasBdt && TryParseComparableDate(before.AnalyzedAtText, out bDt)) hasBdt = true;
            bool hasAdt = after.AnalyzedAt.HasValue;
            aDt = hasAdt ? after.AnalyzedAt.Value : DateTime.MinValue;
            if (!hasAdt && TryParseComparableDate(after.AnalyzedAtText, out aDt)) hasAdt = true;
            if (hasBdt && hasAdt)
            {
                double diffSec = (aDt - bDt).TotalSeconds;
                if (diffSec < -1.0)
                {
                    res.ClockReversed = true;
                    res.ElapsedDisplay = "日時逆転 (" + FormatElapsed(-diffSec) + ")";
                }
                else if (diffSec >= 1.0)
                {
                    res.HasElapsedTime = true;
                    res.ElapsedSeconds = diffSec;
                    res.ElapsedDisplay = FormatElapsed(diffSec);
                }
            }

            if (before.HasAcuEstimate && after.HasAcuEstimate)
            {
                res.HasAcuComparison = true;
                res.BeforeEstMinAcuCurrent = before.EstMinAcuCurrent;
                res.AfterEstMinAcuCurrent = after.EstMinAcuCurrent;
                res.BeforeEstMinAcu10x = before.EstMinAcu10x;
                res.AfterEstMinAcu10x = after.EstMinAcu10x;
            }

            foreach (SnapshotTable bt in before.TableList)
            {
                res.BeforeTotalRows += bt.RowCount;
                res.BeforeTotalSizeKB += bt.FileSizeKB;
                res.BeforeTotalInnoDbKB += bt.EstInnoDbKB;
                if (bt.DbMeta != null && bt.DbMeta.HasMetadata)
                {
                    res.BeforeTotalDataLengthKB += bt.DbMeta.DataLength / 1024.0;
                    res.BeforeTotalIndexLengthKB += bt.DbMeta.IndexLength / 1024.0;
                    res.BeforeTotalDataFreeKB += bt.DbMeta.DataFree / 1024.0;
                }
            }
            foreach (SnapshotTable at in after.TableList)
            {
                res.AfterTotalRows += at.RowCount;
                res.AfterTotalSizeKB += at.FileSizeKB;
                res.AfterTotalInnoDbKB += at.EstInnoDbKB;
                if (at.DbMeta != null && at.DbMeta.HasMetadata)
                {
                    res.AfterTotalDataLengthKB += at.DbMeta.DataLength / 1024.0;
                    res.AfterTotalIndexLengthKB += at.DbMeta.IndexLength / 1024.0;
                    res.AfterTotalDataFreeKB += at.DbMeta.DataFree / 1024.0;
                }
            }
            res.TotalRowDelta = res.AfterTotalRows - res.BeforeTotalRows;
            if (res.BeforeTotalRows > 0)
            {
                res.HasTotalRowGrowthRate = true;
                res.TotalRowGrowthRate = (double)res.TotalRowDelta / res.BeforeTotalRows;
            }
            res.TotalSizeDeltaKB = Math.Round(res.AfterTotalSizeKB - res.BeforeTotalSizeKB, 2);
            res.TotalInnoDbDeltaKB = Math.Round(res.AfterTotalInnoDbKB - res.BeforeTotalInnoDbKB, 1);

            // Compare MySQL Server status counters first so Uptime can supplement elapsed time if needed
            CompareServerMetrics(before, after, res);
            if (!res.HasElapsedTime && !res.ClockReversed && res.ServerUptimeDeltaSeconds >= 1)
            {
                res.HasElapsedTime = true;
                res.ElapsedSeconds = res.ServerUptimeDeltaSeconds;
                res.ElapsedDisplay = FormatElapsed(res.ElapsedSeconds);
            }
            if (res.HasElapsedTime && res.ElapsedSeconds > 0.0)
            {
                double mins = res.ElapsedSeconds / 60.0;
                res.TotalRowsPerMinute = res.TotalRowDelta / mins;
                res.TotalInnoDbKBPerMinute = res.TotalInnoDbDeltaKB / mins;
            }

            // Collect union of all table names (preserving After's casing when present)
            List<string> allNames = new List<string>();
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (SnapshotTable at in after.TableList)
            {
                if (seen.Add(at.TableName)) allNames.Add(at.TableName);
            }
            foreach (SnapshotTable bt in before.TableList)
            {
                if (seen.Add(bt.TableName)) allNames.Add(bt.TableName);
            }
            res.TotalUniqueTableCount = allNames.Count;

            bool allChurnExact = true;
            foreach (string tblName in allNames)
            {
                SnapshotTable bt = null;
                SnapshotTable at = null;
                before.Tables.TryGetValue(tblName, out bt);
                after.Tables.TryGetValue(tblName, out at);

                TableComparisonItem item = CompareSingleTable(bt, at, res);
                if (res.HasElapsedTime && res.ElapsedSeconds > 0.0)
                {
                    item.HasVelocity = true;
                    item.RowsPerMinute = item.RowDelta / (res.ElapsedSeconds / 60.0);
                }

                res.AllTables.Add(item);
                foreach (ColumnDiffItem cd in item.ColumnDiffs)
                {
                    res.AllColumnDiffs.Add(cd);
                }
                foreach (ColumnComparisonItem cc in item.ColumnComparisons)
                {
                    res.AllColumnComparisons.Add(cc);
                    res.TotalUniqueColumnCount++;
                    if (cc.IsChanged) res.ChangedColumnTotal++;
                    else res.UnchangedColumnTotal++;
                    if (cc.ColumnStatus == "追加（新規カラム）" || cc.ColumnStatus == "追加（新規テーブル）") res.AddedColumnTotal++;
                    else if (cc.ColumnStatus == "削除（カラム消失）" || cc.ColumnStatus == "削除（テーブル消失）") res.RemovedColumnTotal++;
                    if (cc.ColumnStatus.IndexOf("スキーマ", StringComparison.Ordinal) >= 0) res.SchemaChangedColumnTotal++;
                }

                if (item.HasDbPhysicalMeta) res.HasDbPhysicalMeta = true;
                if (item.HasRowChurn)
                {
                    res.HasAnyRowChurn = true;
                    res.RowChurnCoveredTables++;
                    if (!item.RowChurnExact) allChurnExact = false;
                    res.TotalInsertedRows += item.InsertedRows;
                    res.TotalUpdatedRows += item.UpdatedRows;
                    res.TotalDeletedRows += item.DeletedRows;
                    res.TotalUnchangedRows += item.UnchangedRows;
                }

                if (item.IsNewTable) res.NewTableCount++;
                else if (item.IsDeleted) res.DeletedTableCount++;
                else if (item.IsUnchanged) res.UnchangedTableCount++;
                else if (item.PrimaryStatus == "追加（レコード増加）") res.RowAddedOnlyCount++;
                else if (item.PrimaryStatus == "追加・更新") res.AddedAndUpdatedCount++;
                else if (item.PrimaryStatus == "更新（データ更新）") res.UpdatedOnlyCount++;
                else if (item.PrimaryStatus == "更新（レコード減少）") res.RowDecreasedCount++;

                if (item.IsAdded) res.AddedTables.Add(item);
                if (item.IsUpdated) res.UpdatedTables.Add(item);
                if (item.IsUnchanged) res.UnchangedTables.Add(item);
                if (item.IsDeleted) res.DeletedTables.Add(item);
                if (item.HasAlert) res.AlertTableCount++;
            }

            res.RowChurnAllExact = res.HasAnyRowChurn && allChurnExact;
            res.ChangedTableCount = res.TotalUniqueTableCount - res.UnchangedTableCount;
            res.AddedSheetCount = res.AddedTables.Count;
            res.UpdatedSheetCount = res.UpdatedTables.Count;

            if (res.HasDbPhysicalMeta)
            {
                res.BeforeTotalDataLengthKB = Math.Round(res.BeforeTotalDataLengthKB, 1);
                res.AfterTotalDataLengthKB = Math.Round(res.AfterTotalDataLengthKB, 1);
                res.TotalDataLengthDeltaKB = Math.Round(res.AfterTotalDataLengthKB - res.BeforeTotalDataLengthKB, 1);
                res.BeforeTotalIndexLengthKB = Math.Round(res.BeforeTotalIndexLengthKB, 1);
                res.AfterTotalIndexLengthKB = Math.Round(res.AfterTotalIndexLengthKB, 1);
                res.TotalIndexLengthDeltaKB = Math.Round(res.AfterTotalIndexLengthKB - res.BeforeTotalIndexLengthKB, 1);
                res.BeforeTotalDataFreeKB = Math.Round(res.BeforeTotalDataFreeKB, 1);
                res.AfterTotalDataFreeKB = Math.Round(res.AfterTotalDataFreeKB, 1);
                res.TotalDataFreeDeltaKB = Math.Round(res.AfterTotalDataFreeKB - res.BeforeTotalDataFreeKB, 1);
            }

            // Precision provenance label
            string modePair = string.Format("{0}×{1}", res.BeforeSourceMode, res.AfterSourceMode);
            if (res.UsedFullSnapshot)
            {
                if (res.HasDbPhysicalMeta || res.HasDbServerStatus)
                {
                    res.PrecisionLabel = string.Format("完全照合・ハッシュ＋DBメタ ({0})", modePair);
                }
                else
                {
                    res.PrecisionLabel = string.Format("完全照合・ハッシュ＋全統計 ({0})", modePair);
                }
            }
            else if (res.BeforeHasEmbeddedSnapshot || res.AfterHasEmbeddedSnapshot)
            {
                res.PrecisionLabel = string.Format("混在互換照合・一部旧形式 ({0})", modePair);
            }
            else
            {
                res.PrecisionLabel = string.Format("表示シート互換解析 ({0})", modePair);
            }

            // Cross-table quality, integrity, JMeter & time-series comparisons
            CompareWarnings(before, after, res);
            CompareRelationships(before, after, res);
            CompareJMeterCandidates(before, after, res);
            CompareTimeSeries(before, after, res);

            // Sort lists deterministically for clear presentation
            res.AllTables = res.AllTables
                .OrderBy(t => GetStatusSortOrder(t.PrimaryStatus))
                .ThenByDescending(t => Math.Abs(t.RowDelta))
                .ThenByDescending(t => t.ChangedColumnCount)
                .ThenBy(t => t.TableName, StringComparer.OrdinalIgnoreCase)
                .ToList();

            res.AddedTables = res.AddedTables
                .OrderBy(t => t.IsNewTable ? 0 : (t.HasInPlaceMutation ? 1 : 2))
                .ThenByDescending(t => t.RowDelta)
                .ThenBy(t => t.TableName, StringComparer.OrdinalIgnoreCase)
                .ToList();

            res.UpdatedTables = res.UpdatedTables
                .OrderBy(t => t.PrimaryStatus == "更新（レコード減少）" ? 0 : (t.PrimaryStatus == "追加・更新" ? 1 : 2))
                .ThenByDescending(t => t.ChangedColumnCount)
                .ThenByDescending(t => Math.Abs(t.RowDelta))
                .ThenBy(t => t.TableName, StringComparer.OrdinalIgnoreCase)
                .ToList();

            res.UnchangedTables = res.UnchangedTables
                .OrderByDescending(t => t.AfterRows)
                .ThenBy(t => t.TableName, StringComparer.OrdinalIgnoreCase)
                .ToList();

            res.DeletedTables = res.DeletedTables
                .OrderByDescending(t => t.BeforeRows)
                .ThenBy(t => t.TableName, StringComparer.OrdinalIgnoreCase)
                .ToList();

            res.AllColumnDiffs = res.AllColumnDiffs
                .OrderBy(d => d.Severity == "警告" ? 0 : (d.Severity == "注意" ? 1 : 2))
                .ThenBy(d => d.TableName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(d => d.ColumnName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(d => d.Category, StringComparer.Ordinal)
                .ToList();

            res.AllColumnComparisons = res.AllColumnComparisons
                .OrderBy(c => GetStatusSortOrder(c.TableStatus))
                .ThenBy(c => c.TableName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(c => c.IsChanged ? 0 : 1)
                .ThenBy(c => c.ColumnIndex)
                .ThenBy(c => c.ColumnName, StringComparer.OrdinalIgnoreCase)
                .ToList();

            BuildComparisonInsights(res);
            return res;
        }

        private static int GetStatusSortOrder(string status)
        {
            if (status == "削除（テーブル消失）") return 0;
            if (status == "追加（新規テーブル）") return 1;
            if (status == "追加・更新") return 2;
            if (status == "追加（レコード増加）") return 3;
            if (status == "更新（レコード減少）") return 4;
            if (status == "更新（データ更新）") return 5;
            return 6; // 変更なし
        }

        private static string FormatDdlKeyDisplay(SnapshotColumn sc)
        {
            if (sc == null) return "-";
            if (sc.DbMeta != null && !string.IsNullOrEmpty(sc.DbMeta.ColumnType))
            {
                List<string> tags = new List<string>();
                if (!string.IsNullOrEmpty(sc.DbMeta.ColumnKey)) tags.Add(sc.DbMeta.ColumnKey);
                if (string.Equals(sc.DbMeta.IsNullable, "NO", StringComparison.OrdinalIgnoreCase)) tags.Add("NOT NULL");
                if (!string.IsNullOrEmpty(sc.DbMeta.Extra)) tags.Add(sc.DbMeta.Extra);
                if (tags.Count > 0) return sc.DbMeta.ColumnType + " [" + string.Join(", ", tags.ToArray()) + "]";
                return sc.DbMeta.ColumnType;
            }
            return string.IsNullOrEmpty(sc.KeyRole) ? (string.IsNullOrEmpty(sc.SelectivityRating) ? "-" : sc.SelectivityRating) : sc.KeyRole;
        }

        private static string FormatBeforeAfterText(string b, string a)
        {
            string sb = string.IsNullOrEmpty(b) ? "-" : b;
            string sa = string.IsNullOrEmpty(a) ? "-" : a;
            if (string.Equals(sb, sa, StringComparison.Ordinal)) return sa;
            return sb + " → " + sa;
        }

        private static void ComputeRowChurnFromSketches(SnapshotTable bt, SnapshotTable at, TableComparisonItem item)
        {
            if (bt == null || at == null) return;
            if (bt.RowCount == 0 && at.RowCount == 0)
            {
                item.HasRowChurn = true;
                item.RowChurnExact = true;
                item.ChurnSummary = "0件 (変動なし)";
                return;
            }
            if (bt.RowCount == 0 && at.RowCount > 0)
            {
                item.HasRowChurn = true;
                item.RowChurnExact = true;
                item.InsertedRows = at.RowCount;
                item.ChurnSummary = string.Format("挿入 +{0:#,##0} / 更新 0 / 削除 0", at.RowCount);
                return;
            }
            if (bt.RowCount > 0 && at.RowCount == 0)
            {
                item.HasRowChurn = true;
                item.RowChurnExact = true;
                item.DeletedRows = bt.RowCount;
                item.ChurnSummary = string.Format("挿入 0 / 更新 0 / 削除 -{0:#,##0}", bt.RowCount);
                return;
            }

            RowChurnSketch bs = bt.RowSketch;
            RowChurnSketch as_ = at.RowSketch;
            if (bs == null || !bs.IsValid || as_ == null || !as_.IsValid) return;
            if (!string.Equals(bs.KeyColumns, as_.KeyColumns, StringComparison.OrdinalIgnoreCase)) return;

            ulong threshold = Math.Min(bs.IsExact ? ulong.MaxValue : bs.MaxKeptPkHash, as_.IsExact ? ulong.MaxValue : as_.MaxKeptPkHash);
            int bi = 0, ai = 0;
            int bLim = bs.Entries.Count;
            int aLim = as_.Entries.Count;
            long wIns = 0, wDel = 0, wUpd = 0, wUnc = 0;

            while (bi < bLim && bs.Entries[bi].PkHash <= threshold && ai < aLim && as_.Entries[ai].PkHash <= threshold)
            {
                ulong bp = bs.Entries[bi].PkHash;
                ulong ap = as_.Entries[ai].PkHash;
                if (bp == ap)
                {
                    if (bs.Entries[bi].RowHash == as_.Entries[ai].RowHash) wUnc++;
                    else wUpd++;
                    bi++; ai++;
                }
                else if (bp < ap)
                {
                    wDel++;
                    bi++;
                }
                else
                {
                    wIns++;
                    ai++;
                }
            }
            while (bi < bLim && bs.Entries[bi].PkHash <= threshold) { wDel++; bi++; }
            while (ai < aLim && as_.Entries[ai].PkHash <= threshold) { wIns++; ai++; }

            if (bs.IsExact && as_.IsExact)
            {
                item.HasRowChurn = true;
                item.RowChurnExact = true;
                item.InsertedRows = wIns;
                item.DeletedRows = wDel;
                item.UpdatedRows = wUpd;
                item.UnchangedRows = wUnc;
                item.ChurnSummary = string.Format("挿入 +{0:#,##0} / 更新 {1:#,##0} / 削除 -{2:#,##0} / 無変更 {3:#,##0}", wIns, wUpd, wDel, wUnc);
            }
            else if ((wIns + wDel + wUpd + wUnc) >= 20)
            {
                long wBefore = wUnc + wUpd + wDel;
                long wAfter = wUnc + wUpd + wIns;
                double scale = (wBefore + wAfter) > 0 ? ((double)(bt.RowCount + at.RowCount) / (double)(wBefore + wAfter)) : 1.0;
                long estUpd = (long)Math.Round(wUpd * scale);
                long estDel = (long)Math.Round(wDel * scale);
                long estIns = estDel + item.RowDelta;
                if (estIns < 0)
                {
                    estDel = -item.RowDelta;
                    estIns = 0;
                }
                if (estUpd > bt.RowCount - estDel) estUpd = Math.Max(0L, bt.RowCount - estDel);
                long estUnc = Math.Max(0L, bt.RowCount - estDel - estUpd);

                item.HasRowChurn = true;
                item.RowChurnExact = false;
                item.InsertedRows = estIns;
                item.DeletedRows = estDel;
                item.UpdatedRows = estUpd;
                item.UnchangedRows = estUnc;
                item.ChurnSummary = string.Format("推定 挿入 +{0:#,##0} / 更新 ~{1:#,##0} / 削除 -{2:#,##0}", estIns, estUpd, estDel);
            }
        }

        private static TableComparisonItem CompareSingleTable(SnapshotTable bt, SnapshotTable at, AnalysisComparisonResult parentRes)
        {
            TableComparisonItem item = new TableComparisonItem();

            // Case 1: Brand-new table added in After
            if (bt == null && at != null)
            {
                item.TableName = at.TableName;
                item.PrimaryStatus = "追加（新規テーブル）";
                item.ChangeGroup = "追加";
                item.AddedType = "新規テーブル追加";
                item.IsNewTable = true;
                item.IsAdded = true;
                item.InferredRole = at.InferredRole;
                item.BeforeRows = 0;
                item.AfterRows = at.RowCount;
                item.RowDelta = at.RowCount;
                item.BeforeCols = 0;
                item.AfterCols = at.ColumnCount;
                item.ColDelta = at.ColumnCount;
                item.BeforeSizeKB = 0.0;
                item.AfterSizeKB = at.FileSizeKB;
                item.SizeDeltaKB = Math.Round(at.FileSizeKB, 2);
                item.BeforeInnoDbKB = 0.0;
                item.AfterInnoDbKB = at.EstInnoDbKB;
                item.InnoDbDeltaKB = Math.Round(at.EstInnoDbKB, 1);
                item.AfterAvgRowBytes = at.AvgRowBytes;
                item.AvgRowBytesDelta = Math.Round(at.AvgRowBytes, 1);
                item.AfterMaxRowBytes = at.MaxRowBytes;
                item.MaxRowBytesDelta = at.MaxRowBytes;
                item.AfterAvgNullRate = at.AvgNullRate;
                item.AvgNullRateDelta = at.AvgNullRate;
                item.CandidatePks = at.CandidatePks;
                item.DateColumns = at.DateColumns;
                item.BeforeWarning = "-";
                item.AfterWarning = string.IsNullOrEmpty(at.WarningSummary) ? "正常" : at.WarningSummary;
                item.AfterWarningCount = at.WarningCount;
                item.WarningCountDelta = at.WarningCount;
                item.AfterJMeterReadiness = at.JMeterReadiness;
                item.ContentHashStatus = !string.IsNullOrEmpty(at.RowMultisetHash) ? "追加" : "未取得(旧形式)";
                item.ContentHashChanged = true;
                item.HasRowChurn = true;
                item.RowChurnExact = true;
                item.InsertedRows = at.RowCount;
                item.ChurnSummary = string.Format("新規追加 +{0:#,##0}件", at.RowCount);

                if (at.DbMeta != null && at.DbMeta.HasMetadata)
                {
                    item.HasDbPhysicalMeta = true;
                    item.AfterDataLengthKB = Math.Round(at.DbMeta.DataLength / 1024.0, 1);
                    item.DataLengthDeltaKB = item.AfterDataLengthKB;
                    item.AfterIndexLengthKB = Math.Round(at.DbMeta.IndexLength / 1024.0, 1);
                    item.IndexLengthDeltaKB = item.AfterIndexLengthKB;
                    item.AfterDataFreeKB = Math.Round(at.DbMeta.DataFree / 1024.0, 1);
                    item.DataFreeDeltaKB = item.AfterDataFreeKB;
                    if (at.DbMeta.HasAutoIncrement)
                    {
                        item.HasAutoIncrement = true;
                        item.AfterAutoIncrement = at.DbMeta.AutoIncrement;
                        item.AutoIncrementDelta = at.DbMeta.AutoIncrement;
                    }
                    item.AfterUpdateTime = at.DbMeta.UpdateTime;
                }

                item.ChangedColumnCount = at.ColumnCount;
                item.AddedColumnCount = at.ColumnCount;
                item.UnchangedColumnCount = 0;
                item.ChangedColumnsSummary = string.Format("新規テーブル ({0}列)", at.ColumnCount);
                item.UnchangedColumnsSummary = "-";
                item.ChangeDetail = string.Format("新規テーブル追加 ({0:#,##0}件 / {1}列)", at.RowCount, at.ColumnCount);
                item.DiffSummary = string.Format("テスト開始後に新規追加: {0:#,##0}件, {1}列, CSV {2:#,##0.00}KB", at.RowCount, at.ColumnCount, at.FileSizeKB);

                if (at.RowCount == 0 || (!string.IsNullOrEmpty(at.WarningSummary) && at.WarningSummary != "正常"))
                {
                    item.HasAlert = true;
                    item.AlertNote = at.RowCount == 0 ? "新規テーブルだが0件(空テーブル)" : ("新規テーブルに警告あり: " + at.WarningSummary);
                }
                else
                {
                    item.AlertNote = "正常";
                }

                AddDiff(item, "-", "テーブル指標", "テーブル存在", "なし", string.Format("新規追加 ({0:#,##0}件 / {1}列)", at.RowCount, at.ColumnCount), "新規テーブル追加", "注意");

                foreach (SnapshotColumn ac in GetOrderedColumns(at))
                {
                    ColumnComparisonItem cc = new ColumnComparisonItem();
                    cc.TableName = at.TableName;
                    cc.ColumnName = ac.ColumnName;
                    cc.ColumnIndex = ac.ColumnIndex;
                    cc.TableStatus = item.PrimaryStatus;
                    cc.ColumnStatus = "追加（新規テーブル）";
                    cc.IsChanged = true;
                    cc.AfterType = ac.InferredType;
                    cc.TypeDisplay = ac.InferredType;
                    cc.DdlKeyDisplay = FormatDdlKeyDisplay(ac);
                    cc.AfterNonNull = ac.NonNullCount > 0 ? ac.NonNullCount : Math.Max(0L, ac.TotalRows - ac.NullCount);
                    cc.NonNullDelta = cc.AfterNonNull;
                    cc.AfterNullCount = ac.NullCount;
                    cc.NullCountDelta = ac.NullCount;
                    cc.NullChangeDisplay = string.Format("0 → {0:#,##0}", ac.NullCount);
                    cc.AfterDistinct = ac.DistinctCount;
                    cc.DistinctDelta = ac.DistinctCount;
                    cc.DistinctChangeDisplay = string.Format("0 → {0:#,##0}", ac.DistinctCount);
                    cc.ByteLenDisplay = string.Format("平均 {0:0.0} / 最大 {1:#,##0} B", ac.AvgByteLen, ac.MaxByteLen);
                    cc.MinValDisplay = string.IsNullOrEmpty(ac.MinValueText) ? "-" : ac.MinValueText;
                    cc.MaxValDisplay = string.IsNullOrEmpty(ac.MaxValueText) ? "-" : ac.MaxValueText;
                    cc.Top1Display = string.IsNullOrEmpty(ac.Top1Value) ? "-" : ac.Top1Value;
                    cc.ContentHashStatus = !string.IsNullOrEmpty(ac.ContentHash) ? "追加" : "未取得(旧形式)";
                    cc.DiffCount = 1;
                    cc.ChangeSummary = string.Format("新規テーブルのカラム ({0}, ユニーク {1:#,##0})", ac.InferredType, ac.DistinctCount);
                    cc.Severity = "注意";
                    item.ColumnComparisons.Add(cc);
                }
                return item;
            }

            // Case 2: Table disappeared / deleted in After
            if (bt != null && at == null)
            {
                item.TableName = bt.TableName;
                item.PrimaryStatus = "削除（テーブル消失）";
                item.ChangeGroup = "削除";
                item.IsDeleted = true;
                item.InferredRole = bt.InferredRole;
                item.BeforeRows = bt.RowCount;
                item.AfterRows = 0;
                item.RowDelta = -bt.RowCount;
                if (bt.RowCount > 0)
                {
                    item.HasRowGrowthRate = true;
                    item.RowGrowthRate = -1.0;
                }
                item.BeforeCols = bt.ColumnCount;
                item.AfterCols = 0;
                item.ColDelta = -bt.ColumnCount;
                item.BeforeSizeKB = bt.FileSizeKB;
                item.AfterSizeKB = 0.0;
                item.SizeDeltaKB = Math.Round(-bt.FileSizeKB, 2);
                item.BeforeInnoDbKB = bt.EstInnoDbKB;
                item.AfterInnoDbKB = 0.0;
                item.InnoDbDeltaKB = Math.Round(-bt.EstInnoDbKB, 1);
                item.BeforeAvgRowBytes = bt.AvgRowBytes;
                item.AvgRowBytesDelta = Math.Round(-bt.AvgRowBytes, 1);
                item.BeforeMaxRowBytes = bt.MaxRowBytes;
                item.MaxRowBytesDelta = -bt.MaxRowBytes;
                item.BeforeAvgNullRate = bt.AvgNullRate;
                item.AvgNullRateDelta = -bt.AvgNullRate;
                item.CandidatePks = bt.CandidatePks;
                item.DateColumns = bt.DateColumns;
                item.BeforeWarning = string.IsNullOrEmpty(bt.WarningSummary) ? "正常" : bt.WarningSummary;
                item.AfterWarning = "テーブル消失";
                item.BeforeWarningCount = bt.WarningCount;
                item.WarningCountDelta = -bt.WarningCount;
                item.BeforeJMeterReadiness = bt.JMeterReadiness;
                item.ContentHashStatus = !string.IsNullOrEmpty(bt.RowMultisetHash) ? "削除" : "未取得(旧形式)";
                item.ContentHashChanged = true;
                item.HasRowChurn = true;
                item.RowChurnExact = true;
                item.DeletedRows = bt.RowCount;
                item.ChurnSummary = string.Format("テーブル消失 -{0:#,##0}件", bt.RowCount);

                if (bt.DbMeta != null && bt.DbMeta.HasMetadata)
                {
                    item.HasDbPhysicalMeta = true;
                    item.BeforeDataLengthKB = Math.Round(bt.DbMeta.DataLength / 1024.0, 1);
                    item.DataLengthDeltaKB = -item.BeforeDataLengthKB;
                    item.BeforeIndexLengthKB = Math.Round(bt.DbMeta.IndexLength / 1024.0, 1);
                    item.IndexLengthDeltaKB = -item.BeforeIndexLengthKB;
                    item.BeforeDataFreeKB = Math.Round(bt.DbMeta.DataFree / 1024.0, 1);
                    item.DataFreeDeltaKB = -item.BeforeDataFreeKB;
                    if (bt.DbMeta.HasAutoIncrement)
                    {
                        item.HasAutoIncrement = true;
                        item.BeforeAutoIncrement = bt.DbMeta.AutoIncrement;
                        item.AutoIncrementDelta = -bt.DbMeta.AutoIncrement;
                    }
                    item.BeforeUpdateTime = bt.DbMeta.UpdateTime;
                }

                item.ChangedColumnCount = bt.ColumnCount;
                item.RemovedColumnCount = bt.ColumnCount;
                item.UnchangedColumnCount = 0;
                item.ChangedColumnsSummary = string.Format("テーブル消失 (開始前{0}列)", bt.ColumnCount);
                item.UnchangedColumnsSummary = "-";
                item.ChangeDetail = string.Format("テーブル消失 (開始前 {0:#,##0}件)", bt.RowCount);
                item.DiffSummary = string.Format("テスト開始後の分析ファイルに存在せず (開始前: {0:#,##0}件, {1}列)", bt.RowCount, bt.ColumnCount);
                item.HasAlert = true;
                item.AlertNote = "要注意: テスト開始後にテーブルが消失/未出力";

                AddDiff(item, "-", "テーブル指標", "テーブル存在", string.Format("存在 ({0:#,##0}件 / {1}列)", bt.RowCount, bt.ColumnCount), "消失", "テーブル消失", "警告");

                foreach (SnapshotColumn bc in GetOrderedColumns(bt))
                {
                    ColumnComparisonItem cc = new ColumnComparisonItem();
                    cc.TableName = bt.TableName;
                    cc.ColumnName = bc.ColumnName;
                    cc.ColumnIndex = bc.ColumnIndex;
                    cc.TableStatus = item.PrimaryStatus;
                    cc.ColumnStatus = "削除（テーブル消失）";
                    cc.IsChanged = true;
                    cc.BeforeType = bc.InferredType;
                    cc.TypeDisplay = bc.InferredType + " → 消失";
                    cc.DdlKeyDisplay = FormatDdlKeyDisplay(bc);
                    cc.BeforeNonNull = bc.NonNullCount > 0 ? bc.NonNullCount : Math.Max(0L, bc.TotalRows - bc.NullCount);
                    cc.NonNullDelta = -cc.BeforeNonNull;
                    cc.BeforeNullCount = bc.NullCount;
                    cc.NullCountDelta = -bc.NullCount;
                    cc.NullChangeDisplay = string.Format("{0:#,##0} → 0", bc.NullCount);
                    cc.BeforeDistinct = bc.DistinctCount;
                    cc.DistinctDelta = -bc.DistinctCount;
                    cc.DistinctChangeDisplay = string.Format("{0:#,##0} → 0", bc.DistinctCount);
                    cc.ByteLenDisplay = string.Format("平均 {0:0.0} / 最大 {1:#,##0} B", bc.AvgByteLen, bc.MaxByteLen);
                    cc.MinValDisplay = string.IsNullOrEmpty(bc.MinValueText) ? "-" : bc.MinValueText;
                    cc.MaxValDisplay = string.IsNullOrEmpty(bc.MaxValueText) ? "-" : bc.MaxValueText;
                    cc.Top1Display = string.IsNullOrEmpty(bc.Top1Value) ? "-" : bc.Top1Value;
                    cc.ContentHashStatus = !string.IsNullOrEmpty(bc.ContentHash) ? "削除" : "未取得(旧形式)";
                    cc.DiffCount = 1;
                    cc.ChangeSummary = "親テーブル消失により未検出";
                    cc.Severity = "警告";
                    item.ColumnComparisons.Add(cc);
                }
                return item;
            }

            // Case 3: Table exists in BOTH Before and After
            item.TableName = at.TableName;
            item.InferredRole = !string.IsNullOrEmpty(at.InferredRole) ? at.InferredRole : bt.InferredRole;
            item.BeforeRows = bt.RowCount;
            item.AfterRows = at.RowCount;
            item.RowDelta = at.RowCount - bt.RowCount;
            if (bt.RowCount > 0)
            {
                item.HasRowGrowthRate = true;
                item.RowGrowthRate = (double)item.RowDelta / bt.RowCount;
            }
            item.BeforeCols = bt.ColumnCount;
            item.AfterCols = at.ColumnCount;
            item.ColDelta = at.ColumnCount - bt.ColumnCount;
            item.BeforeSizeKB = bt.FileSizeKB;
            item.AfterSizeKB = at.FileSizeKB;
            item.SizeDeltaKB = Math.Round(at.FileSizeKB - bt.FileSizeKB, 2);
            item.BeforeInnoDbKB = bt.EstInnoDbKB;
            item.AfterInnoDbKB = at.EstInnoDbKB;
            item.InnoDbDeltaKB = Math.Round(at.EstInnoDbKB - bt.EstInnoDbKB, 1);
            item.BeforeAvgRowBytes = bt.AvgRowBytes;
            item.AfterAvgRowBytes = at.AvgRowBytes;
            item.AvgRowBytesDelta = Math.Round(at.AvgRowBytes - bt.AvgRowBytes, 1);
            item.BeforeMaxRowBytes = bt.MaxRowBytes;
            item.AfterMaxRowBytes = at.MaxRowBytes;
            item.MaxRowBytesDelta = at.MaxRowBytes - bt.MaxRowBytes;
            item.BeforeAvgNullRate = bt.AvgNullRate;
            item.AfterAvgNullRate = at.AvgNullRate;
            item.AvgNullRateDelta = at.AvgNullRate - bt.AvgNullRate;
            item.CandidatePks = !string.IsNullOrEmpty(at.CandidatePks) ? at.CandidatePks : bt.CandidatePks;
            item.DateColumns = !string.IsNullOrEmpty(at.DateColumns) ? at.DateColumns : bt.DateColumns;
            item.BeforeWarning = string.IsNullOrEmpty(bt.WarningSummary) ? "正常" : bt.WarningSummary;
            item.AfterWarning = string.IsNullOrEmpty(at.WarningSummary) ? "正常" : at.WarningSummary;
            item.BeforeWarningCount = bt.WarningCount;
            item.AfterWarningCount = at.WarningCount;
            item.WarningCountDelta = at.WarningCount - bt.WarningCount;
            item.BeforeJMeterReadiness = bt.JMeterReadiness;
            item.AfterJMeterReadiness = at.JMeterReadiness;

            bool hasIntrinsicTableChange = false;
            bool hasInPlaceMutation = false;
            bool hasSchemaChange = false;
            HashSet<string> changedColSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            HashSet<string> schemaChangedColSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            HashSet<string> inPlaceColSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            List<string> summaryParts = new List<string>();
            List<string> alertParts = new List<string>();

            // 3.0 Exact Row Multiset Hash & PK Row Churn Sketch comparison
            bool bothHaveRowHash = !string.IsNullOrEmpty(bt.RowMultisetHash) && !string.IsNullOrEmpty(at.RowMultisetHash);
            if (bothHaveRowHash)
            {
                if (!string.Equals(bt.RowMultisetHash, at.RowMultisetHash, StringComparison.Ordinal))
                {
                    item.ContentHashStatus = "変化";
                    item.ContentHashChanged = true;
                    hasIntrinsicTableChange = true;
                    if (item.RowDelta == 0) hasInPlaceMutation = true;
                }
                else
                {
                    item.ContentHashStatus = "一致";
                    item.ContentHashChanged = false;
                }
            }
            else
            {
                item.ContentHashStatus = "未取得(旧形式)";
            }

            ComputeRowChurnFromSketches(bt, at, item);
            if (item.HasRowChurn && (item.UpdatedRows > 0 || (item.RowDelta >= 0 && item.DeletedRows > 0)))
            {
                hasIntrinsicTableChange = true;
                hasInPlaceMutation = true;
                AddDiff(item, "-", "内容ハッシュ", "行レベル変動内訳(" + (item.RowChurnExact ? "完全計測" : "Bottom-K推定") + ")",
                    bt.RowCount.ToString("#,##0") + "件", at.RowCount.ToString("#,##0") + "件",
                    item.ChurnSummary, (item.DeletedRows > 0 && item.RowDelta >= 0) ? "注意" : "通常");
            }

            // 3.1 Table-level intrinsic comparisons (EXCLUDING No, RowShare, SizeShare)
            if (item.RowDelta != 0)
            {
                hasIntrinsicTableChange = true;
                string sev = item.RowDelta < 0 ? "警告" : "通常";
                AddDiff(item, "-", "テーブル指標", "レコード数",
                    bt.RowCount.ToString("#,##0"), at.RowCount.ToString("#,##0"),
                    FormatSignedInt(item.RowDelta) + "件", sev);
            }
            if (item.ColDelta != 0)
            {
                hasIntrinsicTableChange = true;
                hasSchemaChange = true;
                hasInPlaceMutation = true;
                AddDiff(item, "-", "テーブル指標", "列数",
                    bt.ColumnCount.ToString("#,##0"), at.ColumnCount.ToString("#,##0"),
                    FormatSignedInt(item.ColDelta) + "列", "警告");
            }
            if (!NumbersNearlyEqual(bt.AvgRowBytes, at.AvgRowBytes, 0.05))
            {
                hasIntrinsicTableChange = true;
                AddDiff(item, "-", "テーブル指標", "平均行長(B)",
                    bt.AvgRowBytes.ToString("#,##0.0"), at.AvgRowBytes.ToString("#,##0.0"),
                    FormatSignedDouble(at.AvgRowBytes - bt.AvgRowBytes, "0.0") + " B", "通常");
            }
            if (bt.MaxRowBytes != at.MaxRowBytes)
            {
                hasIntrinsicTableChange = true;
                if (item.RowDelta > 0 && at.MaxRowBytes < bt.MaxRowBytes) hasInPlaceMutation = true;
                AddDiff(item, "-", "テーブル指標", "最大行長(B)",
                    bt.MaxRowBytes.ToString("#,##0"), at.MaxRowBytes.ToString("#,##0"),
                    FormatSignedInt(at.MaxRowBytes - bt.MaxRowBytes) + " B", "通常");
            }
            if (!NumbersNearlyEqual(bt.EstInnoDbKB, at.EstInnoDbKB, 0.05))
            {
                hasIntrinsicTableChange = true;
                AddDiff(item, "-", "テーブル指標", "推定InnoDB(KB)",
                    bt.EstInnoDbKB.ToString("#,##0.0"), at.EstInnoDbKB.ToString("#,##0.0"),
                    FormatSignedDouble(at.EstInnoDbKB - bt.EstInnoDbKB, "0.0") + " KB", "通常");
            }
            if (!NumbersNearlyEqual(bt.FileSizeKB, at.FileSizeKB, 0.005))
            {
                bool encodingDiffOnly = !string.Equals(bt.EncodingName, at.EncodingName, StringComparison.OrdinalIgnoreCase) &&
                                        Math.Abs(at.FileSizeKB - bt.FileSizeKB) <= 0.015;
                if (!encodingDiffOnly)
                {
                    hasIntrinsicTableChange = true;
                    AddDiff(item, "-", "テーブル指標", "CSVサイズ(KB)",
                        bt.FileSizeKB.ToString("#,##0.00"), at.FileSizeKB.ToString("#,##0.00"),
                        FormatSignedDouble(at.FileSizeKB - bt.FileSizeKB, "0.00") + " KB", "通常");
                }
            }
            if (!NumbersNearlyEqual(bt.AvgNullRate, at.AvgNullRate, 1e-6))
            {
                hasIntrinsicTableChange = true;
                AddDiff(item, "-", "テーブル指標", "平均NULL率",
                    bt.AvgNullRate.ToString("0.0%"), at.AvgNullRate.ToString("0.0%"),
                    FormatSignedPercent(at.AvgNullRate - bt.AvgNullRate), "通常");
            }
            if (!string.Equals(NormalizeDash(bt.WarningSummary), NormalizeDash(at.WarningSummary), StringComparison.Ordinal))
            {
                string bw = string.IsNullOrEmpty(bt.WarningSummary) ? "正常" : bt.WarningSummary;
                string aw = string.IsNullOrEmpty(at.WarningSummary) ? "正常" : at.WarningSummary;
                if (bw != aw)
                {
                    hasIntrinsicTableChange = true;
                    bool newWarn = (aw != "正常");
                    AddDiff(item, "-", "テーブル指標", "警告・注意点", bw, aw, newWarn ? "警告内容変化" : "警告解消", newWarn ? "注意" : "通常");
                    if (newWarn && bw == "正常")
                    {
                        alertParts.Add("新規警告発生(" + aw + ")");
                    }
                }
            }

            // 3.1B DirectDB Physical InnoDB / AUTO_INCREMENT / DDL / Index / Declared FK comparisons
            bool bothHaveDbMeta = bt.DbMeta != null && bt.DbMeta.HasMetadata && at.DbMeta != null && at.DbMeta.HasMetadata;
            if (bothHaveDbMeta)
            {
                item.HasDbPhysicalMeta = true;
                item.BeforeDataLengthKB = Math.Round(bt.DbMeta.DataLength / 1024.0, 1);
                item.AfterDataLengthKB = Math.Round(at.DbMeta.DataLength / 1024.0, 1);
                item.DataLengthDeltaKB = Math.Round(item.AfterDataLengthKB - item.BeforeDataLengthKB, 1);
                item.BeforeIndexLengthKB = Math.Round(bt.DbMeta.IndexLength / 1024.0, 1);
                item.AfterIndexLengthKB = Math.Round(at.DbMeta.IndexLength / 1024.0, 1);
                item.IndexLengthDeltaKB = Math.Round(item.AfterIndexLengthKB - item.BeforeIndexLengthKB, 1);
                item.BeforeDataFreeKB = Math.Round(bt.DbMeta.DataFree / 1024.0, 1);
                item.AfterDataFreeKB = Math.Round(at.DbMeta.DataFree / 1024.0, 1);
                item.DataFreeDeltaKB = Math.Round(item.AfterDataFreeKB - item.BeforeDataFreeKB, 1);
                item.BeforeUpdateTime = bt.DbMeta.UpdateTime;
                item.AfterUpdateTime = at.DbMeta.UpdateTime;

                if (bt.DbMeta.DataLength != at.DbMeta.DataLength)
                {
                    AddDiff(item, "-", "物理DB指標", "InnoDB実データ容量(KB)",
                        item.BeforeDataLengthKB.ToString("#,##0.0"), item.AfterDataLengthKB.ToString("#,##0.0"),
                        FormatSignedDouble(item.DataLengthDeltaKB, "0.0") + " KB", "通常");
                }
                if (bt.DbMeta.IndexLength != at.DbMeta.IndexLength)
                {
                    AddDiff(item, "-", "物理DB指標", "InnoDB実インデックス容量(KB)",
                        item.BeforeIndexLengthKB.ToString("#,##0.0"), item.AfterIndexLengthKB.ToString("#,##0.0"),
                        FormatSignedDouble(item.IndexLengthDeltaKB, "0.0") + " KB", "通常");
                }
                if (bt.DbMeta.DataFree != at.DbMeta.DataFree)
                {
                    string fSev = (item.DataFreeDeltaKB >= 1024.0) ? "注意" : "通常";
                    AddDiff(item, "-", "物理DB指標", "InnoDB空き領域/断片化(KB)",
                        item.BeforeDataFreeKB.ToString("#,##0.0"), item.AfterDataFreeKB.ToString("#,##0.0"),
                        FormatSignedDouble(item.DataFreeDeltaKB, "0.0") + " KB", fSev);
                }

                if (bt.DbMeta.HasAutoIncrement && at.DbMeta.HasAutoIncrement)
                {
                    item.HasAutoIncrement = true;
                    item.BeforeAutoIncrement = bt.DbMeta.AutoIncrement;
                    item.AfterAutoIncrement = at.DbMeta.AutoIncrement;
                    item.AutoIncrementDelta = at.DbMeta.AutoIncrement - bt.DbMeta.AutoIncrement;
                    if (item.AutoIncrementDelta != 0)
                    {
                        hasIntrinsicTableChange = true;
                        long netAdded = Math.Max(0L, item.RowDelta);
                        if (item.AutoIncrementDelta > netAdded)
                        {
                            item.AutoIncrementGap = item.AutoIncrementDelta - netAdded;
                            hasInPlaceMutation = true;
                            string gapDesc = string.Format("採番 +{0:#,##0} に対し純増 +{1:#,##0} (差 {2:#,##0}: ROLLBACKまたはINSERT後DELETEの可能性)",
                                item.AutoIncrementDelta, netAdded, item.AutoIncrementGap);
                            AddDiff(item, "-", "物理DB指標", "AUTO_INCREMENT進行(ギャップ検出)",
                                bt.DbMeta.AutoIncrement.ToString("#,##0"), at.DbMeta.AutoIncrement.ToString("#,##0"),
                                gapDesc, "注意");
                            alertParts.Add(string.Format("AUTO_INCREMENTギャップ({0:#,##0}件)", item.AutoIncrementGap));
                        }
                        else
                        {
                            AddDiff(item, "-", "物理DB指標", "AUTO_INCREMENT",
                                bt.DbMeta.AutoIncrement.ToString("#,##0"), at.DbMeta.AutoIncrement.ToString("#,##0"),
                                FormatSignedInt(item.AutoIncrementDelta), "通常");
                        }
                    }
                }

                if (!string.Equals(bt.DbMeta.Engine, at.DbMeta.Engine, StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(bt.DbMeta.RowFormat, at.DbMeta.RowFormat, StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(bt.DbMeta.TableCollation, at.DbMeta.TableCollation, StringComparison.OrdinalIgnoreCase))
                {
                    hasIntrinsicTableChange = true;
                    hasSchemaChange = true;
                    hasInPlaceMutation = true;
                    string bCfg = string.Format("{0}/{1}/{2}", bt.DbMeta.Engine, bt.DbMeta.RowFormat, bt.DbMeta.TableCollation);
                    string aCfg = string.Format("{0}/{1}/{2}", at.DbMeta.Engine, at.DbMeta.RowFormat, at.DbMeta.TableCollation);
                    AddDiff(item, "-", "DDLスキーマ", "テーブル定義(Engine/Format/Collation)", bCfg, aCfg, "テーブル定義変更", "警告");
                    AddDbSchemaDiff(parentRes, item.TableName, "テーブル設定", item.TableName, "変更", bCfg, aCfg, "警告", "Engine/RowFormat/Collation変更");
                }

                // Compare Indexes
                CompareTableIndexes(bt.DbMeta, at.DbMeta, item, parentRes, ref hasIntrinsicTableChange, ref hasSchemaChange, ref hasInPlaceMutation);
                // Compare Declared Foreign Keys
                CompareTableDeclaredFks(bt.DbMeta, at.DbMeta, item, parentRes, ref hasIntrinsicTableChange, ref hasSchemaChange, ref hasInPlaceMutation);
            }

            // 3.2 Column-level comparison (preserving schema order)
            List<string> orderedColNames = new List<string>();
            HashSet<string> seenCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (SnapshotColumn ac in GetOrderedColumns(at)) { if (seenCols.Add(ac.ColumnName)) orderedColNames.Add(ac.ColumnName); }
            foreach (SnapshotColumn bc in GetOrderedColumns(bt)) { if (seenCols.Add(bc.ColumnName)) orderedColNames.Add(bc.ColumnName); }

            Dictionary<string, List<ColumnDiffItem>> diffsByCol = new Dictionary<string, List<ColumnDiffItem>>(StringComparer.OrdinalIgnoreCase);

            foreach (string colName in orderedColNames)
            {
                SnapshotColumn bc = null;
                SnapshotColumn ac = null;
                bt.Columns.TryGetValue(colName, out bc);
                at.Columns.TryGetValue(colName, out ac);

                if (bc == null && ac != null)
                {
                    hasIntrinsicTableChange = true;
                    hasSchemaChange = true;
                    hasInPlaceMutation = true;
                    changedColSet.Add(ac.ColumnName);
                    schemaChangedColSet.Add(ac.ColumnName);
                    item.AddedColumnCount++;
                    ColumnDiffItem df = AddDiff(item, ac.ColumnName, "スキーマ追加", "カラム追加", "なし", FormatDdlKeyDisplay(ac) + " (" + ac.InferredType + ")", "新規カラム追加", "注意");
                    TrackColDiff(diffsByCol, ac.ColumnName, df);
                    if (ac.DbMeta != null)
                    {
                        AddDbSchemaDiff(parentRes, item.TableName, "カラムDDL", ac.ColumnName, "追加", "-", FormatDdlKeyDisplay(ac), "注意", "テスト開始後にカラム追加");
                    }
                    continue;
                }
                if (bc != null && ac == null)
                {
                    hasIntrinsicTableChange = true;
                    hasSchemaChange = true;
                    hasInPlaceMutation = true;
                    changedColSet.Add(bc.ColumnName);
                    schemaChangedColSet.Add(bc.ColumnName);
                    item.RemovedColumnCount++;
                    ColumnDiffItem df = AddDiff(item, bc.ColumnName, "スキーマ削除", "カラム削除", FormatDdlKeyDisplay(bc) + " (" + bc.InferredType + ")", "なし", "カラム削除", "警告");
                    TrackColDiff(diffsByCol, bc.ColumnName, df);
                    alertParts.Add("カラム削除:" + bc.ColumnName);
                    if (bc.DbMeta != null)
                    {
                        AddDbSchemaDiff(parentRes, item.TableName, "カラムDDL", bc.ColumnName, "削除", FormatDdlKeyDisplay(bc), "-", "警告", "テスト開始後にカラム削除");
                    }
                    continue;
                }

                bool colChanged = false;
                bool colSchemaChanged = false;
                bool colInPlace = false;
                int prevDiffCount = item.ColumnDiffs.Count;

                // DDL metadata comparison (when both DirectDB snapshots have DbMeta)
                if (bc.DbMeta != null && ac.DbMeta != null)
                {
                    string bDdl = FormatDdlKeyDisplay(bc);
                    string aDdl = FormatDdlKeyDisplay(ac);
                    if (!string.Equals(bc.DbMeta.ColumnType, ac.DbMeta.ColumnType, StringComparison.OrdinalIgnoreCase) ||
                        !string.Equals(bc.DbMeta.IsNullable, ac.DbMeta.IsNullable, StringComparison.OrdinalIgnoreCase) ||
                        !string.Equals(bc.DbMeta.ColumnKey, ac.DbMeta.ColumnKey, StringComparison.OrdinalIgnoreCase) ||
                        !string.Equals(bc.DbMeta.ColumnDefault, ac.DbMeta.ColumnDefault, StringComparison.Ordinal) ||
                        !string.Equals(bc.DbMeta.Extra, ac.DbMeta.Extra, StringComparison.OrdinalIgnoreCase))
                    {
                        colChanged = true;
                        colSchemaChanged = true;
                        colInPlace = true;
                        hasSchemaChange = true;
                        hasInPlaceMutation = true;
                        item.SchemaChangedColumnCount++;
                        string bFull = bDdl + (string.IsNullOrEmpty(bc.DbMeta.ColumnDefault) ? "" : (" DEFAULT=" + bc.DbMeta.ColumnDefault));
                        string aFull = aDdl + (string.IsNullOrEmpty(ac.DbMeta.ColumnDefault) ? "" : (" DEFAULT=" + ac.DbMeta.ColumnDefault));
                        AddDiff(item, ac.ColumnName, "DDLスキーマ", "MySQLカラム定義", bFull, aFull, "DDL変更", "警告");
                        AddDbSchemaDiff(parentRes, item.TableName, "カラムDDL", ac.ColumnName, "変更", bFull, aFull, "警告", "型・NULL制約・キー・デフォルト変更");
                    }
                }

                if (!string.Equals(bc.InferredType, ac.InferredType, StringComparison.Ordinal))
                {
                    colChanged = true;
                    colInPlace = true;
                    hasInPlaceMutation = true;
                    AddDiff(item, ac.ColumnName, "カラム統計", "推定データ型", bc.InferredType, ac.InferredType, "型推定変化", "注意");
                }
                if (bc.NullCount != ac.NullCount)
                {
                    colChanged = true;
                    if (item.RowDelta > 0 && ac.NullCount < bc.NullCount) { colInPlace = true; hasInPlaceMutation = true; }
                    AddDiff(item, ac.ColumnName, "カラム統計", "NULL件数",
                        bc.NullCount.ToString("#,##0"), ac.NullCount.ToString("#,##0"),
                        FormatSignedInt(ac.NullCount - bc.NullCount) + "件", "通常");
                }
                if (bc.DistinctCount != ac.DistinctCount)
                {
                    colChanged = true;
                    if (item.RowDelta > 0 && ac.DistinctCount < bc.DistinctCount) { colInPlace = true; hasInPlaceMutation = true; }
                    AddDiff(item, ac.ColumnName, "カラム統計", "ユニーク数",
                        bc.DistinctCount.ToString("#,##0"), ac.DistinctCount.ToString("#,##0"),
                        FormatSignedInt(ac.DistinctCount - bc.DistinctCount) + "種類", "通常");
                }
                if (!NumbersNearlyEqual(bc.AvgByteLen, ac.AvgByteLen, 0.05))
                {
                    colChanged = true;
                    AddDiff(item, ac.ColumnName, "カラム統計", "平均長(B)",
                        bc.AvgByteLen.ToString("#,##0.0"), ac.AvgByteLen.ToString("#,##0.0"),
                        FormatSignedDouble(ac.AvgByteLen - bc.AvgByteLen, "0.0") + " B", "通常");
                }
                if (bc.P95ByteLen != ac.P95ByteLen)
                {
                    colChanged = true;
                    AddDiff(item, ac.ColumnName, "カラム統計", "P95長(B)",
                        bc.P95ByteLen.ToString("#,##0"), ac.P95ByteLen.ToString("#,##0"),
                        FormatSignedInt(ac.P95ByteLen - bc.P95ByteLen) + " B", "通常");
                }
                if (bc.MaxByteLen != ac.MaxByteLen)
                {
                    colChanged = true;
                    if (item.RowDelta > 0 && ac.MaxByteLen < bc.MaxByteLen) { colInPlace = true; hasInPlaceMutation = true; }
                    string bSev = (ac.MaxByteLen >= 4000 && bc.MaxByteLen < 4000) ? "注意" : "通常";
                    AddDiff(item, ac.ColumnName, "カラム統計", "最大長(B)",
                        bc.MaxByteLen.ToString("#,##0"), ac.MaxByteLen.ToString("#,##0"),
                        FormatSignedInt(ac.MaxByteLen - bc.MaxByteLen) + " B", bSev);
                }
                if (!Top1ValuesEqual(bc.Top1Value, ac.Top1Value))
                {
                    colChanged = true;
                    AddDiff(item, ac.ColumnName, "カラム統計", "最頻値(Top1)", bc.Top1Value, ac.Top1Value, "最頻値変化", "通常");
                }
                if (!NumbersNearlyEqual(bc.Top1Share, ac.Top1Share, 1e-6))
                {
                    colChanged = true;
                    AddDiff(item, ac.ColumnName, "カラム統計", "Top1占有率",
                        bc.Top1Share.ToString("0.0%"), ac.Top1Share.ToString("0.0%"),
                        FormatSignedPercent(ac.Top1Share - bc.Top1Share), "通常");
                }
                if (!NumbersNearlyEqual(bc.Top3Share, ac.Top3Share, 1e-6))
                {
                    colChanged = true;
                    AddDiff(item, ac.ColumnName, "カラム統計", "Top3占有率",
                        bc.Top3Share.ToString("0.0%"), ac.Top3Share.ToString("0.0%"),
                        FormatSignedPercent(ac.Top3Share - bc.Top3Share), "通常");
                }
                if (!NumbersNearlyEqual(bc.GiniCoefficient, ac.GiniCoefficient, 0.0005))
                {
                    colChanged = true;
                    AddDiff(item, ac.ColumnName, "カラム統計", "ジニ係数",
                        bc.GiniCoefficient.ToString("0.000"), ac.GiniCoefficient.ToString("0.000"),
                        FormatSignedDouble(ac.GiniCoefficient - bc.GiniCoefficient, "0.000"), "通常");
                }
                if (bc.HasNormalizedEntropy && ac.HasNormalizedEntropy && !NumbersNearlyEqual(bc.NormalizedEntropy, ac.NormalizedEntropy, 0.0005))
                {
                    colChanged = true;
                    AddDiff(item, ac.ColumnName, "カラム統計", "正規化エントロピー",
                        bc.NormalizedEntropy.ToString("0.000"), ac.NormalizedEntropy.ToString("0.000"),
                        FormatSignedDouble(ac.NormalizedEntropy - bc.NormalizedEntropy, "0.000"), "通常");
                }
                if (!string.Equals(bc.SkewRisk, ac.SkewRisk, StringComparison.Ordinal))
                {
                    colChanged = true;
                    string sev = ac.SkewRisk.StartsWith("高リスク") ? "警告" : "注意";
                    AddDiff(item, ac.ColumnName, "カラム統計", "偏りリスク", bc.SkewRisk, ac.SkewRisk, "偏り評価変化", sev);
                }
                if (bc.HasCharQualityCounts && ac.HasCharQualityCounts)
                {
                    if (bc.WhitespaceOnlyCount != ac.WhitespaceOnlyCount)
                    {
                        colChanged = true;
                        AddDiff(item, ac.ColumnName, "カラム統計", "空白のみ件数",
                            bc.WhitespaceOnlyCount.ToString("#,##0"), ac.WhitespaceOnlyCount.ToString("#,##0"),
                            FormatSignedInt(ac.WhitespaceOnlyCount - bc.WhitespaceOnlyCount) + "件", ac.WhitespaceOnlyCount > bc.WhitespaceOnlyCount ? "注意" : "通常");
                    }
                    if (bc.MultilineCount != ac.MultilineCount)
                    {
                        colChanged = true;
                        AddDiff(item, ac.ColumnName, "カラム統計", "改行含件数",
                            bc.MultilineCount.ToString("#,##0"), ac.MultilineCount.ToString("#,##0"),
                            FormatSignedInt(ac.MultilineCount - bc.MultilineCount) + "件", "通常");
                    }
                }

                // Fallback top-value distribution & Min/Max comparison when detailed dist/num/dt stats are absent
                bool hasBothDist = bt.Distributions.ContainsKey(ac.ColumnName) && at.Distributions.ContainsKey(ac.ColumnName);
                if (!hasBothDist && bc.TopDistCounts.Count > 0 && ac.TopDistCounts.Count > 0)
                {
                    HashSet<string> topKeys = new HashSet<string>(StringComparer.Ordinal);
                    foreach (string k in bc.TopDistCounts.Keys) topKeys.Add(k);
                    foreach (string k in ac.TopDistCounts.Keys) topKeys.Add(k);
                    foreach (string k in topKeys)
                    {
                        long bCnt = 0;
                        long aCnt = 0;
                        bool inB = bc.TopDistCounts.TryGetValue(k, out bCnt);
                        bool inA = ac.TopDistCounts.TryGetValue(k, out aCnt);
                        if ((inB && inA && bCnt != aCnt) || (bc.DistinctCount <= 5 && ac.DistinctCount <= 5 && inB != inA))
                        {
                            colChanged = true;
                            if (item.RowDelta > 0 && aCnt < bCnt) { colInPlace = true; hasInPlaceMutation = true; }
                            string sev = "通常";
                            if (aCnt > bCnt && AnomalyValueRegex.IsMatch(k)) sev = (bCnt == 0) ? "警告" : "注意";
                            AddDiff(item, ac.ColumnName, "カテゴリ分布", "区分値: " + k,
                                bCnt.ToString("#,##0") + "件", aCnt.ToString("#,##0") + "件",
                                FormatSignedInt(aCnt - bCnt) + "件" + (sev != "通常" ? " (異常系値増加)" : ""), sev);
                        }
                    }
                }

                bool hasBothNum = bt.NumericStats.ContainsKey(ac.ColumnName) && at.NumericStats.ContainsKey(ac.ColumnName);
                bool hasBothDt = bt.DateTimeStats.ContainsKey(ac.ColumnName) && at.DateTimeStats.ContainsKey(ac.ColumnName);
                if (!hasBothNum && !hasBothDt &&
                    (!string.IsNullOrEmpty(bc.MinValueText) || !string.IsNullOrEmpty(ac.MinValueText)))
                {
                    bool isNumCol = ac.InferredType.IndexOf("整数", StringComparison.Ordinal) >= 0 || ac.InferredType.IndexOf("小数", StringComparison.Ordinal) >= 0;
                    bool isDtCol = ac.InferredType.IndexOf("日時", StringComparison.Ordinal) >= 0 || ac.InferredType.IndexOf("日付", StringComparison.Ordinal) >= 0;
                    if (isNumCol)
                    {
                        double bMin, aMin, bMax, aMax;
                        if (TryParseFormattedNumber(bc.MinValueText, out bMin) && TryParseFormattedNumber(ac.MinValueText, out aMin) && !NumbersNearlyEqual(bMin, aMin, 0.005))
                        {
                            colChanged = true;
                            if (item.RowDelta > 0 && aMin > bMin + 0.005) { colInPlace = true; hasInPlaceMutation = true; }
                            AddDiff(item, ac.ColumnName, "数値統計", "最小値", bc.MinValueText, ac.MinValueText, FormatSignedDouble(aMin - bMin, "0.##"), "通常");
                        }
                        if (TryParseFormattedNumber(bc.MaxValueText, out bMax) && TryParseFormattedNumber(ac.MaxValueText, out aMax) && !NumbersNearlyEqual(bMax, aMax, 0.005))
                        {
                            colChanged = true;
                            if (item.RowDelta > 0 && aMax < bMax - 0.005) { colInPlace = true; hasInPlaceMutation = true; }
                            AddDiff(item, ac.ColumnName, "数値統計", "最大値", bc.MaxValueText, ac.MaxValueText, FormatSignedDouble(aMax - bMax, "0.##"), "通常");
                        }
                    }
                    else if (isDtCol)
                    {
                        DateTime bMinDt, aMinDt, bMaxDt, aMaxDt;
                        if (TryParseComparableDate(bc.MinValueText, out bMinDt) && TryParseComparableDate(ac.MinValueText, out aMinDt) && Math.Abs((aMinDt - bMinDt).TotalSeconds) >= 1.0)
                        {
                            colChanged = true;
                            if (item.RowDelta > 0 && aMinDt > bMinDt) { colInPlace = true; hasInPlaceMutation = true; }
                            AddDiff(item, ac.ColumnName, "時系列・日時", "最古日時", bc.MinValueText, ac.MinValueText, "最古日時変化", "通常");
                        }
                        if (TryParseComparableDate(bc.MaxValueText, out bMaxDt) && TryParseComparableDate(ac.MaxValueText, out aMaxDt) && Math.Abs((aMaxDt - bMaxDt).TotalSeconds) >= 1.0)
                        {
                            colChanged = true;
                            if (item.RowDelta > 0 && aMaxDt < bMaxDt) { colInPlace = true; hasInPlaceMutation = true; }
                            AddDiff(item, ac.ColumnName, "時系列・日時", "最新日時", bc.MaxValueText, ac.MaxValueText, "最新日時更新", "通常");
                        }
                    }
                }

                // Exact Column Content Hash & Distinct Set Hash comparison (catches hidden middle-row/same-length updates!)
                bool bothHaveColHash = !string.IsNullOrEmpty(bc.ContentHash) && !string.IsNullOrEmpty(ac.ContentHash);
                if (bothHaveColHash && !string.Equals(bc.ContentHash, ac.ContentHash, StringComparison.Ordinal))
                {
                    bool statAlreadyDetected = colChanged;
                    colChanged = true;
                    if (item.RowDelta == 0) { colInPlace = true; hasInPlaceMutation = true; }
                    if (!statAlreadyDetected && item.RowDelta == 0)
                    {
                        AddDiff(item, ac.ColumnName, "内容ハッシュ", "カラム値マルチセットハッシュ",
                            bc.ContentHash.Substring(0, 12) + "...", ac.ContentHash.Substring(0, 12) + "...",
                            "統計指標不変のデータ更新検出 (同一長/同一件数の値書換)", "注意");
                    }
                }
                if (!string.IsNullOrEmpty(bc.DistinctSetHash) && !string.IsNullOrEmpty(ac.DistinctSetHash) &&
                    bc.DistinctCount == ac.DistinctCount &&
                    !string.Equals(bc.DistinctSetHash, ac.DistinctSetHash, StringComparison.Ordinal))
                {
                    colChanged = true;
                    colInPlace = true;
                    hasInPlaceMutation = true;
                    AddDiff(item, ac.ColumnName, "内容ハッシュ", "ユニーク値集合ハッシュ",
                        bc.DistinctSetHash.Substring(0, 12) + "...", ac.DistinctSetHash.Substring(0, 12) + "...",
                        string.Format("ユニーク数同一({0:#,##0}種類)だが構成値が変化", ac.DistinctCount), "注意");
                }

                for (int di = prevDiffCount; di < item.ColumnDiffs.Count; di++)
                {
                    TrackColDiff(diffsByCol, ac.ColumnName, item.ColumnDiffs[di]);
                }

                if (colChanged)
                {
                    hasIntrinsicTableChange = true;
                    changedColSet.Add(ac.ColumnName);
                    if (colSchemaChanged) schemaChangedColSet.Add(ac.ColumnName);
                    if (colInPlace) inPlaceColSet.Add(ac.ColumnName);
                }
            }

            // 3.3 Category distribution comparison (Full Top20 from snapshot or Sheet 04)
            HashSet<string> distCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string dc in at.Distributions.Keys) distCols.Add(dc);
            foreach (string dc in bt.Distributions.Keys) distCols.Add(dc);

            foreach (string colName in distCols)
            {
                Dictionary<string, SnapshotCategoryItem> bd = null;
                Dictionary<string, SnapshotCategoryItem> ad = null;
                bt.Distributions.TryGetValue(colName, out bd);
                at.Distributions.TryGetValue(colName, out ad);
                if (bd == null || ad == null) continue;

                HashSet<string> cats = new HashSet<string>(StringComparer.Ordinal);
                foreach (string k in bd.Keys) cats.Add(k);
                foreach (string k in ad.Keys) cats.Add(k);

                foreach (string catVal in cats)
                {
                    SnapshotCategoryItem bi = null;
                    SnapshotCategoryItem ai = null;
                    bd.TryGetValue(catVal, out bi);
                    ad.TryGetValue(catVal, out ai);

                    long bCnt = bi != null ? bi.Count : 0;
                    long aCnt = ai != null ? ai.Count : 0;
                    if (bCnt != aCnt)
                    {
                        hasIntrinsicTableChange = true;
                        changedColSet.Add(colName);
                        if (item.RowDelta > 0 && aCnt < bCnt)
                        {
                            hasInPlaceMutation = true;
                            inPlaceColSet.Add(colName);
                        }
                        string sev = "通常";
                        string extra = "";
                        if (aCnt > bCnt && AnomalyValueRegex.IsMatch(catVal))
                        {
                            sev = (bCnt == 0) ? "警告" : "注意";
                            extra = (bCnt == 0) ? " (異常系値新規出現)" : " (異常系値増加)";
                            if (bCnt == 0) alertParts.Add(string.Format("{0}に異常系値'{1}'出現({2:#,##0}件)", colName, catVal, aCnt));
                        }
                        ColumnDiffItem df = AddDiff(item, colName, "カテゴリ分布", "区分値: " + catVal,
                            bCnt.ToString("#,##0") + "件", aCnt.ToString("#,##0") + "件",
                            FormatSignedInt(aCnt - bCnt) + "件" + extra, sev);
                        TrackColDiff(diffsByCol, colName, df);
                    }
                }
            }

            // 3.4 Numeric statistics comparison (Snapshot or Sheet 05)
            HashSet<string> numCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string nc in at.NumericStats.Keys) numCols.Add(nc);
            foreach (string nc in bt.NumericStats.Keys) numCols.Add(nc);

            foreach (string colName in numCols)
            {
                SnapshotNumericStats bn = null;
                SnapshotNumericStats an = null;
                bt.NumericStats.TryGetValue(colName, out bn);
                at.NumericStats.TryGetValue(colName, out an);
                if (bn == null || an == null) continue;

                bool numChanged = false;
                int prevCnt = item.ColumnDiffs.Count;
                if (bn.HasSum && an.HasSum && !NumbersNearlyEqual(bn.Sum, an.Sum, 0.005))
                {
                    numChanged = true;
                    AddDiff(item, colName, "数値統計", "合計値(SUM)", bn.Sum.ToString("#,##0.##"), an.Sum.ToString("#,##0.##"), FormatSignedDouble(an.Sum - bn.Sum, "0.##"), "通常");
                }
                if (!NumbersNearlyEqual(bn.Min, an.Min, 0.005))
                {
                    numChanged = true;
                    if (item.RowDelta > 0 && an.Min > bn.Min + 0.005) { hasInPlaceMutation = true; inPlaceColSet.Add(colName); }
                    AddDiff(item, colName, "数値統計", "最小値", bn.Min.ToString("#,##0.##"), an.Min.ToString("#,##0.##"), FormatSignedDouble(an.Min - bn.Min, "0.##"), "通常");
                }
                if (!NumbersNearlyEqual(bn.Max, an.Max, 0.005))
                {
                    numChanged = true;
                    if (item.RowDelta > 0 && an.Max < bn.Max - 0.005) { hasInPlaceMutation = true; inPlaceColSet.Add(colName); }
                    AddDiff(item, colName, "数値統計", "最大値", bn.Max.ToString("#,##0.##"), an.Max.ToString("#,##0.##"), FormatSignedDouble(an.Max - bn.Max, "0.##"), "通常");
                }
                if (!NumbersNearlyEqual(bn.Mean, an.Mean, 0.005))
                {
                    numChanged = true;
                    AddDiff(item, colName, "数値統計", "平均値", bn.Mean.ToString("#,##0.##"), an.Mean.ToString("#,##0.##"), FormatSignedDouble(an.Mean - bn.Mean, "0.##"), "通常");
                }
                if (!NumbersNearlyEqual(bn.Median, an.Median, 0.005))
                {
                    numChanged = true;
                    AddDiff(item, colName, "数値統計", "中央値(P50)", bn.Median.ToString("#,##0.##"), an.Median.ToString("#,##0.##"), FormatSignedDouble(an.Median - bn.Median, "0.##"), "通常");
                }
                if (!NumbersNearlyEqual(bn.P25, an.P25, 0.005))
                {
                    numChanged = true;
                    AddDiff(item, colName, "数値統計", "P25", bn.P25.ToString("#,##0.##"), an.P25.ToString("#,##0.##"), FormatSignedDouble(an.P25 - bn.P25, "0.##"), "通常");
                }
                if (!NumbersNearlyEqual(bn.P75, an.P75, 0.005))
                {
                    numChanged = true;
                    AddDiff(item, colName, "数値統計", "P75", bn.P75.ToString("#,##0.##"), an.P75.ToString("#,##0.##"), FormatSignedDouble(an.P75 - bn.P75, "0.##"), "通常");
                }
                if (bn.HasP90 && an.HasP90 && !NumbersNearlyEqual(bn.P90, an.P90, 0.005))
                {
                    numChanged = true;
                    AddDiff(item, colName, "数値統計", "P90", bn.P90.ToString("#,##0.##"), an.P90.ToString("#,##0.##"), FormatSignedDouble(an.P90 - bn.P90, "0.##"), "通常");
                }
                if (!NumbersNearlyEqual(bn.P95, an.P95, 0.005))
                {
                    numChanged = true;
                    AddDiff(item, colName, "数値統計", "P95", bn.P95.ToString("#,##0.##"), an.P95.ToString("#,##0.##"), FormatSignedDouble(an.P95 - bn.P95, "0.##"), "通常");
                }
                if (!NumbersNearlyEqual(bn.P99, an.P99, 0.005))
                {
                    numChanged = true;
                    AddDiff(item, colName, "数値統計", "P99", bn.P99.ToString("#,##0.##"), an.P99.ToString("#,##0.##"), FormatSignedDouble(an.P99 - bn.P99, "0.##"), "通常");
                }
                if (!NumbersNearlyEqual(bn.StdDev, an.StdDev, 0.005))
                {
                    numChanged = true;
                    AddDiff(item, colName, "数値統計", "標準偏差", bn.StdDev.ToString("#,##0.##"), an.StdDev.ToString("#,##0.##"), FormatSignedDouble(an.StdDev - bn.StdDev, "0.##"), "通常");
                }
                if (bn.HasZeroRate && an.HasZeroRate && !NumbersNearlyEqual(bn.ZeroRate, an.ZeroRate, 1e-6))
                {
                    numChanged = true;
                    AddDiff(item, colName, "数値統計", "ゼロ率", bn.ZeroRate.ToString("0.0%"), an.ZeroRate.ToString("0.0%"), FormatSignedPercent(an.ZeroRate - bn.ZeroRate), "通常");
                }
                if (bn.HasNegativeRate && an.HasNegativeRate && !NumbersNearlyEqual(bn.NegativeRate, an.NegativeRate, 1e-6))
                {
                    numChanged = true;
                    string nSev = (an.NegativeRate > bn.NegativeRate + 0.01) ? "注意" : "通常";
                    AddDiff(item, colName, "数値統計", "負値率", bn.NegativeRate.ToString("0.0%"), an.NegativeRate.ToString("0.0%"), FormatSignedPercent(an.NegativeRate - bn.NegativeRate), nSev);
                }
                if (!NumbersNearlyEqual(bn.OutlierRate, an.OutlierRate, 1e-6))
                {
                    numChanged = true;
                    string oSev = (an.OutlierRate >= 0.05 && an.OutlierRate > bn.OutlierRate + 0.01) ? "注意" : "通常";
                    AddDiff(item, colName, "数値統計", "外れ値率", bn.OutlierRate.ToString("0.0%"), an.OutlierRate.ToString("0.0%"), FormatSignedPercent(an.OutlierRate - bn.OutlierRate), oSev);
                }

                for (int di = prevCnt; di < item.ColumnDiffs.Count; di++) TrackColDiff(diffsByCol, colName, item.ColumnDiffs[di]);
                if (numChanged)
                {
                    hasIntrinsicTableChange = true;
                    changedColSet.Add(colName);
                }
            }

            // 3.5 DateTime statistics comparison (Snapshot or Sheet 06)
            HashSet<string> dtCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string dc in at.DateTimeStats.Keys) dtCols.Add(dc);
            foreach (string dc in bt.DateTimeStats.Keys) dtCols.Add(dc);

            foreach (string colName in dtCols)
            {
                SnapshotDateTimeStats bd = null;
                SnapshotDateTimeStats ad = null;
                bt.DateTimeStats.TryGetValue(colName, out bd);
                at.DateTimeStats.TryGetValue(colName, out ad);
                if (bd == null || ad == null) continue;

                bool dtChanged = false;
                int prevCnt = item.ColumnDiffs.Count;
                if (!string.Equals(bd.MinDate, ad.MinDate, StringComparison.Ordinal))
                {
                    dtChanged = true;
                    if (item.RowDelta > 0 && string.CompareOrdinal(ad.MinDate, bd.MinDate) > 0) { hasInPlaceMutation = true; inPlaceColSet.Add(colName); }
                    AddDiff(item, colName, "時系列・日時", "最古日時", bd.MinDate, ad.MinDate, "最古日時変化", "通常");
                }
                if (!string.Equals(bd.MaxDate, ad.MaxDate, StringComparison.Ordinal))
                {
                    dtChanged = true;
                    if (item.RowDelta > 0 && string.CompareOrdinal(ad.MaxDate, bd.MaxDate) < 0) { hasInPlaceMutation = true; inPlaceColSet.Add(colName); }
                    AddDiff(item, colName, "時系列・日時", "最新日時", bd.MaxDate, ad.MaxDate, "最新日時更新", "通常");
                }
                if (!string.Equals(bd.PeakDay, ad.PeakDay, StringComparison.Ordinal) || bd.PeakDayCount != ad.PeakDayCount)
                {
                    dtChanged = true;
                    AddDiff(item, colName, "時系列・日時", "ピーク日/件数",
                        string.Format("{0} ({1:#,##0}件)", bd.PeakDay, bd.PeakDayCount),
                        string.Format("{0} ({1:#,##0}件)", ad.PeakDay, ad.PeakDayCount),
                        FormatSignedInt(ad.PeakDayCount - bd.PeakDayCount) + "件", "通常");
                }
                for (int di = prevCnt; di < item.ColumnDiffs.Count; di++) TrackColDiff(diffsByCol, colName, item.ColumnDiffs[di]);
                if (dtChanged)
                {
                    hasIntrinsicTableChange = true;
                    changedColSet.Add(colName);
                }
            }

            // If table-level RowMultisetHash changed with zero column diffs (pure cross-column row swap!), record explicit evidence
            if (item.ContentHashChanged && item.ColumnDiffs.Count == 0 && item.RowDelta == 0)
            {
                hasInPlaceMutation = true;
                AddDiff(item, "-", "内容ハッシュ", "行マルチセットハッシュ",
                    bt.RowMultisetHash.Substring(0, 12) + "...", at.RowMultisetHash.Substring(0, 12) + "...",
                    "各列の集計は一致するが、行単位の列値組合せが変化 (行間値入替を検出)", "注意");
            }

            // 3.6 Final Classification
            item.ChangedColumnCount = changedColSet.Count;
            item.HasInPlaceMutation = hasInPlaceMutation;
            List<string> sortedChangedCols = changedColSet.OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToList();
            List<string> unchangedCols = new List<string>();
            foreach (string cn in orderedColNames)
            {
                if (!changedColSet.Contains(cn) && bt.Columns.ContainsKey(cn) && at.Columns.ContainsKey(cn))
                {
                    unchangedCols.Add(cn);
                }
            }
            item.UnchangedColumnCount = unchangedCols.Count;
            if (sortedChangedCols.Count > 6)
            {
                item.ChangedColumnsSummary = string.Join(", ", sortedChangedCols.Take(6).ToArray()) + string.Format(" 他{0}列", sortedChangedCols.Count - 6);
            }
            else if (sortedChangedCols.Count > 0)
            {
                item.ChangedColumnsSummary = string.Join(", ", sortedChangedCols.ToArray());
            }
            else
            {
                item.ChangedColumnsSummary = "-";
            }

            if (unchangedCols.Count == 0)
            {
                item.UnchangedColumnsSummary = "-";
            }
            else if (sortedChangedCols.Count == 0)
            {
                item.UnchangedColumnsSummary = string.Format("全{0}列変更なし", unchangedCols.Count);
            }
            else if (unchangedCols.Count > 6)
            {
                item.UnchangedColumnsSummary = string.Join(", ", unchangedCols.Take(6).ToArray()) + string.Format(" 他{0}列", unchangedCols.Count - 6);
            }
            else
            {
                item.UnchangedColumnsSummary = string.Join(", ", unchangedCols.ToArray());
            }

            if (item.RowDelta > 0)
            {
                item.IsAdded = true;
                if (hasInPlaceMutation)
                {
                    item.PrimaryStatus = "追加・更新";
                    item.ChangeGroup = "追加・更新";
                    item.AddedType = "レコード追加・更新";
                    item.IsUpdated = true;
                    item.ChangeDetail = string.Format("レコード追加 ({0}件) ＋ 既存値/構造更新 ({1}列変更 / {2}列不変)", FormatSignedInt(item.RowDelta), item.ChangedColumnCount, item.UnchangedColumnCount);
                }
                else
                {
                    item.PrimaryStatus = "追加（レコード増加）";
                    item.ChangeGroup = "追加";
                    item.AddedType = "レコード追加";
                    item.ChangeDetail = string.Format("レコード追加 ({0}件 / CSV {1}KB)", FormatSignedInt(item.RowDelta), FormatSignedDouble(item.SizeDeltaKB, "0.00"));
                }
                summaryParts.Add(string.Format("レコード {0}件 ({1:#,##0}→{2:#,##0})", FormatSignedInt(item.RowDelta), bt.RowCount, at.RowCount));
                if (item.HasRowChurn && (item.UpdatedRows > 0 || item.DeletedRows > 0)) summaryParts.Add(item.ChurnSummary);
                if (hasSchemaChange)
                {
                    summaryParts.Add(string.Format("列数 {0}列", FormatSignedInt(item.ColDelta)));
                    alertParts.Add("スキーマ変更あり");
                }
                if (item.ChangedColumnCount > 0)
                {
                    summaryParts.Add(string.Format("変化カラム {0}列 ({1})", item.ChangedColumnCount, item.ChangedColumnsSummary));
                }
            }
            else if (item.RowDelta < 0)
            {
                item.PrimaryStatus = "更新（レコード減少）";
                item.ChangeGroup = "更新";
                item.IsUpdated = true;
                item.ChangeDetail = string.Format("レコード減少 ({0}件 / {1:#,##0}→{2:#,##0})", FormatSignedInt(item.RowDelta), bt.RowCount, at.RowCount);
                summaryParts.Add(string.Format("レコード減少 {0}件 ({1:#,##0}→{2:#,##0})", FormatSignedInt(item.RowDelta), bt.RowCount, at.RowCount));
                if (item.HasRowChurn && (item.InsertedRows > 0 || item.UpdatedRows > 0)) summaryParts.Add(item.ChurnSummary);
                if (item.ChangedColumnCount > 0) summaryParts.Add("変化カラム: " + item.ChangedColumnsSummary);
                alertParts.Add(at.RowCount == 0 ? "全レコード消失(0件化)" : string.Format("レコード減少({0}件)", FormatSignedInt(item.RowDelta)));
            }
            else if (hasIntrinsicTableChange)
            {
                item.PrimaryStatus = "更新（データ更新）";
                item.ChangeGroup = "更新";
                item.IsUpdated = true;
                if (hasSchemaChange)
                {
                    item.ChangeDetail = string.Format("件数変動なし / スキーマ・統計更新 ({0}列変更 / {1}列不変)", item.ChangedColumnCount, item.UnchangedColumnCount);
                    alertParts.Add("スキーマ変更あり");
                }
                else if (item.HasRowChurn && (item.InsertedRows > 0 || item.DeletedRows > 0))
                {
                    item.ChangeDetail = string.Format("件数変動なし / 行入替・更新検出 ({0})", item.ChurnSummary);
                }
                else if (item.ChangedColumnCount > 0)
                {
                    item.ChangeDetail = string.Format("件数変動なし / データ更新検出 ({0}列変更 / {1}列不変)", item.ChangedColumnCount, item.UnchangedColumnCount);
                }
                else
                {
                    item.ChangeDetail = "件数変動なし / テーブル指標・行組合せ変化";
                }
                if (item.ChangedColumnCount > 0) summaryParts.Add("変化カラム: " + item.ChangedColumnsSummary);
                int shown = 0;
                foreach (ColumnDiffItem d in item.ColumnDiffs)
                {
                    if (d.ColumnName == "-" && item.ChangedColumnCount > 0) continue;
                    summaryParts.Add(string.Format("{0}.{1}: {2}→{3}", d.ColumnName == "-" ? item.TableName : d.ColumnName, d.MetricName, d.BeforeValue, d.AfterValue));
                    shown++;
                    if (shown >= 3) break;
                }
            }
            else
            {
                item.PrimaryStatus = "変更なし";
                item.ChangeGroup = "変更なし";
                item.IsUnchanged = true;
                item.ChangeDetail = "変更なし";
                item.DiffSummary = bothHaveRowHash
                    ? "開始前後でレコード数・サイズ・全カラム統計・分布・数値・日時統計・行内容ハッシュすべて完全一致"
                    : "開始前後でレコード数・サイズ・全カラム統計・分布・数値・日時統計に差分なし";
                if (!item.HasRowChurn && at.RowCount >= 0)
                {
                    item.HasRowChurn = true;
                    item.RowChurnExact = true;
                    item.UnchangedRows = at.RowCount;
                    item.ChurnSummary = string.Format("無変更 {0:#,##0}件", at.RowCount);
                }
            }

            if (summaryParts.Count > 0)
            {
                item.DiffSummary = string.Join(" / ", summaryParts.ToArray());
            }
            if (alertParts.Count > 0)
            {
                item.HasAlert = true;
                item.AlertNote = string.Join(" / ", alertParts.ToArray());
            }
            else if (item.AfterWarning != "正常" && !string.IsNullOrEmpty(item.AfterWarning))
            {
                item.AlertNote = "継続警告: " + item.AfterWarning;
            }
            else
            {
                item.AlertNote = "正常";
            }

            // 3.7 Build complete per-column comparison matrix (BOTH Changed AND Unchanged columns!)
            foreach (string colName in orderedColNames)
            {
                SnapshotColumn bc = null;
                SnapshotColumn ac = null;
                bt.Columns.TryGetValue(colName, out bc);
                at.Columns.TryGetValue(colName, out ac);

                ColumnComparisonItem cc = new ColumnComparisonItem();
                cc.TableName = item.TableName;
                cc.ColumnName = ac != null ? ac.ColumnName : bc.ColumnName;
                cc.ColumnIndex = ac != null ? ac.ColumnIndex : bc.ColumnIndex;
                cc.TableStatus = item.PrimaryStatus;

                List<ColumnDiffItem> cDiffs = null;
                diffsByCol.TryGetValue(cc.ColumnName, out cDiffs);
                cc.DiffCount = cDiffs != null ? cDiffs.Count : 0;

                if (bc == null && ac != null)
                {
                    cc.ColumnStatus = "追加（新規カラム）";
                    cc.IsChanged = true;
                    cc.AfterType = ac.InferredType;
                    cc.TypeDisplay = "なし → " + ac.InferredType;
                    cc.DdlKeyDisplay = FormatDdlKeyDisplay(ac);
                    cc.AfterNonNull = ac.NonNullCount > 0 ? ac.NonNullCount : Math.Max(0L, ac.TotalRows - ac.NullCount);
                    cc.NonNullDelta = cc.AfterNonNull;
                    cc.AfterNullCount = ac.NullCount;
                    cc.NullCountDelta = ac.NullCount;
                    cc.NullChangeDisplay = string.Format("0 → {0:#,##0}", ac.NullCount);
                    cc.AfterDistinct = ac.DistinctCount;
                    cc.DistinctDelta = ac.DistinctCount;
                    cc.DistinctChangeDisplay = string.Format("0 → {0:#,##0}", ac.DistinctCount);
                    cc.ByteLenDisplay = string.Format("平均 {0:0.0} / 最大 {1:#,##0} B", ac.AvgByteLen, ac.MaxByteLen);
                    cc.MinValDisplay = string.IsNullOrEmpty(ac.MinValueText) ? "-" : ac.MinValueText;
                    cc.MaxValDisplay = string.IsNullOrEmpty(ac.MaxValueText) ? "-" : ac.MaxValueText;
                    cc.Top1Display = string.IsNullOrEmpty(ac.Top1Value) ? "-" : ac.Top1Value;
                    cc.ContentHashStatus = !string.IsNullOrEmpty(ac.ContentHash) ? "追加" : "未取得(旧形式)";
                    cc.ChangeSummary = "テスト開始後に新規カラム追加";
                    cc.Severity = "注意";
                }
                else if (bc != null && ac == null)
                {
                    cc.ColumnStatus = "削除（カラム消失）";
                    cc.IsChanged = true;
                    cc.BeforeType = bc.InferredType;
                    cc.TypeDisplay = bc.InferredType + " → 削除";
                    cc.DdlKeyDisplay = FormatDdlKeyDisplay(bc);
                    cc.BeforeNonNull = bc.NonNullCount > 0 ? bc.NonNullCount : Math.Max(0L, bc.TotalRows - bc.NullCount);
                    cc.NonNullDelta = -cc.BeforeNonNull;
                    cc.BeforeNullCount = bc.NullCount;
                    cc.NullCountDelta = -bc.NullCount;
                    cc.NullChangeDisplay = string.Format("{0:#,##0} → 0", bc.NullCount);
                    cc.BeforeDistinct = bc.DistinctCount;
                    cc.DistinctDelta = -bc.DistinctCount;
                    cc.DistinctChangeDisplay = string.Format("{0:#,##0} → 0", bc.DistinctCount);
                    cc.ByteLenDisplay = string.Format("平均 {0:0.0} / 最大 {1:#,##0} B", bc.AvgByteLen, bc.MaxByteLen);
                    cc.MinValDisplay = string.IsNullOrEmpty(bc.MinValueText) ? "-" : bc.MinValueText;
                    cc.MaxValDisplay = string.IsNullOrEmpty(bc.MaxValueText) ? "-" : bc.MaxValueText;
                    cc.Top1Display = string.IsNullOrEmpty(bc.Top1Value) ? "-" : bc.Top1Value;
                    cc.ContentHashStatus = !string.IsNullOrEmpty(bc.ContentHash) ? "削除" : "未取得(旧形式)";
                    cc.ChangeSummary = "テスト開始後にカラム削除・消失";
                    cc.Severity = "警告";
                }
                else
                {
                    bool isColChanged = changedColSet.Contains(ac.ColumnName);
                    bool isColSchema = schemaChangedColSet.Contains(ac.ColumnName);
                    bool isColInPlace = inPlaceColSet.Contains(ac.ColumnName);
                    bool bothH = !string.IsNullOrEmpty(bc.ContentHash) && !string.IsNullOrEmpty(ac.ContentHash);
                    bool hashChanged = bothH && !string.Equals(bc.ContentHash, ac.ContentHash, StringComparison.Ordinal);

                    cc.BeforeType = bc.InferredType;
                    cc.AfterType = ac.InferredType;
                    cc.TypeDisplay = FormatBeforeAfterText(bc.InferredType, ac.InferredType);
                    cc.DdlKeyDisplay = FormatBeforeAfterText(FormatDdlKeyDisplay(bc), FormatDdlKeyDisplay(ac));
                    cc.BeforeNonNull = bc.NonNullCount > 0 ? bc.NonNullCount : Math.Max(0L, bc.TotalRows - bc.NullCount);
                    cc.AfterNonNull = ac.NonNullCount > 0 ? ac.NonNullCount : Math.Max(0L, ac.TotalRows - ac.NullCount);
                    cc.NonNullDelta = cc.AfterNonNull - cc.BeforeNonNull;
                    cc.BeforeNullCount = bc.NullCount;
                    cc.AfterNullCount = ac.NullCount;
                    cc.NullCountDelta = ac.NullCount - bc.NullCount;
                    cc.NullChangeDisplay = bc.NullCount == ac.NullCount
                        ? bc.NullCount.ToString("#,##0")
                        : string.Format("{0:#,##0} → {1:#,##0} ({2})", bc.NullCount, ac.NullCount, FormatSignedInt(cc.NullCountDelta));
                    cc.BeforeDistinct = bc.DistinctCount;
                    cc.AfterDistinct = ac.DistinctCount;
                    cc.DistinctDelta = ac.DistinctCount - bc.DistinctCount;
                    cc.DistinctChangeDisplay = bc.DistinctCount == ac.DistinctCount
                        ? bc.DistinctCount.ToString("#,##0")
                        : string.Format("{0:#,##0} → {1:#,##0} ({2})", bc.DistinctCount, ac.DistinctCount, FormatSignedInt(cc.DistinctDelta));
                    cc.ByteLenDisplay = (NumbersNearlyEqual(bc.AvgByteLen, ac.AvgByteLen, 0.05) && bc.MaxByteLen == ac.MaxByteLen)
                        ? string.Format("平均 {0:0.0} / 最大 {1:#,##0} B", ac.AvgByteLen, ac.MaxByteLen)
                        : string.Format("平均 {0:0.0}→{1:0.0} / 最大 {2:#,##0}→{3:#,##0} B", bc.AvgByteLen, ac.AvgByteLen, bc.MaxByteLen, ac.MaxByteLen);
                    cc.MinValDisplay = FormatBeforeAfterText(bc.MinValueText, ac.MinValueText);
                    cc.MaxValDisplay = FormatBeforeAfterText(bc.MaxValueText, ac.MaxValueText);
                    cc.Top1Display = Top1ValuesEqual(bc.Top1Value, ac.Top1Value)
                        ? (string.IsNullOrEmpty(ac.Top1Value) ? "-" : ac.Top1Value)
                        : FormatBeforeAfterText(bc.Top1Value, ac.Top1Value);
                    cc.ContentHashStatus = bothH ? (hashChanged ? "変化" : "一致") : "未取得(旧形式)";

                    SnapshotNumericStats bnStat = null, anStat = null;
                    if (bt.NumericStats.TryGetValue(ac.ColumnName, out bnStat) && at.NumericStats.TryGetValue(ac.ColumnName, out anStat) && bnStat.HasSum && anStat.HasSum)
                    {
                        double sumD = anStat.Sum - bnStat.Sum;
                        cc.NumericSumDeltaDisplay = NumbersNearlyEqual(bnStat.Sum, anStat.Sum, 0.005) ? "0" : FormatSignedDouble(sumD, "0.##");
                    }

                    if (isColChanged)
                    {
                        cc.IsChanged = true;
                        bool hasDataDiff = hashChanged || (cDiffs != null && cDiffs.Any(x => x.Category != "DDLスキーマ"));
                        if (isColSchema && hasDataDiff) cc.ColumnStatus = "変更あり（スキーマ・データ）";
                        else if (isColSchema) cc.ColumnStatus = "変更あり（スキーマ変更）";
                        else if (item.RowDelta == 0 || isColInPlace || item.RowDelta < 0) cc.ColumnStatus = "変更あり（データ更新）";
                        else cc.ColumnStatus = "変更あり（追加・統計変化）";

                        string maxSev = "通常";
                        List<string> cSumParts = new List<string>();
                        if (cDiffs != null)
                        {
                            for (int di = 0; di < cDiffs.Count; di++)
                            {
                                if (cDiffs[di].Severity == "警告") maxSev = "警告";
                                else if (cDiffs[di].Severity == "注意" && maxSev != "警告") maxSev = "注意";
                                if (di < 3) cSumParts.Add(string.Format("{0}: {1}", cDiffs[di].MetricName, cDiffs[di].DeltaSummary));
                            }
                            if (cDiffs.Count > 3) cSumParts.Add(string.Format("他{0}指標", cDiffs.Count - 3));
                        }
                        if (cSumParts.Count == 0 && hashChanged) cSumParts.Add("値マルチセットハッシュ変化 (レコード追加/更新)");
                        cc.Severity = maxSev;
                        cc.ChangeSummary = string.Join(" / ", cSumParts.ToArray());
                    }
                    else
                    {
                        cc.IsUnchanged = true;
                        cc.ColumnStatus = "変更なし";
                        cc.Severity = "変化なし";
                        cc.ChangeSummary = bothH ? "型・統計・値ハッシュすべて完全一致" : "型・全統計指標に変化なし";
                    }
                }
                item.ColumnComparisons.Add(cc);
            }

            return item;
        }

        private static List<SnapshotColumn> GetOrderedColumns(SnapshotTable st)
        {
            if (st == null) return new List<SnapshotColumn>();
            if (st.ColumnList != null && st.ColumnList.Count > 0) return st.ColumnList;
            return st.Columns.Values.OrderBy(c => c.ColumnIndex).ThenBy(c => c.ColumnName, StringComparer.OrdinalIgnoreCase).ToList();
        }

        private static void TrackColDiff(Dictionary<string, List<ColumnDiffItem>> map, string colName, ColumnDiffItem diff)
        {
            List<ColumnDiffItem> list;
            if (!map.TryGetValue(colName, out list))
            {
                list = new List<ColumnDiffItem>();
                map[colName] = list;
            }
            list.Add(diff);
        }

        private static void AddDbSchemaDiff(AnalysisComparisonResult res, string tbl, string objType, string objName, string changeType, string beforeDef, string afterDef, string sev, string note)
        {
            if (res == null) return;
            DbSchemaDiffItem d = new DbSchemaDiffItem();
            d.TableName = tbl;
            d.ObjectType = objType;
            d.ObjectName = objName;
            d.ChangeType = changeType;
            d.BeforeDef = beforeDef ?? "-";
            d.AfterDef = afterDef ?? "-";
            d.Severity = sev ?? "注意";
            d.Note = note ?? "";
            res.DbSchemaDiffs.Add(d);
        }

        private static void CompareTableIndexes(DbTableMetadata bm, DbTableMetadata am, TableComparisonItem item, AnalysisComparisonResult res,
            ref bool hasIntrinsicTableChange, ref bool hasSchemaChange, ref bool hasInPlaceMutation)
        {
            Dictionary<string, DbIndexMetadata> bIdx = new Dictionary<string, DbIndexMetadata>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, DbIndexMetadata> aIdx = new Dictionary<string, DbIndexMetadata>(StringComparer.OrdinalIgnoreCase);
            foreach (DbIndexMetadata im in bm.Indexes) bIdx[im.IndexName] = im;
            foreach (DbIndexMetadata im in am.Indexes) aIdx[im.IndexName] = im;

            HashSet<string> allNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string k in aIdx.Keys) allNames.Add(k);
            foreach (string k in bIdx.Keys) allNames.Add(k);

            foreach (string idxName in allNames)
            {
                DbIndexMetadata bi = null, ai = null;
                bIdx.TryGetValue(idxName, out bi);
                aIdx.TryGetValue(idxName, out ai);
                if (bi == null && ai != null)
                {
                    hasIntrinsicTableChange = true; hasSchemaChange = true; hasInPlaceMutation = true;
                    string def = string.Format("{0} ({1}) [{2}]", ai.Columns, ai.NonUnique ? "INDEX" : "UNIQUE", ai.IndexType);
                    AddDiff(item, "-", "インデックス", "INDEX追加: " + ai.IndexName, "なし", def, "インデックス追加", "注意");
                    AddDbSchemaDiff(res, item.TableName, "インデックス", ai.IndexName, "追加", "-", def, "注意", "テスト後にインデックス追加");
                }
                else if (bi != null && ai == null)
                {
                    hasIntrinsicTableChange = true; hasSchemaChange = true; hasInPlaceMutation = true;
                    string def = string.Format("{0} ({1}) [{2}]", bi.Columns, bi.NonUnique ? "INDEX" : "UNIQUE", bi.IndexType);
                    AddDiff(item, "-", "インデックス", "INDEX削除: " + bi.IndexName, def, "なし", "インデックス削除", "警告");
                    AddDbSchemaDiff(res, item.TableName, "インデックス", bi.IndexName, "削除", def, "-", "警告", "テスト後にインデックス削除 (クエリ計画劣化注意)");
                }
                else if (bi != null && ai != null)
                {
                    if (!string.Equals(bi.Columns, ai.Columns, StringComparison.OrdinalIgnoreCase) || bi.NonUnique != ai.NonUnique)
                    {
                        hasIntrinsicTableChange = true; hasSchemaChange = true; hasInPlaceMutation = true;
                        string bDef = string.Format("{0} ({1})", bi.Columns, bi.NonUnique ? "INDEX" : "UNIQUE");
                        string aDef = string.Format("{0} ({1})", ai.Columns, ai.NonUnique ? "INDEX" : "UNIQUE");
                        AddDiff(item, "-", "インデックス", "INDEX構成変更: " + ai.IndexName, bDef, aDef, "インデックス定義変更", "警告");
                        AddDbSchemaDiff(res, item.TableName, "インデックス", ai.IndexName, "変更", bDef, aDef, "警告", "インデックス列構成変更");
                    }
                    else if (bi.Cardinality != ai.Cardinality && Math.Abs(ai.Cardinality - bi.Cardinality) >= Math.Max(10L, bi.Cardinality / 5))
                    {
                        AddDiff(item, "-", "インデックス", "INDEXカーディナリティ: " + ai.IndexName,
                            bi.Cardinality.ToString("#,##0"), ai.Cardinality.ToString("#,##0"),
                            FormatSignedInt(ai.Cardinality - bi.Cardinality), "通常");
                    }
                }
            }
        }

        private static void CompareTableDeclaredFks(DbTableMetadata bm, DbTableMetadata am, TableComparisonItem item, AnalysisComparisonResult res,
            ref bool hasIntrinsicTableChange, ref bool hasSchemaChange, ref bool hasInPlaceMutation)
        {
            Dictionary<string, DbForeignKeyMetadata> bFk = new Dictionary<string, DbForeignKeyMetadata>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, DbForeignKeyMetadata> aFk = new Dictionary<string, DbForeignKeyMetadata>(StringComparer.OrdinalIgnoreCase);
            foreach (DbForeignKeyMetadata f in bm.DeclaredFks) bFk[f.ConstraintName + ":" + f.ColumnName] = f;
            foreach (DbForeignKeyMetadata f in am.DeclaredFks) aFk[f.ConstraintName + ":" + f.ColumnName] = f;

            HashSet<string> allKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string k in aFk.Keys) allKeys.Add(k);
            foreach (string k in bFk.Keys) allKeys.Add(k);

            foreach (string fkKey in allKeys)
            {
                DbForeignKeyMetadata bf = null, af = null;
                bFk.TryGetValue(fkKey, out bf);
                aFk.TryGetValue(fkKey, out af);
                if (bf == null && af != null)
                {
                    hasIntrinsicTableChange = true; hasSchemaChange = true; hasInPlaceMutation = true;
                    string def = string.Format("{0} → {1}.{2}", af.ColumnName, af.ReferencedTableName, af.ReferencedColumnName);
                    AddDiff(item, af.ColumnName, "外部キー", "FK制約追加: " + af.ConstraintName, "なし", def, "外部キー追加", "注意");
                    AddDbSchemaDiff(res, item.TableName, "外部キー制約", af.ConstraintName, "追加", "-", def, "注意", "外部キー制約追加");
                }
                else if (bf != null && af == null)
                {
                    hasIntrinsicTableChange = true; hasSchemaChange = true; hasInPlaceMutation = true;
                    string def = string.Format("{0} → {1}.{2}", bf.ColumnName, bf.ReferencedTableName, bf.ReferencedColumnName);
                    AddDiff(item, bf.ColumnName, "外部キー", "FK制約削除: " + bf.ConstraintName, def, "なし", "外部キー削除", "警告");
                    AddDbSchemaDiff(res, item.TableName, "外部キー制約", bf.ConstraintName, "削除", def, "-", "警告", "外部キー制約削除");
                }
            }
        }

        private static int SeverityRank(string sev)
        {
            if (sev == "高") return 3;
            if (sev == "中") return 2;
            if (sev == "低") return 1;
            return 0;
        }

        private static void CompareWarnings(AnalysisWorkbookSnapshot before, AnalysisWorkbookSnapshot after, AnalysisComparisonResult res)
        {
            if (before.Warnings.Count == 0 && after.Warnings.Count == 0) return;
            if (before.HasEmbeddedSnapshot != after.HasEmbeddedSnapshot && (before.Warnings.Count == 0 || after.Warnings.Count == 0)) return;

            Dictionary<string, SnapshotWarning> bMap = new Dictionary<string, SnapshotWarning>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, SnapshotWarning> aMap = new Dictionary<string, SnapshotWarning>(StringComparer.OrdinalIgnoreCase);
            foreach (SnapshotWarning w in before.Warnings) bMap[w.Category + "|" + w.TableName + "|" + w.ColumnName] = w;
            foreach (SnapshotWarning w in after.Warnings) aMap[w.Category + "|" + w.TableName + "|" + w.ColumnName] = w;

            HashSet<string> keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string k in aMap.Keys) keys.Add(k);
            foreach (string k in bMap.Keys) keys.Add(k);

            foreach (string k in keys)
            {
                SnapshotWarning bw = null, aw = null;
                bMap.TryGetValue(k, out bw);
                aMap.TryGetValue(k, out aw);

                WarningComparisonItem wi = new WarningComparisonItem();
                if (bw == null && aw != null)
                {
                    wi.Status = "新規発生";
                    wi.Severity = aw.Severity;
                    wi.Category = aw.Category;
                    wi.TableName = aw.TableName;
                    wi.ColumnName = aw.ColumnName;
                    wi.BeforeMetric = "なし";
                    wi.AfterMetric = aw.MetricValue;
                    wi.PerfImpact = aw.PerfImpact;
                    wi.Recommendation = aw.Recommendation;
                    wi.IsAlert = true;
                    res.NewWarningCount++;
                }
                else if (bw != null && aw == null)
                {
                    wi.Status = "解消";
                    wi.Severity = bw.Severity;
                    wi.Category = bw.Category;
                    wi.TableName = bw.TableName;
                    wi.ColumnName = bw.ColumnName;
                    wi.BeforeMetric = bw.MetricValue;
                    wi.AfterMetric = "解消";
                    wi.PerfImpact = bw.PerfImpact;
                    wi.Recommendation = "テスト後に解消済み";
                    res.ResolvedWarningCount++;
                }
                else
                {
                    int bRank = SeverityRank(bw.Severity);
                    int aRank = SeverityRank(aw.Severity);
                    wi.Severity = aw.Severity;
                    wi.Category = aw.Category;
                    wi.TableName = aw.TableName;
                    wi.ColumnName = aw.ColumnName;
                    wi.BeforeMetric = bw.MetricValue;
                    wi.AfterMetric = aw.MetricValue;
                    wi.PerfImpact = aw.PerfImpact;
                    wi.Recommendation = aw.Recommendation;
                    if (aRank > bRank)
                    {
                        wi.Status = "悪化";
                        wi.IsAlert = true;
                        res.WorsenedWarningCount++;
                    }
                    else if (aRank < bRank)
                    {
                        wi.Status = "改善";
                        res.OngoingWarningCount++;
                    }
                    else if (!string.Equals(bw.MetricValue, aw.MetricValue, StringComparison.Ordinal))
                    {
                        wi.Status = "継続 (指標変化)";
                        wi.IsAlert = (aw.Severity == "高");
                        res.OngoingWarningCount++;
                    }
                    else
                    {
                        wi.Status = "継続";
                        res.OngoingWarningCount++;
                    }
                }
                res.WarningComparisons.Add(wi);
            }

            res.WarningComparisons = res.WarningComparisons
                .OrderBy(w => w.Status == "新規発生" ? 0 : (w.Status == "悪化" ? 1 : (w.Status.StartsWith("継続") ? 2 : (w.Status == "改善" ? 3 : 4))))
                .ThenByDescending(w => SeverityRank(w.Severity))
                .ThenBy(w => w.TableName, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static void CompareRelationships(AnalysisWorkbookSnapshot before, AnalysisWorkbookSnapshot after, AnalysisComparisonResult res)
        {
            if (before.Relationships.Count == 0 && after.Relationships.Count == 0) return;
            if (before.HasEmbeddedSnapshot != after.HasEmbeddedSnapshot && (before.Relationships.Count == 0 || after.Relationships.Count == 0)) return;

            Dictionary<string, SnapshotRelationship> bMap = new Dictionary<string, SnapshotRelationship>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, SnapshotRelationship> aMap = new Dictionary<string, SnapshotRelationship>(StringComparer.OrdinalIgnoreCase);
            foreach (SnapshotRelationship r in before.Relationships)
            {
                string rk = r.ParentTable + "." + r.ParentColumn + " → " + r.ChildTable + "." + r.ChildColumn;
                bMap[rk] = r;
                res.BeforeTotalOrphanRows += r.OrphanRows;
            }
            foreach (SnapshotRelationship r in after.Relationships)
            {
                string rk = r.ParentTable + "." + r.ParentColumn + " → " + r.ChildTable + "." + r.ChildColumn;
                aMap[rk] = r;
                res.AfterTotalOrphanRows += r.OrphanRows;
            }
            res.OrphanRowDelta = res.AfterTotalOrphanRows - res.BeforeTotalOrphanRows;

            HashSet<string> keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string k in aMap.Keys) keys.Add(k);
            foreach (string k in bMap.Keys) keys.Add(k);

            foreach (string rk in keys)
            {
                SnapshotRelationship br = null, ar = null;
                bMap.TryGetValue(rk, out br);
                aMap.TryGetValue(rk, out ar);

                RelationshipComparisonItem item = new RelationshipComparisonItem();
                item.RelationKey = rk;
                if (br == null && ar != null)
                {
                    item.Status = ar.OrphanRows > 0 ? "悪化（要注意）" : "新規検出";
                    item.ParentTable = ar.ParentTable; item.ParentColumn = ar.ParentColumn;
                    item.ChildTable = ar.ChildTable; item.ChildColumn = ar.ChildColumn;
                    item.AfterMatchRate = ar.MatchRate; item.MatchRateDelta = ar.MatchRate;
                    item.AfterOrphanRows = ar.OrphanRows; item.OrphanDelta = ar.OrphanRows;
                    item.AfterFanOutAvg = ar.FanOutAvg; item.AfterFanOutP95 = ar.FanOutP95;
                    item.AfterFanOutMax = ar.FanOutMax; item.FanOutMaxDelta = ar.FanOutMax;
                    item.AfterRisk = ar.RiskLevel;
                    item.Note = ar.Note;
                    item.HasAlert = ar.OrphanRows > 0 || ar.RiskLevel.StartsWith("高リスク");
                }
                else if (br != null && ar == null)
                {
                    item.Status = "消失";
                    item.ParentTable = br.ParentTable; item.ParentColumn = br.ParentColumn;
                    item.ChildTable = br.ChildTable; item.ChildColumn = br.ChildColumn;
                    item.BeforeMatchRate = br.MatchRate; item.MatchRateDelta = -br.MatchRate;
                    item.BeforeOrphanRows = br.OrphanRows; item.OrphanDelta = -br.OrphanRows;
                    item.BeforeFanOutAvg = br.FanOutAvg; item.BeforeFanOutP95 = br.FanOutP95;
                    item.BeforeFanOutMax = br.FanOutMax; item.FanOutMaxDelta = -br.FanOutMax;
                    item.BeforeRisk = br.RiskLevel;
                    item.AfterRisk = "未検出";
                    item.Note = "テスト開始後にFKリレーション未検出";
                    item.HasAlert = true;
                }
                else
                {
                    item.ParentTable = ar.ParentTable; item.ParentColumn = ar.ParentColumn;
                    item.ChildTable = ar.ChildTable; item.ChildColumn = ar.ChildColumn;
                    item.BeforeMatchRate = br.MatchRate; item.AfterMatchRate = ar.MatchRate; item.MatchRateDelta = ar.MatchRate - br.MatchRate;
                    item.BeforeOrphanRows = br.OrphanRows; item.AfterOrphanRows = ar.OrphanRows; item.OrphanDelta = ar.OrphanRows - br.OrphanRows;
                    item.BeforeFanOutAvg = br.FanOutAvg; item.AfterFanOutAvg = ar.FanOutAvg;
                    item.BeforeFanOutP95 = br.FanOutP95; item.AfterFanOutP95 = ar.FanOutP95;
                    item.BeforeFanOutMax = br.FanOutMax; item.AfterFanOutMax = ar.FanOutMax; item.FanOutMaxDelta = ar.FanOutMax - br.FanOutMax;
                    item.BeforeRisk = br.RiskLevel; item.AfterRisk = ar.RiskLevel;

                    bool worsened = item.OrphanDelta > 0 ||
                                    item.MatchRateDelta < -0.005 ||
                                    (ar.RiskLevel.StartsWith("高リスク") && !br.RiskLevel.StartsWith("高リスク")) ||
                                    (item.FanOutMaxDelta >= Math.Max(50L, br.FanOutMax));
                    bool improved = item.OrphanDelta < 0 || (!ar.RiskLevel.StartsWith("高リスク") && br.RiskLevel.StartsWith("高リスク"));
                    bool anyDiff = item.OrphanDelta != 0 || !NumbersNearlyEqual(br.MatchRate, ar.MatchRate, 1e-4) ||
                                   item.FanOutMaxDelta != 0 || !NumbersNearlyEqual(br.FanOutAvg, ar.FanOutAvg, 0.05) ||
                                   !string.Equals(br.RiskLevel, ar.RiskLevel, StringComparison.Ordinal);

                    if (worsened)
                    {
                        item.Status = "悪化（要注意）";
                        item.HasAlert = true;
                    }
                    else if (improved)
                    {
                        item.Status = "改善";
                    }
                    else if (anyDiff)
                    {
                        item.Status = "変化あり";
                    }
                    else
                    {
                        item.Status = "変化なし";
                    }
                    List<string> nParts = new List<string>();
                    if (item.OrphanDelta > 0) nParts.Add(string.Format("孤立FK +{0:#,##0}件増加 (参照整合性破壊注意)", item.OrphanDelta));
                    else if (item.OrphanDelta < 0) nParts.Add(string.Format("孤立FK {0:#,##0}件減少", item.OrphanDelta));
                    if (item.FanOutMaxDelta != 0) nParts.Add(string.Format("最大ファンアウト {0:#,##0}→{1:#,##0} ({2})", br.FanOutMax, ar.FanOutMax, FormatSignedInt(item.FanOutMaxDelta)));
                    if (!string.IsNullOrEmpty(ar.MaxFanOutKey) && ar.MaxFanOutKey != br.MaxFanOutKey) nParts.Add("最大親キー変化:" + ar.MaxFanOutKey);
                    item.Note = nParts.Count > 0 ? string.Join(" / ", nParts.ToArray()) : (string.IsNullOrEmpty(ar.Note) ? "変化なし" : ar.Note);
                }
                res.RelationshipComparisons.Add(item);
            }

            res.RelationshipComparisons = res.RelationshipComparisons
                .OrderBy(r => r.HasAlert ? 0 : (r.Status == "変化なし" ? 2 : 1))
                .ThenByDescending(r => Math.Abs(r.OrphanDelta))
                .ThenByDescending(r => Math.Abs(r.FanOutMaxDelta))
                .ToList();
        }

        private static int JMeterSuitRank(string s)
        {
            if (string.IsNullOrEmpty(s)) return 0;
            if (s.StartsWith("最適")) return 3;
            if (s.StartsWith("良好")) return 2;
            return 1;
        }

        private static void CompareJMeterCandidates(AnalysisWorkbookSnapshot before, AnalysisWorkbookSnapshot after, AnalysisComparisonResult res)
        {
            if (before.JMeterCandidates.Count == 0 && after.JMeterCandidates.Count == 0) return;
            if (before.HasEmbeddedSnapshot != after.HasEmbeddedSnapshot && (before.JMeterCandidates.Count == 0 || after.JMeterCandidates.Count == 0)) return;

            Dictionary<string, SnapshotJMeterCandidate> bMap = new Dictionary<string, SnapshotJMeterCandidate>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, SnapshotJMeterCandidate> aMap = new Dictionary<string, SnapshotJMeterCandidate>(StringComparer.OrdinalIgnoreCase);
            foreach (SnapshotJMeterCandidate j in before.JMeterCandidates) bMap[j.TableName + "." + j.KeyColumn] = j;
            foreach (SnapshotJMeterCandidate j in after.JMeterCandidates) aMap[j.TableName + "." + j.KeyColumn] = j;

            HashSet<string> keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string k in aMap.Keys) keys.Add(k);
            foreach (string k in bMap.Keys) keys.Add(k);

            foreach (string k in keys)
            {
                SnapshotJMeterCandidate bj = null, aj = null;
                bMap.TryGetValue(k, out bj);
                aMap.TryGetValue(k, out aj);

                JMeterComparisonItem ji = new JMeterComparisonItem();
                if (bj == null && aj != null)
                {
                    ji.Status = "新規候補";
                    ji.TableName = aj.TableName; ji.KeyColumn = aj.KeyColumn; ji.DataType = aj.DataType;
                    ji.AfterUnique = aj.UniqueCount; ji.UniqueDelta = aj.UniqueCount;
                    ji.AfterSuitability = aj.Suitability; ji.RecommendedUsage = aj.RecommendedUsage;
                }
                else if (bj != null && aj == null)
                {
                    ji.Status = "消失";
                    ji.TableName = bj.TableName; ji.KeyColumn = bj.KeyColumn; ji.DataType = bj.DataType;
                    ji.BeforeUnique = bj.UniqueCount; ji.UniqueDelta = -bj.UniqueCount;
                    ji.BeforeSuitability = bj.Suitability; ji.AfterSuitability = "候補外";
                    ji.RecommendedUsage = "キー数減少またはテーブル消失により候補外";
                    ji.HasAlert = true;
                }
                else
                {
                    ji.TableName = aj.TableName; ji.KeyColumn = aj.KeyColumn; ji.DataType = aj.DataType;
                    ji.BeforeUnique = bj.UniqueCount; ji.AfterUnique = aj.UniqueCount; ji.UniqueDelta = aj.UniqueCount - bj.UniqueCount;
                    ji.BeforeSuitability = bj.Suitability; ji.AfterSuitability = aj.Suitability;
                    ji.RecommendedUsage = aj.RecommendedUsage;
                    int bRank = JMeterSuitRank(bj.Suitability);
                    int aRank = JMeterSuitRank(aj.Suitability);
                    if (aRank < bRank) { ji.Status = "悪化（要注意）"; ji.HasAlert = true; }
                    else if (aRank > bRank) { ji.Status = "改善"; }
                    else if (ji.UniqueDelta > 0) { ji.Status = "キー増加"; }
                    else if (ji.UniqueDelta < 0) { ji.Status = "キー減少"; ji.HasAlert = true; }
                    else { ji.Status = "変化なし"; }
                }
                res.JMeterComparisons.Add(ji);
            }

            res.JMeterComparisons = res.JMeterComparisons
                .OrderBy(j => j.HasAlert ? 0 : (j.Status == "変化なし" ? 2 : 1))
                .ThenByDescending(j => Math.Abs(j.UniqueDelta))
                .ThenBy(j => j.TableName, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static void CompareTimeSeries(AnalysisWorkbookSnapshot before, AnalysisWorkbookSnapshot after, AnalysisComparisonResult res)
        {
            // Compare monthly, recent daily, and hourly activity across matching datetime columns that actually changed
            foreach (SnapshotTable at in after.TableList)
            {
                SnapshotTable bt = null;
                before.Tables.TryGetValue(at.TableName, out bt);
                foreach (KeyValuePair<string, SnapshotDateTimeStats> kv in at.DateTimeStats)
                {
                    SnapshotDateTimeStats ad = kv.Value;
                    if (!ad.HasTimeSeriesDist) continue;
                    SnapshotDateTimeStats bd = null;
                    if (bt != null) bt.DateTimeStats.TryGetValue(kv.Key, out bd);
                    if (bd != null && !bd.HasTimeSeriesDist) continue;

                    string tc = at.TableName + "." + kv.Key;

                    // Monthly deltas
                    HashSet<string> months = new HashSet<string>(StringComparer.Ordinal);
                    if (bd != null) foreach (string m in bd.MonthlyCounts.Keys) months.Add(m);
                    foreach (string m in ad.MonthlyCounts.Keys) months.Add(m);
                    foreach (string m in months.OrderBy(x => x, StringComparer.Ordinal))
                    {
                        long bc = (bd != null && bd.MonthlyCounts.ContainsKey(m)) ? bd.MonthlyCounts[m] : 0L;
                        long ac = ad.MonthlyCounts.ContainsKey(m) ? ad.MonthlyCounts[m] : 0L;
                        if (bc != ac)
                        {
                            res.TimeSeriesDeltas.Add(MakeTsDelta("月別", tc, m, bc, ac));
                        }
                    }

                    // Recent Daily deltas
                    HashSet<string> days = new HashSet<string>(StringComparer.Ordinal);
                    if (bd != null) foreach (string d in bd.RecentDailyCounts.Keys) days.Add(d);
                    foreach (string d in ad.RecentDailyCounts.Keys) days.Add(d);
                    foreach (string d in days.OrderBy(x => x, StringComparer.Ordinal))
                    {
                        long bc = (bd != null && bd.RecentDailyCounts.ContainsKey(d)) ? bd.RecentDailyCounts[d] : 0L;
                        long ac = ad.RecentDailyCounts.ContainsKey(d) ? ad.RecentDailyCounts[d] : 0L;
                        if (bc != ac)
                        {
                            res.TimeSeriesDeltas.Add(MakeTsDelta("日別", tc, d, bc, ac));
                        }
                    }

                    // Hourly deltas (00:00..23:00)
                    if (ad.HasTimeComponent)
                    {
                        for (int h = 0; h < 24; h++)
                        {
                            long bc = (bd != null && bd.HourlyCounts != null && bd.HourlyCounts.Length == 24) ? bd.HourlyCounts[h] : 0L;
                            long ac = (ad.HourlyCounts != null && ad.HourlyCounts.Length == 24) ? ad.HourlyCounts[h] : 0L;
                            if (bc != ac)
                            {
                                res.TimeSeriesDeltas.Add(MakeTsDelta("時間帯別", tc, string.Format(CultureInfo.InvariantCulture, "{0:00}:00台", h), bc, ac));
                            }
                        }
                    }
                }
            }

            // Cap excessive time-series rows cleanly while preserving top deltas
            if (res.TimeSeriesDeltas.Count > 250)
            {
                res.TimeSeriesDeltas = res.TimeSeriesDeltas
                    .OrderByDescending(x => Math.Abs(x.DeltaCount))
                    .Take(250)
                    .OrderBy(x => x.Granularity, StringComparer.Ordinal)
                    .ThenBy(x => x.TableColumn, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(x => x.PeriodLabel, StringComparer.Ordinal)
                    .ToList();
            }
        }

        private static TimeSeriesDeltaItem MakeTsDelta(string gran, string tc, string period, long bc, long ac)
        {
            TimeSeriesDeltaItem ti = new TimeSeriesDeltaItem();
            ti.Granularity = gran;
            ti.TableColumn = tc;
            ti.PeriodLabel = period;
            ti.BeforeCount = bc;
            ti.AfterCount = ac;
            ti.DeltaCount = ac - bc;
            if (bc > 0)
            {
                ti.HasGrowthRate = true;
                ti.GrowthRate = (double)ti.DeltaCount / bc;
            }
            return ti;
        }

        private static void CompareServerMetrics(AnalysisWorkbookSnapshot before, AnalysisWorkbookSnapshot after, AnalysisComparisonResult res)
        {
            if (before.DbServer == null || !before.DbServer.HasStatus || after.DbServer == null || !after.DbServer.HasStatus) return;
            res.HasDbServerStatus = true;

            Dictionary<string, long> bs = before.DbServer.StatusCounters;
            Dictionary<string, long> as_ = after.DbServer.StatusCounters;

            long bUp = 0, aUp = 0;
            bool hasUp = bs.TryGetValue("Uptime", out bUp) && as_.TryGetValue("Uptime", out aUp);
            long upDelta = hasUp ? (aUp - bUp) : 0L;
            if (hasUp && upDelta < 0)
            {
                res.DbRestartDetected = true;
            }
            else if (hasUp && upDelta > 0)
            {
                res.ServerUptimeDeltaSeconds = upDelta;
            }
            double denomSec = (upDelta > 0) ? (double)upDelta : (res.HasElapsedTime && res.ElapsedSeconds > 0 ? res.ElapsedSeconds : 0.0);

            // Helper local logic via static method
            AddSrvCounter(res, bs, as_, "稼働・スループット", "サーバー稼働時間 (Uptime)", "Uptime", "秒", denomSec, false, res.DbRestartDetected ? "警告" : "通常",
                res.DbRestartDetected ? "テスト期間中にDB再起動/フェイルオーバー発生の可能性あり！" : "テスト前後間のDB稼働時間増分");
            AddSrvCounter(res, bs, as_, "稼働・スループット", "総クエリ数 (Questions)", "Questions", "回", denomSec, true, "通常", "テスト期間中の総実行クエリ数・平均QPS");
            AddSrvCounter(res, bs, as_, "SQL種別", "SELECT実行数 (Com_select)", "Com_select", "回", denomSec, true, "通常", "読取クエリ実行数 (開始前スキャン分含む)");
            AddSrvCounter(res, bs, as_, "SQL種別", "INSERT実行数 (Com_insert)", "Com_insert", "回", denomSec, true, "通常", "新規挿入クエリ数");
            AddSrvCounter(res, bs, as_, "SQL種別", "UPDATE実行数 (Com_update)", "Com_update", "回", denomSec, true, "通常", "既存更新クエリ数");
            AddSrvCounter(res, bs, as_, "SQL種別", "DELETE実行数 (Com_delete)", "Com_delete", "回", denomSec, true, "通常", "削除クエリ数");
            AddSrvCounter(res, bs, as_, "SQL種別", "COMMIT実行数 (Com_commit)", "Com_commit", "回", denomSec, true, "通常", "コミット回数");
            AddSrvCounter(res, bs, as_, "SQL種別", "ROLLBACK実行数 (Com_rollback)", "Com_rollback", "回", denomSec, true, "AUTO_WARN_POS", "ロールバック発生回数 (エラー/デッドロック/リトライ指標)");
            AddSrvCounter(res, bs, as_, "クエリ効率", "スロークエリ数 (Slow_queries)", "Slow_queries", "件", denomSec, true, "AUTO_ALERT_POS", "スロークエリ発生件数 (性能ボトルネック指標)");
            AddSrvCounter(res, bs, as_, "InnoDB行操作", "InnoDB読取行数 (Innodb_rows_read)", "Innodb_rows_read", "行", denomSec, true, "通常", "InnoDBエンジン層での総読取行数");
            AddSrvCounter(res, bs, as_, "InnoDB行操作", "InnoDB挿入行数 (Innodb_rows_inserted)", "Innodb_rows_inserted", "行", denomSec, true, "通常", "InnoDB層での総挿入行数");
            AddSrvCounter(res, bs, as_, "InnoDB行操作", "InnoDB更新行数 (Innodb_rows_updated)", "Innodb_rows_updated", "行", denomSec, true, "通常", "InnoDB層での総更新行数");
            AddSrvCounter(res, bs, as_, "InnoDB行操作", "InnoDB削除行数 (Innodb_rows_deleted)", "Innodb_rows_deleted", "行", denomSec, true, "通常", "InnoDB層での総削除行数");
            AddSrvCounter(res, bs, as_, "ロック競合", "行ロック待機回数 (Innodb_row_lock_waits)", "Innodb_row_lock_waits", "回", denomSec, true, "AUTO_ALERT_POS", "同時更新による行ロック競合の発生回数");
            AddSrvCounter(res, bs, as_, "ロック競合", "行ロック待機累計時間 (Innodb_row_lock_time)", "Innodb_row_lock_time", "ms", denomSec, false, "AUTO_WARN_POS", "行ロック待機の合計時間(ms)");
            AddSrvCounter(res, bs, as_, "Buffer Pool / I/O", "Buffer Pool論理読込 (read_requests)", "Innodb_buffer_pool_read_requests", "回", denomSec, true, "通常", "メモリキャッシュ論理読込回数");
            AddSrvCounter(res, bs, as_, "Buffer Pool / I/O", "Buffer Pool物理読込 (reads)", "Innodb_buffer_pool_reads", "回", denomSec, true, "AUTO_WARN_POS", "ディスク/ストレージ層への物理読込回数");
            AddSrvCounter(res, bs, as_, "Buffer Pool / I/O", "InnoDB Redoログ書込量 (os_log_written)", "Innodb_os_log_written", "B", denomSec, true, "通常", "書込トランザクションによるRedoログ発生量");
            AddSrvCounter(res, bs, as_, "クエリ効率", "ディスク一時テーブル (Created_tmp_disk_tables)", "Created_tmp_disk_tables", "回", denomSec, true, "AUTO_WARN_POS", "メモリ超過によるディスク一時テーブル作成回数");
            AddSrvCounter(res, bs, as_, "クエリ効率", "INDEX未使用結合 (Select_full_join)", "Select_full_join", "回", denomSec, true, "AUTO_ALERT_POS", "インデックスを使用しないフルテーブルJOIN回数");
            AddSrvCounter(res, bs, as_, "クエリ効率", "ソートマージ発生 (Sort_merge_passes)", "Sort_merge_passes", "回", denomSec, true, "AUTO_WARN_POS", "sort_buffer_size超過によるソートマージ回数");
            AddSrvCounter(res, bs, as_, "接続・ネットワーク", "最大同時接続数 (Max_used_connections)", "Max_used_connections", "接続", 0.0, false, "通常", "サーバー起動来の最大同時接続ピーク値");
            AddSrvCounter(res, bs, as_, "接続・ネットワーク", "異常切断クライアント (Aborted_clients)", "Aborted_clients", "回", denomSec, false, "AUTO_WARN_POS", "クライアント異常切断回数");
            AddSrvCounter(res, bs, as_, "接続・ネットワーク", "受信バイト数 (Bytes_received)", "Bytes_received", "B", denomSec, true, "通常", "DBサーバー受信トラフィック");
            AddSrvCounter(res, bs, as_, "接続・ネットワーク", "送信バイト数 (Bytes_sent)", "Bytes_sent", "B", denomSec, true, "通常", "DBサーバー送信トラフィック");

            // Derived metric: Buffer Pool Hit Ratio during test window
            long bReq, aReq, bPh, aPh;
            if (!res.DbRestartDetected &&
                bs.TryGetValue("Innodb_buffer_pool_read_requests", out bReq) && as_.TryGetValue("Innodb_buffer_pool_read_requests", out aReq) &&
                bs.TryGetValue("Innodb_buffer_pool_reads", out bPh) && as_.TryGetValue("Innodb_buffer_pool_reads", out aPh))
            {
                long dReq = aReq - bReq;
                long dPh = aPh - bPh;
                if (dReq > 0)
                {
                    double hitRate = Math.Max(0.0, Math.Min(1.0, 1.0 - ((double)dPh / (double)dReq)));
                    ServerMetricComparisonItem hr = new ServerMetricComparisonItem();
                    hr.Category = "Buffer Pool / I/O";
                    hr.MetricName = "テスト期間 Buffer Poolヒット率 (算出)";
                    hr.Unit = "%";
                    hr.DisplayBefore = "-";
                    hr.DisplayAfter = (hitRate * 100.0).ToString("0.00", CultureInfo.InvariantCulture) + "%";
                    hr.DisplayDelta = string.Format("物理読込 {0:#,##0} / 論理 {1:#,##0}", dPh, dReq);
                    hr.RatePerSecDisplay = "-";
                    hr.Severity = hitRate < 0.99 ? "警告" : (hitRate < 0.995 ? "注意" : "通常");
                    hr.Note = hitRate < 0.99 ? "ヒット率99%未満: ACU/Buffer Pool不足またはフルスキャン発生の可能性" : "良好なBuffer Poolキャッシュ効率";
                    res.ServerMetricComparisons.Add(hr);
                }
            }
        }

        private static void AddSrvCounter(AnalysisComparisonResult res, Dictionary<string, long> bs, Dictionary<string, long> as_,
            string cat, string name, string key, string unit, double denomSec, bool showRate, string sevMode, string note)
        {
            long bv, av;
            if (!bs.TryGetValue(key, out bv) || !as_.TryGetValue(key, out av)) return;
            long dv = av - bv;
            ServerMetricComparisonItem item = new ServerMetricComparisonItem();
            item.Category = cat;
            item.MetricName = name;
            item.Unit = unit;
            item.BeforeValue = bv;
            item.AfterValue = av;
            item.DeltaValue = dv;
            item.DisplayBefore = bv.ToString("#,##0", CultureInfo.InvariantCulture) + " " + unit;
            item.DisplayAfter = av.ToString("#,##0", CultureInfo.InvariantCulture) + " " + unit;
            item.DisplayDelta = FormatSignedInt(dv) + " " + unit;
            if (showRate && denomSec > 0.0 && dv >= 0 && !res.DbRestartDetected)
            {
                item.RatePerSecDisplay = (dv / denomSec).ToString("#,##0.00", CultureInfo.InvariantCulture) + " " + unit + "/秒";
            }
            if (sevMode == "AUTO_ALERT_POS") item.Severity = dv > 0 ? "警告" : "通常";
            else if (sevMode == "AUTO_WARN_POS") item.Severity = dv > 0 ? "注意" : "通常";
            else item.Severity = sevMode;
            item.Note = note;
            res.ServerMetricComparisons.Add(item);
        }

        private static void BuildComparisonInsights(AnalysisComparisonResult res)
        {
            if (string.Equals(res.BeforeFilePath, res.AfterFilePath, StringComparison.OrdinalIgnoreCase))
            {
                res.Insights.Add("【同一ファイル比較】テスト開始前・開始後に同一パスが指定されたため、全テーブルが「変更なし」と判定されています。");
            }
            if (res.ClockReversed)
            {
                res.Insights.Add(string.Format(
                    "【要注意：分析日時の逆転】開始前ファイル ({0}) より開始後ファイル ({1}) の方が古い日時です。比較対象の指定順序が逆になっていないか確認してください。",
                    res.BeforeAnalyzedAt, res.AfterAnalyzedAt));
            }
            if (res.DbRestartDetected)
            {
                res.Insights.Add("【要注意：DBサーバー再起動/フェイルオーバー検出】MySQLのUptimeが開始前より減少しています。テスト期間中にAuroraライター再起動またはフェイルオーバーが発生した可能性があります。");
            }

            string elapsedNote = res.HasElapsedTime
                ? string.Format(" / 計測間隔 {0} (平均増減 {1} 件/分, InnoDB {2} KB/分)",
                    res.ElapsedDisplay, FormatSignedDouble(res.TotalRowsPerMinute, "0.0"), FormatSignedDouble(res.TotalInnoDbKBPerMinute, "0.0"))
                : "";
            res.Insights.Add(string.Format(
                "【比較方式・照合精度】{0}{1}。全 {2} テーブル（変更あり {3} / 変更なし {4}）・全 {5} カラム（変更あり {6} / 変更なし {7}）を照合しました。",
                res.PrecisionLabel, elapsedNote,
                res.TotalUniqueTableCount, res.ChangedTableCount, res.UnchangedTableCount,
                res.TotalUniqueColumnCount, res.ChangedColumnTotal, res.UnchangedColumnTotal));

            string growthTxt = res.HasTotalRowGrowthRate
                ? string.Format(" ({0})", FormatSignedPercent(res.TotalRowGrowthRate))
                : "";
            string physTxt = res.HasDbPhysicalMeta
                ? string.Format("、実InnoDB Data {0} KB / Index {1} KB (断片化空き {2} KB)",
                    FormatSignedDouble(res.TotalDataLengthDeltaKB, "0.0"), FormatSignedDouble(res.TotalIndexLengthDeltaKB, "0.0"), FormatSignedDouble(res.TotalDataFreeDeltaKB, "0.0"))
                : "";
            string acuTxt = res.HasAcuComparison
                ? string.Format("、推定最低ACU {0:0.0}→{1:0.0} ACU (10倍時 {2:0.0}→{3:0.0} ACU)",
                    res.BeforeEstMinAcuCurrent, res.AfterEstMinAcuCurrent, res.BeforeEstMinAcu10x, res.AfterEstMinAcu10x)
                : "";
            res.Insights.Add(string.Format(
                "【全体規模・容量推移】比較対象 {0} テーブル（開始前 {1} → 開始後 {2}）。総レコード数 {3:#,##0} → {4:#,##0} 件（増減 {5} 件{6}）、推定InnoDB容量 {7:#,##0.0} → {8:#,##0.0} KB（増減 {9} KB）{10}{11}。",
                res.TotalUniqueTableCount, res.BeforeTableCount, res.AfterTableCount,
                res.BeforeTotalRows, res.AfterTotalRows, FormatSignedInt(res.TotalRowDelta), growthTxt,
                res.BeforeTotalInnoDbKB, res.AfterTotalInnoDbKB, FormatSignedDouble(res.TotalInnoDbDeltaKB, "0.0"),
                physTxt, acuTxt));

            res.Insights.Add(string.Format(
                "【テーブル変更内訳】追加あり {0} テーブル（新規テーブル {1} / レコード増加 {2} / 追加・更新併用 {3}）、更新あり {4} テーブル（データ更新 {5} / レコード減少 {6} / 追加・更新併用 {3}）、削除・消失 {7} テーブル、変更なし {8} テーブル。",
                res.AddedSheetCount, res.NewTableCount, res.RowAddedOnlyCount, res.AddedAndUpdatedCount,
                res.UpdatedSheetCount, res.UpdatedOnlyCount, res.RowDecreasedCount,
                res.DeletedTableCount, res.UnchangedTableCount));

            if (res.HasAnyRowChurn)
            {
                res.Insights.Add(string.Format(
                    "【行レベル変動内訳 ({0} / 対象 {1} テーブル)】挿入 +{2:#,##0} 行 / 既存更新 {3:#,##0} 行 / 削除 -{4:#,##0} 行 / 無変更 {5:#,##0} 行。純増減だけでは見えない同時INSERT+UPDATE+DELETEの実挙動を確認できます。",
                    res.RowChurnAllExact ? "PK完全一致計測" : "PK計測＋Bottom-K推定", res.RowChurnCoveredTables,
                    res.TotalInsertedRows, res.TotalUpdatedRows, res.TotalDeletedRows, res.TotalUnchangedRows));
            }

            if (res.AddedTables.Count > 0)
            {
                TableComparisonItem topAdd = res.AddedTables.OrderByDescending(t => t.RowDelta).First();
                string aiNote = topAdd.HasAutoIncrement && topAdd.AutoIncrementGap > 0
                    ? string.Format(" / AUTO_INCREMENTギャップ {0:#,##0}件 (ROLLBACK等)", topAdd.AutoIncrementGap)
                    : "";
                res.Insights.Add(string.Format(
                    "【最大レコード増加テーブル】`{0}` が最大増加（{1} 件: {2:#,##0} → {3:#,##0} 件, InnoDB増減 {4} KB{5}）。JMeter 性能試験の主要 INSERT 負荷先として、AUTO_INCREMENT / INDEX 肥大化および Buffer Pool 消費推移を確認してください。",
                    topAdd.TableName, FormatSignedInt(topAdd.RowDelta), topAdd.BeforeRows, topAdd.AfterRows, FormatSignedDouble(topAdd.InnoDbDeltaKB, "0.0"), aiNote));
            }

            List<TableComparisonItem> inPlaceUpdated = res.UpdatedTables.Where(t => t.PrimaryStatus == "更新（データ更新）" || t.PrimaryStatus == "追加・更新").ToList();
            if (inPlaceUpdated.Count > 0)
            {
                string names = string.Join(", ", inPlaceUpdated.Take(5).Select(t => "`" + t.TableName + "`").ToArray());
                res.Insights.Add(string.Format(
                    "【既存レコード更新検出】{0} テーブル（{1}{2}）で既存データの更新を検出しました。Aurora MySQL (InnoDB) では同時 UPDATE による行ロック競合・Undo ログ増加・デッドロック要因となり得るため、負荷試験中の Lock Wait 時間とデッドロック有無を確認してください。",
                    inPlaceUpdated.Count, names, inPlaceUpdated.Count > 5 ? " 他" : ""));
            }

            if (res.DeletedTableCount > 0 || res.RowDecreasedCount > 0)
            {
                res.Insights.Add(string.Format(
                    "【要注意：削除・減少検出】テーブル消失 {0} 件、レコード減少 {1} 件が検出されました。テストシナリオ上の意図的な DELETE / パージ処理か、データ欠損・取得漏れかを必ず確認してください。",
                    res.DeletedTableCount, res.RowDecreasedCount));
            }

            if (res.NewWarningCount > 0 || res.WorsenedWarningCount > 0 || res.ResolvedWarningCount > 0)
            {
                string prefix = (res.NewWarningCount > 0 || res.WorsenedWarningCount > 0) ? "【要注意：性能リスク・警告の変化】" : "【性能リスク・警告の推移】";
                res.Insights.Add(string.Format(
                    "{0}テスト後に警告が 新規発生 {1} 件 / 悪化 {2} 件 / 解消 {3} 件 / 継続 {4} 件 です。詳細はシート「08_性能・DB・時系列・品質変化」を確認してください。",
                    prefix, res.NewWarningCount, res.WorsenedWarningCount, res.ResolvedWarningCount, res.OngoingWarningCount));
            }

            if (res.OrphanRowDelta > 0)
            {
                res.Insights.Add(string.Format(
                    "【要注意：参照整合性(孤立FK)悪化】テスト前後でFK孤立子レコードが +{0:#,##0} 行増加しました（{1:#,##0} → {2:#,##0} 行）。トランザクション境界や削除順序の不具合がないか確認してください。",
                    res.OrphanRowDelta, res.BeforeTotalOrphanRows, res.AfterTotalOrphanRows));
            }

            if (res.DbSchemaDiffs.Count > 0)
            {
                res.Insights.Add(string.Format(
                    "【要注意：DBスキーマ/DDL・インデックス変化】テスト前後でDDL/INDEX/FK定義の差分を {0} 件検出しました。テスト中のマイグレーションや意図しないスキーマ差異がないか確認してください。",
                    res.DbSchemaDiffs.Count));
            }

            if (res.HasDbServerStatus)
            {
                ServerMetricComparisonItem q = res.ServerMetricComparisons.FirstOrDefault(x => x.MetricName.IndexOf("Questions", StringComparison.Ordinal) >= 0);
                ServerMetricComparisonItem slow = res.ServerMetricComparisons.FirstOrDefault(x => x.MetricName.IndexOf("Slow_queries", StringComparison.Ordinal) >= 0);
                ServerMetricComparisonItem lck = res.ServerMetricComparisons.FirstOrDefault(x => x.MetricName.IndexOf("Innodb_row_lock_waits", StringComparison.Ordinal) >= 0);
                ServerMetricComparisonItem rb = res.ServerMetricComparisons.FirstOrDefault(x => x.MetricName.IndexOf("Com_rollback", StringComparison.Ordinal) >= 0);
                List<string> sParts = new List<string>();
                if (q != null) sParts.Add("クエリ " + q.DisplayDelta + (q.RatePerSecDisplay != "-" ? (" (" + q.RatePerSecDisplay + ")") : ""));
                if (slow != null) sParts.Add("SlowQuery " + slow.DisplayDelta);
                if (lck != null) sParts.Add("行ロック待機 " + lck.DisplayDelta);
                if (rb != null) sParts.Add("ROLLBACK " + rb.DisplayDelta);
                if (sParts.Count > 0)
                {
                    bool srvWarn = (slow != null && slow.DeltaValue > 0) || (lck != null && lck.DeltaValue > 0) || (rb != null && rb.DeltaValue > 0);
                    res.Insights.Add((srvWarn ? "【要注意：Aurora/MySQLサーバー負荷差分】" : "【Aurora/MySQLサーバー負荷差分】") + string.Join(" / ", sParts.ToArray()) + "。");
                }
            }

            if (res.TimeSeriesDeltas.Count > 0)
            {
                TimeSeriesDeltaItem topHr = res.TimeSeriesDeltas.Where(x => x.Granularity == "時間帯別" && x.DeltaCount > 0).OrderByDescending(x => x.DeltaCount).FirstOrDefault();
                TimeSeriesDeltaItem topDay = res.TimeSeriesDeltas.Where(x => x.Granularity == "日別" && x.DeltaCount > 0).OrderByDescending(x => x.DeltaCount).FirstOrDefault();
                List<string> tsParts = new List<string>();
                if (topDay != null) tsParts.Add(string.Format("最大増加日 {0} ({1}: +{2:#,##0}件)", topDay.PeriodLabel, topDay.TableColumn, topDay.DeltaCount));
                if (topHr != null) tsParts.Add(string.Format("最大増加時間帯 {0} ({1}: +{2:#,##0}件)", topHr.PeriodLabel, topHr.TableColumn, topHr.DeltaCount));
                if (tsParts.Count > 0)
                {
                    res.Insights.Add("【時系列アクティビティ集中】" + string.Join(" / ", tsParts.ToArray()) + "。テスト実行時間帯のデータ偏りを確認できます。");
                }
            }

            if (res.UnchangedTableCount > 0)
            {
                string basis = res.UsedFullSnapshot ? "行・カラムの値マルチセットハッシュおよび全統計" : "分析ファイル内の統計プロファイル";
                res.Insights.Add(string.Format(
                    "【変更なしテーブル・カラム】{0} テーブル（および全テーブル中 {1} カラム）はテスト前後で{2}が完全一致しており、参照専用（マスタ系または未更新領域）と判定されました。テストシナリオで更新想定の対象が含まれていないかご確認ください。",
                    res.UnchangedTableCount, res.UnchangedColumnTotal, basis));
            }
        }

        private static ColumnDiffItem AddDiff(TableComparisonItem item, string colName, string cat, string metric, string beforeVal, string afterVal, string delta, string sev)
        {
            ColumnDiffItem d = new ColumnDiffItem();
            d.TableName = item.TableName;
            d.ColumnName = colName;
            d.Category = cat;
            d.MetricName = metric;
            d.BeforeValue = beforeVal ?? "";
            d.AfterValue = afterVal ?? "";
            d.DeltaSummary = delta ?? "";
            d.Severity = sev ?? "通常";
            item.ColumnDiffs.Add(d);
            return d;
        }

        private static bool NumbersNearlyEqual(double a, double b, double absTol)
        {
            double diff = Math.Abs(a - b);
            double relTol = 1e-9 * Math.Max(1.0, Math.Max(Math.Abs(a), Math.Abs(b)));
            return diff <= (absTol + relTol);
        }

        private static bool TryParseComparableDate(string s, out DateTime dt)
        {
            dt = DateTime.MinValue;
            if (string.IsNullOrWhiteSpace(s)) return false;
            string t = s.Trim();
            double oa;
            if (double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out oa) && oa > 20000.0 && oa < 80000.0)
            {
                try { dt = DateTime.FromOADate(oa); return true; }
                catch { return false; }
            }
            string[] fmts = new string[] {
                "yyyy-MM-dd HH:mm:ss", "yyyy/MM/dd HH:mm:ss",
                "yyyy-MM-dd HH:mm:ss.fff", "yyyy/MM/dd HH:mm:ss.fff",
                "yyyy-MM-ddTHH:mm:ss", "yyyy-MM-ddTHH:mm:ssZ",
                "yyyy-MM-dd", "yyyy/MM/dd"
            };
            return DateTime.TryParseExact(t, fmts, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt);
        }

        private static string NormalizeDateTop1(string raw, bool dateOnly)
        {
            if (string.IsNullOrWhiteSpace(raw)) return "";
            string t = raw.Trim();
            double oa;
            if (double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out oa) && oa > 20000.0 && oa < 80000.0)
            {
                try
                {
                    DateTime dt = DateTime.FromOADate(oa);
                    return dateOnly ? dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                }
                catch { }
            }
            return t;
        }

        private static bool Top1ValuesEqual(string a, string b)
        {
            string sa = a ?? "";
            string sb = b ?? "";
            if (string.Equals(sa, sb, StringComparison.Ordinal)) return true;
            // Handle Excel COM's 35-char Top1Value truncation vs OpenXML full value
            if (sa.Length == 35 && sa.EndsWith("...") && sb.Length > 35 && sb.StartsWith(sa.Substring(0, 32), StringComparison.Ordinal)) return true;
            if (sb.Length == 35 && sb.EndsWith("...") && sa.Length > 35 && sa.StartsWith(sb.Substring(0, 32), StringComparison.Ordinal)) return true;
            // Handle Excel COM OA serial date conversion vs OpenXML raw datetime string
            DateTime dtA, dtB;
            if (TryParseComparableDate(sa, out dtA) && TryParseComparableDate(sb, out dtB))
            {
                if (Math.Abs((dtA - dtB).TotalSeconds) < 1.0) return true;
            }
            return false;
        }

        private static string NormalizeDash(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            string t = s.Trim();
            return (t == "-" || t == "正常") ? "" : t;
        }

        private static string FormatSignedInt(long v)
        {
            if (v > 0) return "+" + v.ToString("#,##0", CultureInfo.InvariantCulture);
            if (v < 0) return v.ToString("#,##0", CultureInfo.InvariantCulture);
            return "0";
        }

        private static string FormatSignedDouble(double v, string fmt)
        {
            if (Math.Abs(v) < 1e-9) return (0.0).ToString(fmt, CultureInfo.InvariantCulture);
            if (v > 0) return "+" + v.ToString(fmt, CultureInfo.InvariantCulture);
            return v.ToString(fmt, CultureInfo.InvariantCulture);
        }

        private static string FormatSignedPercent(double v)
        {
            double pct = v * 100.0;
            if (Math.Abs(pct) < 0.005) return "0.0%";
            if (pct > 0) return "+" + pct.ToString("0.0", CultureInfo.InvariantCulture) + "%";
            return pct.ToString("0.0", CultureInfo.InvariantCulture) + "%";
        }

        // ---- OpenXML (.xlsx) Reader Helpers ----
        private static ZipArchiveEntry FindZipEntry(ZipArchive zip, string entryPath)
        {
            string norm = entryPath.Replace('\\', '/').TrimStart('/');
            foreach (ZipArchiveEntry e in zip.Entries)
            {
                if (string.Equals(e.FullName.Replace('\\', '/').TrimStart('/'), norm, StringComparison.OrdinalIgnoreCase))
                {
                    return e;
                }
            }
            return null;
        }

        private static string DecodeOoxmlText(string s)
        {
            if (string.IsNullOrEmpty(s)) return s ?? "";
            if (s.IndexOf("_x000D_", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                s = Regex.Replace(s, @"_x000[dD]_\r?\n", "\n");
                s = Regex.Replace(s, @"_x000[dD]_", "\n");
            }
            if (s.IndexOf('\r') >= 0)
            {
                s = s.Replace("\r\n", "\n").Replace("\r", "\n");
            }
            return s;
        }

        private static List<string> ReadSharedStrings(ZipArchive zip)
        {
            List<string> list = new List<string>();
            ZipArchiveEntry entry = FindZipEntry(zip, "xl/sharedStrings.xml");
            if (entry == null) return list;

            XmlDocument doc = new XmlDocument();
            doc.XmlResolver = null;
            using (Stream s = entry.Open())
            {
                doc.Load(s);
            }
            XmlNamespaceManager nsm = new XmlNamespaceManager(doc.NameTable);
            nsm.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            XmlNodeList sis = doc.SelectNodes("//x:sst/x:si", nsm);
            if (sis != null)
            {
                foreach (XmlNode si in sis)
                {
                    StringBuilder sb = new StringBuilder();
                    XmlNodeList ts = si.SelectNodes(".//x:t", nsm);
                    if (ts != null)
                    {
                        foreach (XmlNode t in ts) sb.Append(t.InnerText);
                    }
                    list.Add(DecodeOoxmlText(sb.ToString()));
                }
            }
            return list;
        }

        private static Dictionary<string, string> ReadSheetPartMap(ZipArchive zip)
        {
            Dictionary<string, string> map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            ZipArchiveEntry wbEntry = FindZipEntry(zip, "xl/workbook.xml");
            ZipArchiveEntry relEntry = FindZipEntry(zip, "xl/_rels/workbook.xml.rels");
            if (wbEntry == null || relEntry == null)
            {
                throw new InvalidDataException("有効なExcelブック構造 (xl/workbook.xml) が見つかりません。");
            }

            Dictionary<string, string> relMap = new Dictionary<string, string>(StringComparer.Ordinal);
            XmlDocument relDoc = new XmlDocument();
            relDoc.XmlResolver = null;
            using (Stream s = relEntry.Open()) { relDoc.Load(s); }
            foreach (XmlNode rn in relDoc.GetElementsByTagName("Relationship"))
            {
                if (rn.Attributes == null) continue;
                XmlAttribute idAttr = rn.Attributes["Id"];
                XmlAttribute tgtAttr = rn.Attributes["Target"];
                if (idAttr != null && tgtAttr != null)
                {
                    string tgt = tgtAttr.Value.Replace('\\', '/');
                    if (tgt.StartsWith("/")) tgt = tgt.TrimStart('/');
                    else
                    {
                        while (tgt.StartsWith("./")) tgt = tgt.Substring(2);
                        if (!tgt.StartsWith("xl/", StringComparison.OrdinalIgnoreCase)) tgt = "xl/" + tgt;
                    }
                    relMap[idAttr.Value] = tgt;
                }
            }

            XmlDocument wbDoc = new XmlDocument();
            wbDoc.XmlResolver = null;
            using (Stream s = wbEntry.Open()) { wbDoc.Load(s); }
            XmlNamespaceManager nsm = new XmlNamespaceManager(wbDoc.NameTable);
            nsm.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            nsm.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            XmlNodeList sheets = wbDoc.SelectNodes("//x:sheets/x:sheet", nsm);
            if (sheets != null)
            {
                foreach (XmlNode sh in sheets)
                {
                    if (sh.Attributes == null) continue;
                    XmlAttribute nameAttr = sh.Attributes["name"];
                    XmlAttribute ridAttr = sh.Attributes["id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"];
                    if (ridAttr == null) ridAttr = sh.Attributes["r:id"];
                    if (nameAttr != null && ridAttr != null && relMap.ContainsKey(ridAttr.Value))
                    {
                        map[nameAttr.Value] = relMap[ridAttr.Value];
                    }
                }
            }
            return map;
        }

        private static int ParseColIndexFromRef(string cellRef)
        {
            if (string.IsNullOrEmpty(cellRef)) return 0;
            int col = 0;
            for (int i = 0; i < cellRef.Length; i++)
            {
                char ch = cellRef[i];
                if (ch >= 'A' && ch <= 'Z') col = col * 26 + (ch - 'A' + 1);
                else if (ch >= 'a' && ch <= 'z') col = col * 26 + (ch - 'a' + 1);
                else break;
            }
            return col;
        }

        private static List<Dictionary<int, string>> ReadSheetRows(ZipArchive zip, string partPath, List<string> sharedStrings, int maxRows)
        {
            List<Dictionary<int, string>> rows = new List<Dictionary<int, string>>();
            ZipArchiveEntry entry = FindZipEntry(zip, partPath);
            if (entry == null) return rows;

            XmlReaderSettings settings = new XmlReaderSettings();
            settings.DtdProcessing = DtdProcessing.Prohibit;
            settings.IgnoreWhitespace = false;

            using (Stream s = entry.Open())
            using (XmlReader xr = XmlReader.Create(s, settings))
            {
                Dictionary<int, string> currentRow = null;
                int lastRowIndex = 0;
                int autoCol = 0;

                while (xr.Read())
                {
                    if (xr.NodeType == XmlNodeType.Element)
                    {
                        if (xr.LocalName == "row")
                        {
                            int rIdx = lastRowIndex + 1;
                            string rAttr = xr.GetAttribute("r");
                            if (!string.IsNullOrEmpty(rAttr)) int.TryParse(rAttr, out rIdx);
                            while (rows.Count < rIdx - 1)
                            {
                                rows.Add(new Dictionary<int, string>());
                            }
                            if (maxRows > 0 && rIdx > maxRows) break;
                            currentRow = new Dictionary<int, string>();
                            rows.Add(currentRow);
                            lastRowIndex = rIdx;
                            autoCol = 0;
                            if (xr.IsEmptyElement) currentRow = null;
                        }
                        else if (xr.LocalName == "c" && currentRow != null)
                        {
                            string cref = xr.GetAttribute("r");
                            int colIdx = ParseColIndexFromRef(cref);
                            if (colIdx <= 0) colIdx = autoCol + 1;
                            autoCol = colIdx;

                            string cellType = xr.GetAttribute("t") ?? "";
                            if (xr.IsEmptyElement) continue;

                            string val = ReadCellInnerValue(xr, cellType, sharedStrings);
                            currentRow[colIdx] = val;
                        }
                    }
                    else if (xr.NodeType == XmlNodeType.EndElement && xr.LocalName == "row")
                    {
                        currentRow = null;
                    }
                }
            }
            return rows;
        }

        private static string ReadCellInnerValue(XmlReader xr, string cellType, List<string> sharedStrings)
        {
            StringBuilder inlineSb = null;
            StringBuilder vSb = null;
            using (XmlReader sub = xr.ReadSubtree())
            {
                bool inV = false;
                bool inT = false;
                while (sub.Read())
                {
                    if (sub.NodeType == XmlNodeType.Element)
                    {
                        if (sub.LocalName == "v" && !sub.IsEmptyElement)
                        {
                            inV = true;
                            if (vSb == null) vSb = new StringBuilder();
                        }
                        else if (sub.LocalName == "t" && !sub.IsEmptyElement)
                        {
                            inT = true;
                            if (inlineSb == null) inlineSb = new StringBuilder();
                        }
                    }
                    else if (sub.NodeType == XmlNodeType.EndElement)
                    {
                        if (sub.LocalName == "v") inV = false;
                        else if (sub.LocalName == "t") inT = false;
                    }
                    else if (sub.NodeType == XmlNodeType.Text ||
                             sub.NodeType == XmlNodeType.CDATA ||
                             sub.NodeType == XmlNodeType.SignificantWhitespace ||
                             sub.NodeType == XmlNodeType.Whitespace)
                    {
                        if (inV) vSb.Append(sub.Value);
                        else if (inT) inlineSb.Append(sub.Value);
                    }
                }
            }

            string vText = vSb != null ? vSb.ToString() : null;

            if (cellType == "s")
            {
                int idx;
                if (vText != null && int.TryParse(vText, out idx) && idx >= 0 && idx < sharedStrings.Count)
                {
                    return sharedStrings[idx];
                }
                return "";
            }
            if (cellType == "inlineStr")
            {
                return inlineSb != null ? DecodeOoxmlText(inlineSb.ToString()) : "";
            }
            if (cellType == "b")
            {
                return vText == "1" ? "TRUE" : "FALSE";
            }
            if (vText != null) return DecodeOoxmlText(vText);
            if (inlineSb != null) return DecodeOoxmlText(inlineSb.ToString());
            return "";
        }

        private static string GetCell(Dictionary<int, string> row, int col)
        {
            if (row == null || col <= 0) return "";
            string v;
            return row.TryGetValue(col, out v) ? (v ?? "") : "";
        }

        private static long ParseLong(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return 0;
            double d;
            if (double.TryParse(s.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out d))
            {
                return (long)Math.Round(d);
            }
            return 0;
        }

        private static double ParseDouble(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return 0.0;
            double d;
            if (double.TryParse(s.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out d))
            {
                return d;
            }
            return 0.0;
        }

        private static string ParseDateCell(string raw, bool dateOnly)
        {
            if (string.IsNullOrWhiteSpace(raw)) return "-";
            string t = raw.Trim();
            if (t == "-") return "-";
            double oa;
            if (double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out oa) && oa > 20000 && oa < 80000)
            {
                DateTime dt = DateTime.FromOADate(oa);
                return dateOnly ? dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : dt.ToString("yyyy/MM/dd HH:mm:ss", CultureInfo.InvariantCulture);
            }
            return t;
        }

        private static void ParseSubtitleMeta(AnalysisWorkbookSnapshot snap, string subtitle)
        {
            snap.Subtitle = subtitle.Trim();
            Match mAt = Regex.Match(subtitle, @"分析日時:\s*([^|]+)");
            if (mAt.Success)
            {
                snap.AnalyzedAtText = mAt.Groups[1].Value.Trim();
                DateTime parsedDt;
                if (TryParseComparableDate(snap.AnalyzedAtText, out parsedDt)) snap.AnalyzedAt = parsedDt;
            }
            Match mDir = Regex.Match(subtitle, @"(?:対象フォルダ|接続先DB):\s*([^|]+)");
            if (mDir.Success) snap.InputFolder = mDir.Groups[1].Value.Trim();
            if (subtitle.IndexOf("DB直接分析", StringComparison.Ordinal) >= 0)
            {
                snap.IsDirectDb = true;
                snap.SourceMode = "DirectDB";
            }
            else if (string.IsNullOrEmpty(snap.SourceMode))
            {
                snap.SourceMode = "CSV";
            }
            Match mLbl = Regex.Match(subtitle, @"A5M2登録DB:\s*([^|]+)");
            if (mLbl.Success) snap.DbLabel = mLbl.Groups[1].Value.Trim();
        }

        private static int FindHeaderRow(List<Dictionary<int, string>> rows, string col1Expected, string col2Expected, out Dictionary<string, int> colMap)
        {
            colMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int r = 0; r < rows.Count && r < 12; r++)
            {
                string c1 = GetCell(rows[r], 1).Trim();
                string c2 = GetCell(rows[r], 2).Trim();
                if (string.Equals(c1, col1Expected, StringComparison.Ordinal) &&
                    string.Equals(c2, col2Expected, StringComparison.Ordinal))
                {
                    foreach (KeyValuePair<int, string> kv in rows[r])
                    {
                        string h = (kv.Value ?? "").Trim();
                        if (!string.IsNullOrEmpty(h) && !colMap.ContainsKey(h)) colMap[h] = kv.Key;
                    }
                    return r;
                }
            }
            return -1;
        }

        private static int GetColIdx(Dictionary<string, int> map, params string[] names)
        {
            foreach (string n in names)
            {
                int idx;
                if (map.TryGetValue(n, out idx)) return idx;
            }
            return -1;
        }

        private static bool IsRowEmptyOrPlaceholder(Dictionary<int, string> row)
        {
            if (row == null || row.Count == 0) return true;
            string c1 = GetCell(row, 1).Trim();
            string c2 = GetCell(row, 2).Trim();
            if (string.IsNullOrEmpty(c1) && string.IsNullOrEmpty(c2)) return true;
            if (c1 == "該当データなし" && string.IsNullOrEmpty(c2)) return true;
            return false;
        }

        private static bool TryParseFormattedNumber(string raw, out double value)
        {
            value = 0.0;
            if (string.IsNullOrWhiteSpace(raw)) return false;
            string s = raw.Trim();
            if (s == "-") return false;
            return double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out value) && !double.IsNaN(value) && !double.IsInfinity(value);
        }

        private static void ParseTopDistSummary(string summary, Dictionary<string, long> target)
        {
            if (string.IsNullOrWhiteSpace(summary)) return;
            string trimmed = summary.Trim();
            if (trimmed == "-" || trimmed.StartsWith("全件一意", StringComparison.Ordinal)) return;
            string[] parts = trimmed.Split(new string[] { " / " }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < parts.Length; i++)
            {
                Match m = Regex.Match(parts[i].Trim(), @"^(.+?):\s*([0-9,]+)\s*\([0-9\.]+%\)$");
                if (m.Success)
                {
                    string key = m.Groups[1].Value.Trim();
                    long cnt;
                    if (long.TryParse(m.Groups[2].Value.Replace(",", ""), NumberStyles.Integer, CultureInfo.InvariantCulture, out cnt))
                    {
                        target[key] = cnt;
                    }
                }
            }
        }

        private static void ParseTablesSheet(AnalysisWorkbookSnapshot snap, List<Dictionary<int, string>> rows)
        {
            Dictionary<string, int> map;
            int hdrIdx = FindHeaderRow(rows, "No", "テーブル名", out map);
            if (hdrIdx < 0)
            {
                throw new InvalidDataException("'02_テーブル一覧' シートのヘッダー行が見つかりません: " + snap.FilePath);
            }

            int iTbl = GetColIdx(map, "テーブル名");
            int iFile = GetColIdx(map, "CSVファイル名", "取得元 (DB.テーブル)", "取得元");
            int iEnc = GetColIdx(map, "文字コード");
            int iRole = GetColIdx(map, "推定役割");
            int iRows = GetColIdx(map, "レコード数");
            int iRowShare = GetColIdx(map, "件数構成比");
            int iCols = GetColIdx(map, "列数");
            int iSize = GetColIdx(map, "CSVサイズ(KB)", "CSV換算サイズ(KB)", "データサイズ(KB)");
            int iSizeShare = GetColIdx(map, "容量構成比");
            int iMinLen = GetColIdx(map, "最小行長(B)");
            int iAvgLen = GetColIdx(map, "平均行長(B)");
            int iMaxLen = GetColIdx(map, "最大行長(B)");
            int iInno = GetColIdx(map, "推定InnoDB(KB)");
            int iInnoPg = GetColIdx(map, "推定InnoDBページ数");
            int iNull = GetColIdx(map, "平均NULL率");
            int iPk = GetColIdx(map, "PK/UK候補");
            int iDt = GetColIdx(map, "日時カラム");
            int iJm = GetColIdx(map, "JMeter並列適性");
            int iWarnCnt = GetColIdx(map, "警告数");
            int iWarn = GetColIdx(map, "警告・注意点");

            for (int r = hdrIdx + 1; r < rows.Count; r++)
            {
                Dictionary<int, string> row = rows[r];
                if (IsRowEmptyOrPlaceholder(row)) break;
                string c1 = GetCell(row, 1).Trim();
                string tblName = GetCell(row, iTbl).Trim();
                if (string.IsNullOrEmpty(tblName) || c1.StartsWith("合計", StringComparison.Ordinal) || tblName == "合計" || tblName.StartsWith("合計", StringComparison.Ordinal)) break;

                SnapshotTable st = new SnapshotTable();
                st.TableName = tblName;
                st.FileName = GetCell(row, iFile).Trim();
                st.EncodingName = GetCell(row, iEnc).Trim();
                st.InferredRole = GetCell(row, iRole).Trim();
                st.RowCount = ParseLong(GetCell(row, iRows));
                st.RowShare = ParseDouble(GetCell(row, iRowShare));
                st.ColumnCount = (int)ParseLong(GetCell(row, iCols));
                st.FileSizeKB = ParseDouble(GetCell(row, iSize));
                st.SizeShare = ParseDouble(GetCell(row, iSizeShare));
                if (iMinLen > 0) st.MinRowBytes = ParseLong(GetCell(row, iMinLen));
                st.AvgRowBytes = ParseDouble(GetCell(row, iAvgLen));
                st.MaxRowBytes = ParseLong(GetCell(row, iMaxLen));
                st.EstInnoDbKB = ParseDouble(GetCell(row, iInno));
                if (iInnoPg > 0) st.EstInnoDbPages = ParseLong(GetCell(row, iInnoPg));
                st.AvgNullRate = ParseDouble(GetCell(row, iNull));
                st.CandidatePks = NormalizeDash(GetCell(row, iPk));
                st.DateColumns = NormalizeDash(GetCell(row, iDt));
                if (iJm > 0) st.JMeterReadiness = GetCell(row, iJm).Trim();
                if (iWarnCnt > 0) st.WarningCount = (int)ParseLong(GetCell(row, iWarnCnt));
                string warnRaw = GetCell(row, iWarn).Trim();
                st.WarningSummary = (warnRaw == "正常" || warnRaw == "-") ? "" : warnRaw;

                if (!snap.Tables.ContainsKey(st.TableName))
                {
                    snap.Tables[st.TableName] = st;
                    snap.TableList.Add(st);
                }
            }
        }

        private static void ParseColumnsSheet(AnalysisWorkbookSnapshot snap, List<Dictionary<int, string>> rows)
        {
            Dictionary<string, int> map;
            int hdrIdx = FindHeaderRow(rows, "テーブル名", "カラム名", out map);
            if (hdrIdx < 0) return;

            int iTbl = GetColIdx(map, "テーブル名");
            int iCol = GetColIdx(map, "カラム名");
            int iOrd = GetColIdx(map, "列順");
            int iType = GetColIdx(map, "推定データ型");
            int iRole = GetColIdx(map, "キー・役割");
            int iTot = GetColIdx(map, "総行数");
            int iNn = GetColIdx(map, "有効件数");
            int iNull = GetColIdx(map, "NULL件数");
            int iNullR = GetColIdx(map, "NULL率");
            int iDist = GetColIdx(map, "ユニーク数");
            int iCard = GetColIdx(map, "ユニーク率");
            int iSel = GetColIdx(map, "INDEX選択性評価");
            int iMinB = GetColIdx(map, "最小長(B)");
            int iAvgB = GetColIdx(map, "平均長(B)");
            int iP95B = GetColIdx(map, "P95長(B)");
            int iMaxB = GetColIdx(map, "最大長(B)");
            int iTop1 = GetColIdx(map, "最頻値(Top1)");
            int iTop1C = GetColIdx(map, "Top1件数");
            int iTop1S = GetColIdx(map, "Top1占有率");
            int iTop3S = GetColIdx(map, "Top3占有率");
            int iTop5S = GetColIdx(map, "Top5占有率");
            int iGini = GetColIdx(map, "ジニ係数");
            int iEnt = GetColIdx(map, "正規化エントロピー");
            int iSkew = GetColIdx(map, "偏りリスク");
            int iMinVal = GetColIdx(map, "最小値 / 最古日時");
            int iMaxVal = GetColIdx(map, "最大値 / 最新日時");
            int iTopDist = GetColIdx(map, "上位値分布 (Top5)");
            int iDetail = GetColIdx(map, "型別詳細統計 (数値・日時・文字特性)");
            int iAlert = GetColIdx(map, "警告・注意点");

            for (int r = hdrIdx + 1; r < rows.Count; r++)
            {
                Dictionary<int, string> row = rows[r];
                if (IsRowEmptyOrPlaceholder(row)) break;
                string tblName = GetCell(row, iTbl).Trim();
                string colName = GetCell(row, iCol).Trim();
                if (string.IsNullOrEmpty(tblName) || string.IsNullOrEmpty(colName)) break;

                SnapshotTable st;
                if (!snap.Tables.TryGetValue(tblName, out st)) continue;

                SnapshotColumn sc = new SnapshotColumn();
                sc.TableName = tblName;
                sc.ColumnName = colName;
                sc.ColumnIndex = iOrd > 0 ? Math.Max(0, (int)ParseLong(GetCell(row, iOrd)) - 1) : st.ColumnList.Count;
                sc.InferredType = GetCell(row, iType).Trim();
                if (iRole > 0) sc.KeyRole = GetCell(row, iRole).Trim();
                sc.TotalRows = ParseLong(GetCell(row, iTot));
                sc.NullCount = ParseLong(GetCell(row, iNull));
                sc.NonNullCount = iNn > 0 ? ParseLong(GetCell(row, iNn)) : Math.Max(0L, sc.TotalRows - sc.NullCount);
                sc.NullRate = ParseDouble(GetCell(row, iNullR));
                sc.DistinctCount = ParseLong(GetCell(row, iDist));
                sc.CardinalityRate = ParseDouble(GetCell(row, iCard));
                sc.SelectivityRating = GetCell(row, iSel).Trim();
                if (iMinB > 0) sc.MinByteLen = ParseLong(GetCell(row, iMinB));
                sc.AvgByteLen = ParseDouble(GetCell(row, iAvgB));
                sc.P95ByteLen = ParseLong(GetCell(row, iP95B));
                sc.MaxByteLen = ParseLong(GetCell(row, iMaxB));
                string rawTop1 = GetCell(row, iTop1);
                if (sc.InferredType.IndexOf("日時", StringComparison.Ordinal) >= 0)
                {
                    rawTop1 = NormalizeDateTop1(rawTop1, false);
                }
                else if (sc.InferredType.IndexOf("日付", StringComparison.Ordinal) >= 0)
                {
                    rawTop1 = NormalizeDateTop1(rawTop1, true);
                }
                sc.Top1Value = rawTop1;
                if (iTop1C > 0) sc.Top1Count = ParseLong(GetCell(row, iTop1C));
                sc.Top1Share = ParseDouble(GetCell(row, iTop1S));
                sc.Top3Share = ParseDouble(GetCell(row, iTop3S));
                if (iTop5S > 0) sc.Top5Share = ParseDouble(GetCell(row, iTop5S));
                sc.GiniCoefficient = ParseDouble(GetCell(row, iGini));
                if (iEnt > 0) { sc.NormalizedEntropy = ParseDouble(GetCell(row, iEnt)); sc.HasNormalizedEntropy = true; }
                sc.SkewRisk = GetCell(row, iSkew).Trim();
                if (iMinVal > 0)
                {
                    string rvMin = GetCell(row, iMinVal).Trim();
                    if (sc.InferredType.IndexOf("日時", StringComparison.Ordinal) >= 0) rvMin = ParseDateCell(rvMin, false);
                    else if (sc.InferredType.IndexOf("日付", StringComparison.Ordinal) >= 0) rvMin = ParseDateCell(rvMin, true);
                    sc.MinValueText = (rvMin == "-") ? "" : rvMin;
                }
                if (iMaxVal > 0)
                {
                    string rvMax = GetCell(row, iMaxVal).Trim();
                    if (sc.InferredType.IndexOf("日時", StringComparison.Ordinal) >= 0) rvMax = ParseDateCell(rvMax, false);
                    else if (sc.InferredType.IndexOf("日付", StringComparison.Ordinal) >= 0) rvMax = ParseDateCell(rvMax, true);
                    sc.MaxValueText = (rvMax == "-") ? "" : rvMax;
                }
                if (iTopDist > 0)
                {
                    ParseTopDistSummary(GetCell(row, iTopDist), sc.TopDistCounts);
                }
                if (iDetail > 0)
                {
                    sc.DetailStatsSummary = GetCell(row, iDetail).Trim();
                }
                string aRaw = GetCell(row, iAlert).Trim();
                sc.AlertNote = (aRaw == "正常" || aRaw == "-") ? "" : aRaw;

                st.Columns[sc.ColumnName] = sc;
                st.ColumnList.Add(sc);
            }
        }

        private static void ParseDistributionsSheet(AnalysisWorkbookSnapshot snap, List<Dictionary<int, string>> rows)
        {
            Dictionary<string, int> map;
            int hdrIdx = FindHeaderRow(rows, "テーブル名", "カラム名", out map);
            if (hdrIdx < 0) return;

            int iTbl = GetColIdx(map, "テーブル名");
            int iCol = GetColIdx(map, "カラム名");
            int iVal = GetColIdx(map, "カテゴリ値");
            int iCnt = GetColIdx(map, "件数");
            int iShare = GetColIdx(map, "割合(%)");

            for (int r = hdrIdx + 1; r < rows.Count; r++)
            {
                Dictionary<int, string> row = rows[r];
                if (IsRowEmptyOrPlaceholder(row)) break;
                string tblName = GetCell(row, iTbl).Trim();
                string colName = GetCell(row, iCol).Trim();
                if (string.IsNullOrEmpty(tblName) || string.IsNullOrEmpty(colName)) break;

                SnapshotTable st;
                if (!snap.Tables.TryGetValue(tblName, out st)) continue;

                Dictionary<string, SnapshotCategoryItem> colDist;
                if (!st.Distributions.TryGetValue(colName, out colDist))
                {
                    colDist = new Dictionary<string, SnapshotCategoryItem>(StringComparer.Ordinal);
                    st.Distributions[colName] = colDist;
                }

                SnapshotCategoryItem ci = new SnapshotCategoryItem();
                ci.Value = GetCell(row, iVal);
                ci.Count = ParseLong(GetCell(row, iCnt));
                ci.Share = ParseDouble(GetCell(row, iShare));
                colDist[ci.Value] = ci;
            }
        }

        private static void ParseNumericSheet(AnalysisWorkbookSnapshot snap, List<Dictionary<int, string>> rows)
        {
            Dictionary<string, int> map;
            int hdrIdx = FindHeaderRow(rows, "テーブル名", "カラム名", out map);
            if (hdrIdx < 0) return;

            int iTbl = GetColIdx(map, "テーブル名");
            int iCol = GetColIdx(map, "カラム名");
            int iCnt = GetColIdx(map, "有効件数");
            int iMin = GetColIdx(map, "最小値");
            int iP25 = GetColIdx(map, "P25(Q1)", "P25");
            int iMed = GetColIdx(map, "中央値(P50)");
            int iMean = GetColIdx(map, "平均値");
            int iP75 = GetColIdx(map, "P75(Q3)", "P75");
            int iP90 = GetColIdx(map, "P90");
            int iP95 = GetColIdx(map, "P95");
            int iP99 = GetColIdx(map, "P99");
            int iMax = GetColIdx(map, "最大値");
            int iStd = GetColIdx(map, "標準偏差");
            int iZero = GetColIdx(map, "ゼロ率");
            int iNeg = GetColIdx(map, "負値率");
            int iOut = GetColIdx(map, "外れ値率");

            for (int r = hdrIdx + 1; r < rows.Count; r++)
            {
                Dictionary<int, string> row = rows[r];
                if (IsRowEmptyOrPlaceholder(row)) break;
                string tblName = GetCell(row, iTbl).Trim();
                string colName = GetCell(row, iCol).Trim();
                if (string.IsNullOrEmpty(tblName) || string.IsNullOrEmpty(colName)) break;

                SnapshotTable st;
                if (!snap.Tables.TryGetValue(tblName, out st)) continue;

                SnapshotNumericStats ns = new SnapshotNumericStats();
                ns.ColumnName = colName;
                ns.NonNullCount = ParseLong(GetCell(row, iCnt));
                ns.Min = ParseDouble(GetCell(row, iMin));
                ns.P25 = ParseDouble(GetCell(row, iP25));
                ns.Median = ParseDouble(GetCell(row, iMed));
                ns.Mean = ParseDouble(GetCell(row, iMean));
                ns.P75 = ParseDouble(GetCell(row, iP75));
                if (iP90 > 0) { ns.HasP90 = true; ns.P90 = ParseDouble(GetCell(row, iP90)); }
                ns.P95 = ParseDouble(GetCell(row, iP95));
                ns.P99 = ParseDouble(GetCell(row, iP99));
                ns.Max = ParseDouble(GetCell(row, iMax));
                ns.StdDev = ParseDouble(GetCell(row, iStd));
                if (iZero > 0) { ns.HasZeroRate = true; ns.ZeroRate = ParseDouble(GetCell(row, iZero)); }
                if (iNeg > 0) { ns.HasNegativeRate = true; ns.NegativeRate = ParseDouble(GetCell(row, iNeg)); }
                ns.OutlierRate = ParseDouble(GetCell(row, iOut));

                st.NumericStats[colName] = ns;
            }
        }

        private static void ParseDateTimeSheet(AnalysisWorkbookSnapshot snap, List<Dictionary<int, string>> rows)
        {
            Dictionary<string, int> map;
            int hdrIdx = FindHeaderRow(rows, "テーブル名", "日時カラム", out map);
            if (hdrIdx < 0) return;

            int iTbl = GetColIdx(map, "テーブル名");
            int iCol = GetColIdx(map, "日時カラム");
            int iCnt = GetColIdx(map, "有効件数");
            int iMin = GetColIdx(map, "最古日時");
            int iMax = GetColIdx(map, "最新日時");
            int iSpan = GetColIdx(map, "期間(日)");
            int iAvg = GetColIdx(map, "日平均件数");
            int iPkD = GetColIdx(map, "ピーク日");
            int iPkC = GetColIdx(map, "ピーク日件数");
            int iPkR = GetColIdx(map, "ピーク倍率");

            for (int r = hdrIdx + 1; r < rows.Count; r++)
            {
                Dictionary<int, string> row = rows[r];
                if (IsRowEmptyOrPlaceholder(row)) break;
                string tblName = GetCell(row, iTbl).Trim();
                string colName = GetCell(row, iCol).Trim();
                if (string.IsNullOrEmpty(tblName) || string.IsNullOrEmpty(colName)) break;

                SnapshotTable st;
                if (!snap.Tables.TryGetValue(tblName, out st)) continue;

                SnapshotDateTimeStats ds = new SnapshotDateTimeStats();
                ds.ColumnName = colName;
                ds.NonNullCount = ParseLong(GetCell(row, iCnt));
                ds.MinDate = ParseDateCell(GetCell(row, iMin), false);
                ds.MaxDate = ParseDateCell(GetCell(row, iMax), false);
                ds.DateSpanDays = ParseDouble(GetCell(row, iSpan));
                ds.AvgPerDay = ParseDouble(GetCell(row, iAvg));
                ds.PeakDay = ParseDateCell(GetCell(row, iPkD), true);
                ds.PeakDayCount = ParseLong(GetCell(row, iPkC));
                ds.PeakBurstRatio = ParseDouble(GetCell(row, iPkR));

                st.DateTimeStats[colName] = ds;
            }
        }

        private static void ParseRelationshipsSheet(AnalysisWorkbookSnapshot snap, List<Dictionary<int, string>> rows)
        {
            Dictionary<string, int> map;
            int hdrIdx = FindHeaderRow(rows, "親テーブル.PK", "子テーブル.FK", out map);
            if (hdrIdx < 0) return;

            int iP = GetColIdx(map, "親テーブル.PK");
            int iC = GetColIdx(map, "子テーブル.FK");
            int iPd = GetColIdx(map, "親キー数");
            int iCn = GetColIdx(map, "子有効行数");
            int iMr = GetColIdx(map, "参照一致率");
            int iOr = GetColIdx(map, "孤立行数");
            int iFmin = GetColIdx(map, "最小(子/親)");
            int iFavg = GetColIdx(map, "平均(子/親)");
            int iFp50 = GetColIdx(map, "P50(子/親)");
            int iFp95 = GetColIdx(map, "P95(子/親)");
            int iFp99 = GetColIdx(map, "P99(子/親)");
            int iFmax = GetColIdx(map, "最大(子/親)");
            int iRisk = GetColIdx(map, "リスク評価");
            int iNote = GetColIdx(map, "所見・スーパーユーザー情報");

            for (int r = hdrIdx + 1; r < rows.Count; r++)
            {
                Dictionary<int, string> row = rows[r];
                if (IsRowEmptyOrPlaceholder(row)) break;
                string pStr = GetCell(row, iP).Trim();
                string cStr = GetCell(row, iC).Trim();
                if (string.IsNullOrEmpty(pStr) || string.IsNullOrEmpty(cStr)) break;

                string[] pp = pStr.Split('.');
                string[] cp = cStr.Split('.');
                if (pp.Length < 2 || cp.Length < 2) continue;
                SnapshotRelationship sr = new SnapshotRelationship();
                sr.ParentTable = pp[0]; sr.ParentColumn = pp[1];
                sr.ChildTable = cp[0]; sr.ChildColumn = cp[1];
                sr.ParentDistinctCount = ParseLong(GetCell(row, iPd));
                sr.ChildNonNullRows = ParseLong(GetCell(row, iCn));
                sr.MatchRate = ParseDouble(GetCell(row, iMr));
                sr.OrphanRows = ParseLong(GetCell(row, iOr));
                sr.FanOutMin = ParseLong(GetCell(row, iFmin));
                sr.FanOutAvg = ParseDouble(GetCell(row, iFavg));
                sr.FanOutP50 = ParseDouble(GetCell(row, iFp50));
                sr.FanOutP95 = ParseDouble(GetCell(row, iFp95));
                sr.FanOutP99 = ParseDouble(GetCell(row, iFp99));
                sr.FanOutMax = ParseLong(GetCell(row, iFmax));
                sr.RiskLevel = GetCell(row, iRisk).Trim();
                sr.Note = GetCell(row, iNote).Trim();
                snap.Relationships.Add(sr);
            }
        }

        private static void ParseWarningsSheet(AnalysisWorkbookSnapshot snap, List<Dictionary<int, string>> rows)
        {
            Dictionary<string, int> map;
            int hdrIdx = FindHeaderRow(rows, "重要度", "カテゴリ", out map);
            if (hdrIdx < 0) return;

            int iSev = GetColIdx(map, "重要度");
            int iCat = GetColIdx(map, "カテゴリ");
            int iTbl = GetColIdx(map, "テーブル名");
            int iCol = GetColIdx(map, "カラム名");
            int iMet = GetColIdx(map, "検出値・指標");
            int iImp = GetColIdx(map, "Aurora / 性能試験への影響");
            int iRec = GetColIdx(map, "推奨アクション");

            for (int r = hdrIdx + 1; r < rows.Count; r++)
            {
                Dictionary<int, string> row = rows[r];
                if (IsRowEmptyOrPlaceholder(row)) break;
                string sev = GetCell(row, iSev).Trim();
                string cat = GetCell(row, iCat).Trim();
                if (string.IsNullOrEmpty(sev) || string.IsNullOrEmpty(cat)) break;

                SnapshotWarning sw = new SnapshotWarning();
                sw.Severity = sev;
                sw.Category = cat;
                sw.TableName = GetCell(row, iTbl).Trim();
                sw.ColumnName = GetCell(row, iCol).Trim();
                sw.MetricValue = GetCell(row, iMet).Trim();
                sw.PerfImpact = GetCell(row, iImp).Trim();
                sw.Recommendation = GetCell(row, iRec).Trim();
                snap.Warnings.Add(sw);
            }
        }
    }

    // =========================================================================
    //  Background Worker Launcher (Pure C# worker threads, UI thread pumps DoEvents)
    // =========================================================================
    public static class BackgroundTaskLauncher
    {
        public static BackgroundTaskHandle<bool> StartConnect(
            MySql84Client client,
            string connectHost,
            int connectPort,
            string tlsTargetHost,
            string user,
            string password,
            string database,
            string sslMode,
            string caCertPath,
            int timeoutSeconds,
            OperationCancelState cancelState)
        {
            return new BackgroundTaskHandle<bool>(delegate(Action<string> emit)
            {
                client.Connect(connectHost, connectPort, tlsTargetHost, user, password, database, sslMode, caCertPath, timeoutSeconds, cancelState);
                return true;
            });
        }

        public static BackgroundTaskHandle<MySqlQueryResult> StartExecuteQuery(
            MySql84Client client,
            string sql,
            OperationCancelState cancelState)
        {
            return new BackgroundTaskHandle<MySqlQueryResult>(delegate(Action<string> emit)
            {
                if (cancelState != null) cancelState.ThrowIfCancelled();
                return client.ExecuteQuery(sql);
            });
        }

        public static BackgroundTaskHandle<long> StartExportQueryToCsv(
            MySql84Client client,
            string tableName,
            string sql,
            string csvFilePath,
            int maxRows,
            bool writeUtf8Bom,
            OperationCancelState cancelState)
        {
            return StartExportQueryToCsv(client, tableName, sql, csvFilePath, maxRows, writeUtf8Bom, cancelState, null);
        }

        public static BackgroundTaskHandle<long> StartExportQueryToCsv(MySql84Client client, string tableName,
            string sql, string csvFilePath, int maxRows, bool writeUtf8Bom, OperationCancelState cancelState,
            CsvExportAnalysis exportAnalysis)
        {
            return new BackgroundTaskHandle<long>(delegate(Action<string> emit, Action<long> reportRows)
            {
                return client.ExportQueryToCsv(
                    sql,
                    csvFilePath,
                    maxRows,
                    writeUtf8Bom,
                    cancelState,
                    reportRows, exportAnalysis);
            });
        }

        public static BackgroundTaskHandle<TableAnalysis> StartAnalyzeMySqlTable(
            DirectAnalysisSession session,
            MySql84Client client,
            string tableName,
            string databaseName,
            string sql,
            int maxRows,
            OperationCancelState cancelState)
        {
            return new BackgroundTaskHandle<TableAnalysis>(delegate(Action<string> emit)
            {
                return session.AnalyzeMySqlTable(
                    client,
                    tableName,
                    databaseName,
                    sql,
                    maxRows,
                    cancelState,
                    delegate(long rows)
                    {
                        emit(string.Format("    経過: {0} ({1:#,##0} 件読込中...)", tableName, rows));
                    });
            });
        }

        public static BackgroundTaskHandle<FolderAnalysisResult> StartCompleteDirectAnalysis(DirectAnalysisSession session, OperationCancelState cancelState)
        {
            return new BackgroundTaskHandle<FolderAnalysisResult>(emit => session.Complete(cancelState));
        }

        public static BackgroundTaskHandle<FolderAnalysisResult> StartAnalyzeFolder(
            string folderPath,
            OperationCancelState cancelState)
        {
            return StartAnalyzeFolder(folderPath, cancelState, null, null);
        }

        public static BackgroundTaskHandle<FolderAnalysisResult> StartAnalyzeFolder(
            string folderPath, OperationCancelState cancelState, string[] selectedFiles, CsvExportAnalysis exportAnalysis)
        {
            return new BackgroundTaskHandle<FolderAnalysisResult>(delegate(Action<string> emit)
            {
                return CsvAnalysisEngine.AnalyzeFiles(folderPath, selectedFiles, exportAnalysis, cancelState, emit);
            });
        }

        public static BackgroundTaskHandle<AnalysisComparisonResult> StartCompareWorkbooks(
            string beforePath,
            string afterPath,
            OperationCancelState cancelState)
        {
            return new BackgroundTaskHandle<AnalysisComparisonResult>(delegate(Action<string> emit)
            {
                return AnalysisComparisonEngine.CompareAnalysisWorkbooks(beforePath, afterPath, cancelState);
            });
        }
    }
}
'@

if (-not ([System.Management.Automation.PSTypeName]'AuroraAnalyzer.MySql84Client').Type) {
    Add-Type -TypeDefinition $script:CSharpCoreSource -ReferencedAssemblies @('System', 'System.Core', 'System.Data', 'System.Xml', 'System.IO.Compression', 'System.IO.Compression.FileSystem')
}

# ==============================================================================
#  Module 3: A5:SQL Mk-2 Registered Database Discovery & OpenSSH Tunnel Manager
# ==============================================================================
function Parse-A5ConnectionString {
    param(
        [Parameter(Mandatory = $true)][string]$DbName,
        [Parameter(Mandatory = $true)][string]$RawConnectionString,
        [hashtable]$ExtraParams = @{},
        [string]$SourceRegistryPath = '',
        [string]$SourceFile = ''
    )

    if ($SourceFile) {
        # A list entry can contain quoted values with embedded semicolons.
        $settings = [System.Data.Common.DbConnectionStringBuilder]::new()
        $settings.set_ConnectionString($RawConnectionString)
        $kv = [hashtable]$settings
    }
    else {
        $kv = @{}
        foreach ($pair in ($RawConnectionString -split ';')) {
            if ([string]::IsNullOrWhiteSpace($pair)) { continue }
            $eqIdx = $pair.IndexOf('=')
            if ($eqIdx -gt 0) {
                $k = $pair.Substring(0, $eqIdx).Trim()
                $v = $pair.Substring($eqIdx + 1).Trim()
                if (-not [string]::IsNullOrEmpty($k)) {
                    $kv[$k.ToLowerInvariant()] = $v
                }
            }
        }
    }

    foreach ($ek in $ExtraParams.Keys) {
        $lk = $ek.ToLowerInvariant()
        if (-not $kv.ContainsKey($lk) -or [string]::IsNullOrWhiteSpace($kv[$lk])) {
            $kv[$lk] = [string]$ExtraParams[$ek]
        }
    }

    function Get-FirstKey {
        param([hashtable]$Map, [string[]]$Keys, [string]$Default = '')
        foreach ($k in $Keys) {
            $lk = $k.ToLowerInvariant()
            if ($Map.ContainsKey($lk) -and -not [string]::IsNullOrWhiteSpace($Map[$lk])) {
                return [string]$Map[$lk]
            }
        }
        return $Default
    }

    $provider = Get-FirstKey $kv @('ProviderName', 'Provider', '__ProviderName') 'MySQL'
    $server   = Get-FirstKey $kv @('ServerName', 'Server', 'Host', 'HostName', 'DataSource', 'Data Source')
    $portStr  = Get-FirstKey $kv @('Port', 'ServerPort') '3306'
    $port     = 3306
    [void][int]::TryParse($portStr, [ref]$port)
    $database = Get-FirstKey $kv @('Database', 'DatabaseName', 'Initial Catalog', 'Schema', 'InitialSchemaName')
    $user     = Get-FirstKey $kv @('UserName', 'User', 'User ID', 'UID', 'User_Name')

    # The export omits credentials. A mask or password in a DB list is never a login password.
    $rawPass = Get-FirstKey $kv @('Password', 'Pwd')
    $passIsEncrypted = $false
    if (($SourceFile -and $rawPass) -or $rawPass.Length -gt 80 -or $rawPass -match '^[0-9A-Fa-f]{40,}$' -or $ExtraParams.ContainsKey('Password')) {
        $passIsEncrypted = $true
        $rawPass = ''
    }

    # SSH Tunnel parameters (covering A5M2 / SecureBridge / OpenSSH naming variants)
    $sshHost = Get-FirstKey $kv @(
        '__SSHHost', '__SSHHostName', '__SSHServer', '__SSHServerName', '__SSH_Host', '__SSH2Host',
        'SSHHost', 'SSHHostName', 'SSHServer', 'SSHServerName', 'SSH2Host', 'SSH2HostName', 'SSHTunnelHost'
    )
    $sshPortStr = Get-FirstKey $kv @(
        '__SSHPort', '__SSH_Port', '__SSH2Port', 'SSHPort', 'SSH2Port', 'SSHTunnelPort'
    ) '22'
    $sshPort = 22
    [void][int]::TryParse($sshPortStr, [ref]$sshPort)

    $sshUser = Get-FirstKey $kv @(
        '__SSHUser', '__SSHUserName', '__SSH_User', '__SSH2User',
        'SSHUser', 'SSHUserName', 'SSH2User', 'SSH2UserName', 'SSHTunnelUser'
    )
    $sshKeyFile = Get-FirstKey $kv @(
        '__SSHKeyFile', '__SSHPrivateKey', '__SSHPrivateKeyFile', '__SSH_KeyFile', '__SSHIdentityFile',
        'SSHKeyFile', 'SSHPrivateKey', 'SSHPrivateKeyFile', 'SSHIdentityFile', 'PrivateKeyFile', 'KeyFile'
    )
    $useSshRaw = Get-FirstKey $kv @(
        '__UseSSH', '__UseSSHTunnel', '__UseSSH2', '__SSH', 'UseSSH', 'UseSSHTunnel', 'UseSSH2', 'SSHEnabled'
    )
    $useSsh = ($useSshRaw -match '^(1|true|yes|on)$') -or
              ($sshHost -and $useSshRaw -notmatch '^(0|false|no|off)$')

    # SSL/TLS settings for MySQL / Aurora
    $sslRaw = Get-FirstKey $kv @('SSLMode', 'SslMode', 'UseSSL', 'SSL', '__UseSSL')
    $sslMode = 'Preferred'
    if ($sslRaw -match '^(0|false|no|disabled|none)$') { $sslMode = 'Disabled' }
    elseif ($sslRaw -match '^(verify_ca|verifyca|verify_identity|verifyfull)$') { $sslMode = 'VerifyCA' }
    elseif ($sslRaw -match '^(1|true|yes|required|require)$') { $sslMode = 'Required' }
    $sslCa = Get-FirstKey $kv @('SSLCACert', 'SSLCA', 'SslCa', 'CACert')

    if ($SourceFile) {
        $listDir = [System.IO.Path]::GetDirectoryName($SourceFile)
        if ($sshKeyFile -and -not [System.IO.Path]::IsPathRooted($sshKeyFile)) {
            $sshKeyFile = [System.IO.Path]::Combine($listDir, $sshKeyFile)
        }
        if ($sslCa -and -not [System.IO.Path]::IsPathRooted($sslCa)) {
            $sslCa = [System.IO.Path]::Combine($listDir, $sslCa)
        }
    }

    return [PSCustomObject]@{
        RegisteredName  = $DbName
        ProviderName    = $provider
        ServerName      = $server
        Port            = $port
        Database        = $database
        UserName        = $user
        PlainPassword   = $rawPass
        PassIsEncrypted = $passIsEncrypted
        UseSshTunnel    = $useSsh
        SshHost         = $sshHost
        SshPort         = $sshPort
        SshUser         = $sshUser
        SshKeyFile      = $sshKeyFile
        SslMode         = $sslMode
        SslCaCert       = $sslCa
        SourceRegistry  = $SourceRegistryPath
        SourceFile      = $SourceFile
        RawKeys         = @($kv.Keys | Sort-Object)
    }
}

function Read-A5DatabaseList {
    param([Parameter(Mandatory = $true)][string]$Path)

    $file = Get-Item -LiteralPath $Path -ErrorAction Stop
    if ($file -isnot [System.IO.FileInfo] -or $file.Extension -notin '.a5dblist', '.dblist') {
        throw 'A5:SQL Mk-2 の［データベースの追加と削除］からエクスポートした DBリスト (.a5dblist / .dblist) を指定してください。'
    }

    # ReadAllLines detects Unicode BOMs. Older Japanese text files can also be CP932.
    try {
        $lines = [System.IO.File]::ReadAllLines($file.FullName, [System.Text.UTF8Encoding]::new($false, $true))
    }
    catch [System.Text.DecoderFallbackException] {
        $lines = [System.IO.File]::ReadAllLines($file.FullName, [System.Text.Encoding]::GetEncoding(932))
    }

    $entries = [ordered]@{}
    $inList = $false
    $foundList = $false
    for ($i = 0; $i -lt $lines.Length; $i++) {
        $line = $lines[$i].Trim()
        if (-not $line -or $line.StartsWith('#') -or $line.StartsWith(';')) { continue }
        if ($line -match '^\[[^\]]+\]$') {
            $inList = $line -eq '[DatabaseList]'
            if ($inList) { $foundList = $true }
            continue
        }
        if (-not $inList) { continue }

        $name, $connection = $line -split '=', 2
        $name = $name.Trim()
        if (-not $name -or $connection -notmatch '^\s*[^=;]+=') {
            throw "DBリスト '$($file.FullName)' の $($i + 1) 行目に有効な接続設定がありません。"
        }
        if ($entries.Contains($name)) {
            throw "DBリスト '$($file.FullName)' の $($i + 1) 行目で登録DB名が重複しています。"
        }
        try {
            $entries[$name] = Parse-A5ConnectionString -DbName $name -RawConnectionString $connection -SourceFile $file.FullName
        }
        catch [System.ArgumentException] {
            throw "DBリスト '$($file.FullName)' の $($i + 1) 行目の接続設定を読み取れません。"
        }
    }
    if (-not $foundList) { throw "DBリスト '$($file.FullName)' に [DatabaseList] セクションがありません。" }
    return $entries.psbase.Values
}

function Get-A5RegisteredDatabases {
    param(
        [string]$RegistryRoot = 'HKCU:\Software\mmatsubara',
        [string]$DbListPath = ''
    )

    if ($DbListPath) { return (Read-A5DatabaseList -Path $DbListPath) }

    $results = New-Object System.Collections.Generic.List[object]
    if (-not (Test-Path -LiteralPath $RegistryRoot)) {
        return @()
    }

    $a5Keys = Get-ChildItem -LiteralPath $RegistryRoot -ErrorAction SilentlyContinue |
              Where-Object { $_.PSChildName -like 'A5M2*' }

    foreach ($appKey in $a5Keys) {
        $dsPath = Join-Path $appKey.PSPath 'DataSources'
        if (-not (Test-Path -LiteralPath $dsPath)) { continue }

        $props = Get-ItemProperty -LiteralPath $dsPath -ErrorAction SilentlyContinue
        if ($null -eq $props) { continue }

        foreach ($p in $props.PSObject.Properties) {
            if ($p.Name -match '^PS(Path|ParentPath|ChildName|Drive|Provider)$') { continue }
            $val = [string]$p.Value
            if ([string]::IsNullOrWhiteSpace($val)) { continue }

            # Read optional "<DBName>.params" subkey (without exposing encrypted passwords)
            $extra = @{}
            $paramSubKey = Join-Path $dsPath ($p.Name + '.params')
            if (Test-Path -LiteralPath $paramSubKey) {
                $subProps = Get-ItemProperty -LiteralPath $paramSubKey -ErrorAction SilentlyContinue
                if ($null -ne $subProps) {
                    foreach ($sp in $subProps.PSObject.Properties) {
                        if ($sp.Name -match '^PS(Path|ParentPath|ChildName|Drive|Provider)$') { continue }
                        $extra[$sp.Name] = [string]$sp.Value
                    }
                }
            }

            $parsed = Parse-A5ConnectionString -DbName $p.Name -RawConnectionString $val -ExtraParams $extra -SourceRegistryPath $dsPath
            $results.Add($parsed)
        }
    }

    return $results.ToArray()
}

function Get-FreeLocalTcpPort {
    $listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, 0)
    $listener.Start()
    $port = ([System.Net.IPEndPoint]$listener.LocalEndpoint).Port
    $listener.Stop()
    return $port
}

function Wait-BackgroundTask {
    param(
        [Parameter(Mandatory = $true)]$Handle,
        [scriptblock]$LogCallback = $null,
        [scriptblock]$PumpCallback = $null,
        [scriptblock]$RowProgressCallback = $null,
        [AuroraAnalyzer.OperationCancelState]$CancelState = $null
    )

    $pollMs = if ($null -ne $PumpCallback) { 25 } else { 50 }
    $lastRows = 0L
    $finished = $false
    try {
        do {
            $finished = $Handle.WaitStep($pollMs)
            foreach ($m in $Handle.DrainMessages()) {
                if ($null -ne $LogCallback) { & $LogCallback $m -1 }
                else { Write-Host $m }
            }
            if ($null -ne $RowProgressCallback -and -not $finished) {
                $rowsWritten = $Handle.RowsWritten
                if ($rowsWritten -gt $lastRows) {
                    & $RowProgressCallback $rowsWritten
                    $lastRows = $rowsWritten
                }
            }
            if ($null -ne $PumpCallback) { & $PumpCallback }
        } until ($finished)
        return $Handle.GetResult()
    }
    finally {
        if (-not $finished -and $null -ne $CancelState) { $CancelState.Cancel() }
        $Handle.Dispose()
    }
}

function Get-ExportProgressFilePath {
    param([Parameter(Mandatory = $true)][string]$ExportFolder)
    $folderPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportFolder)
    return [IO.Path]::Combine($folderPath, '_export_progress.json')
}

function Open-ExportFolderLock {
    param([Parameter(Mandatory = $true)][string]$ExportFolder)

    $folderPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportFolder)
    [void][IO.Directory]::CreateDirectory($folderPath)
    $lockPath = [IO.Path]::Combine($folderPath, '_export.lock')
    try {
        return [IO.FileStream]::new($lockPath, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite,
            [IO.FileShare]::None, 4096, [IO.FileOptions]::DeleteOnClose)
    }
    catch [IO.IOException] {
        throw "CSV出力フォルダをロックできません。他の処理が使用中か確認してください。 $folderPath ($($_.Exception.Message))"
    }
}

function Read-ExportProgress {
    param([Parameter(Mandatory = $true)][string]$ExportFolder)

    $progPath = Get-ExportProgressFilePath -ExportFolder $ExportFolder
    if (-not (Test-Path -LiteralPath $progPath -PathType Leaf)) {
        return $null
    }
    try {
        $raw = [System.IO.File]::ReadAllText($progPath, (New-Object System.Text.UTF8Encoding($false)))
        $obj = $raw | ConvertFrom-Json
        if ($null -eq $obj -or $obj.Version -notin 1,2 -or [string]::IsNullOrWhiteSpace($obj.Database) -or
            [string]::IsNullOrWhiteSpace($obj.ServerName) -or $null -eq $obj.CompletedTables -or
            $obj.CompletedTables -isnot [PSCustomObject]) { throw '進捗記録の形式が不正です。' }

        $completed = New-Object 'System.Collections.Generic.Dictionary[string, object]' ([System.StringComparer]::OrdinalIgnoreCase)
        if ($null -ne $obj.PSObject.Properties['CompletedTables'] -and $null -ne $obj.CompletedTables) {
            foreach ($p in $obj.CompletedTables.PSObject.Properties) {
                $val = $p.Value
                if ($null -ne $val) {
                    $completed[$p.Name] = @{
                        FileName      = if ($null -ne $val.PSObject.Properties['FileName']) { [string]$val.FileName } else { '' }
                        RowCount      = if ($null -ne $val.PSObject.Properties['RowCount']) { [long]$val.RowCount } else { -1L }
                        FileSizeBytes = if ($null -ne $val.PSObject.Properties['FileSizeBytes']) { [long]$val.FileSizeBytes } else { -1L }
                        CompletedAt   = if ($null -ne $val.PSObject.Properties['CompletedAt']) { [string]$val.CompletedAt } else { '' }
                    }
                }
            }
        }

        return @{
            Version         = if ($null -ne $obj.PSObject.Properties['Version']) { [int]$obj.Version } else { 1 }
            Database        = if ($null -ne $obj.PSObject.Properties['Database']) { [string]$obj.Database } else { '' }
            ServerName      = if ($null -ne $obj.PSObject.Properties['ServerName']) { [string]$obj.ServerName } else { '' }
            Port            = if ($obj.Version -ge 2) { [int]$obj.Port } else { 0 }
            MaxRowsPerTable = if ($null -ne $obj.PSObject.Properties['MaxRowsPerTable']) { [int]$obj.MaxRowsPerTable } else { 0 }
            StartedAt       = if ($null -ne $obj.PSObject.Properties['StartedAt']) { [string]$obj.StartedAt } else { '' }
            UpdatedAt       = if ($null -ne $obj.PSObject.Properties['UpdatedAt']) { [string]$obj.UpdatedAt } else { '' }
            Status          = if ($null -ne $obj.PSObject.Properties['Status']) { [string]$obj.Status } else { '' }
            TotalTables     = if ($null -ne $obj.PSObject.Properties['TotalTables']) { [int]$obj.TotalTables } else { 0 }
            CompletedTables = $completed
        }
    }
    catch {
        throw "進捗記録を読み込めません。 $progPath ($($_.Exception.Message))"
    }
}

function Save-ExportProgress {
    param(
        [Parameter(Mandatory = $true)][string]$ExportFolder,
        [Parameter(Mandatory = $true)][hashtable]$Progress
    )

    if (-not (Test-Path -LiteralPath $ExportFolder)) {
        New-Item -ItemType Directory -Path $ExportFolder -Force | Out-Null
    }
    $progPath = Get-ExportProgressFilePath -ExportFolder $ExportFolder
    $tmpPath = $progPath + '.tmp'

    $payload = [ordered]@{
        Version         = 2
        Database        = [string]$Progress.Database
        ServerName      = [string]$Progress.ServerName
        Port            = [int]$Progress.Port
        MaxRowsPerTable = [int]$Progress.MaxRowsPerTable
        StartedAt       = [string]$Progress.StartedAt
        UpdatedAt       = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Status          = [string]$Progress.Status
        TotalTables     = [int]$Progress.TotalTables
        CompletedTables = $Progress.CompletedTables
    }

    try {
        $json = $payload | ConvertTo-Json -Depth 6 -Compress
        [IO.File]::WriteAllText($tmpPath, $json)
        [AuroraAnalyzer.AtomicFile]::Publish($tmpPath, $progPath)
    }
    finally { [IO.File]::Delete($tmpPath) }
}

function Get-ExportFolderResumeSummary {
    param([string]$ExportFolder)

    if ([string]::IsNullOrWhiteSpace($ExportFolder) -or -not (Test-Path -LiteralPath $ExportFolder -PathType Container)) {
        return $null
    }
    $csvFiles = @(Get-ChildItem -LiteralPath $ExportFolder -Filter '*.csv' -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Length -gt 0 -and -not $_.Name.StartsWith('~$') -and -not $_.Name.StartsWith('_analysis_') })
    $partFiles = @(Get-ChildItem -LiteralPath $ExportFolder -Filter '*.csv.part' -File -ErrorAction SilentlyContinue)
    try { $prog = Read-ExportProgress -ExportFolder $ExportFolder }
    catch { $prog = @{ Status = 'Invalid'; CompletedTables = @{}; TotalTables = 0; Database = '' } }

    return [PSCustomObject]@{
        HasProgressFile = ($null -ne $prog)
        ProgressStatus  = if ($null -ne $prog) { $prog.Status } else { '' }
        RecordedDone    = if ($null -ne $prog) { $prog.CompletedTables.psbase.Count } else { 0 }
        TotalTables     = if ($null -ne $prog) { $prog.TotalTables } else { 0 }
        ExistingCsvCount = $csvFiles.Count
        IncompletePartCount = $partFiles.Count
        Database        = if ($null -ne $prog) { $prog.Database } else { '' }
    }
}

function Start-SshPortForward {
    param(
        [Parameter(Mandatory = $true)][string]$SshHost,
        [int]$SshPort = 22,
        [Parameter(Mandatory = $true)][string]$SshUser,
        [string]$SshKeyPath = '',
        [string]$SshPassphrase = '',
        [Parameter(Mandatory = $true)][string]$RemoteDbHost,
        [int]$RemoteDbPort = 3306,
        [int]$TimeoutSeconds = 25,
        [AuroraAnalyzer.OperationCancelState]$CancelState = $null,
        [scriptblock]$PumpCallback = $null
    )

    if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }

    $sshExe = Join-Path $env:SystemRoot 'System32\OpenSSH\ssh.exe'
    if (-not (Test-Path -LiteralPath $sshExe)) {
        $cmd = Get-Command ssh.exe -ErrorAction SilentlyContinue
        if ($null -eq $cmd) {
            throw "Windows 11 標準 OpenSSH (ssh.exe) が見つかりません。"
        }
        $sshExe = $cmd.Source
    }

    $localPort = Get-FreeLocalTcpPort
    $tempAskPass = $null

    $argList = New-Object System.Collections.Generic.List[string]
    $argList.Add('-N')
    $argList.Add('-L')
    $argList.Add("127.0.0.1:${localPort}:${RemoteDbHost}:${RemoteDbPort}")
    $argList.Add('-p')
    $argList.Add($SshPort.ToString())
    $argList.Add('-o')
    $argList.Add('ExitOnForwardFailure=yes')
    $argList.Add('-o')
    $argList.Add('ServerAliveInterval=30')
    $argList.Add('-o')
    $argList.Add('ServerAliveCountMax=3')
    $argList.Add('-o')
    $argList.Add('StrictHostKeyChecking=accept-new')

    if (-not [string]::IsNullOrWhiteSpace($SshKeyPath)) {
        if (-not (Test-Path -LiteralPath $SshKeyPath)) {
            throw "SSH秘密鍵ファイルが見つかりません: $SshKeyPath"
        }
        $argList.Add('-i')
        $argList.Add($SshKeyPath)
    }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $sshExe
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardError = $true
    $psi.RedirectStandardOutput = $true

    if (-not [string]::IsNullOrEmpty($SshPassphrase)) {
        # Safe temporary askpass helper using delayed expansion so special shell chars (& | < > ^ % !) are never interpreted
        $tempAskPass = Join-Path ([System.IO.Path]::GetTempPath()) ("a5_askpass_" + [Guid]::NewGuid().ToString('N') + ".cmd")
        Set-Content -LiteralPath $tempAskPass -Value "@echo off`r`nsetlocal EnableDelayedExpansion`r`necho(!A5_SSH_SECRET!" -Encoding Ascii
        $psi.EnvironmentVariables['A5_SSH_SECRET'] = $SshPassphrase
        $psi.EnvironmentVariables['SSH_ASKPASS'] = $tempAskPass
        $psi.EnvironmentVariables['SSH_ASKPASS_REQUIRE'] = 'force'
        $psi.EnvironmentVariables['DISPLAY'] = ':0'
    }
    else {
        $argList.Add('-o')
        $argList.Add('BatchMode=yes')
    }

    $argList.Add("${SshUser}@${SshHost}")

    # Safely quote arguments containing spaces
    $quotedArgs = @()
    foreach ($a in $argList) {
        if ($a -match '\s') { $quotedArgs += ('"' + ($a -replace '"', '\"') + '"') }
        else { $quotedArgs += $a }
    }
    $psi.Arguments = ($quotedArgs -join ' ')

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    [void]$proc.Start()

    # Drain stdout/stderr asynchronously in C# so long-running ssh.exe never deadlocks on full pipe buffers
    $drain = New-Object AuroraAnalyzer.ProcessOutputDrain
    $drain.Attach($proc)

    $cancelReg = if ($null -ne $CancelState) { $CancelState.RegisterProcessKill($proc) } else { $null }

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $ready = $false
    try {
        while ((Get-Date) -lt $deadline) {
            if ($null -ne $PumpCallback) { & $PumpCallback }
            if ($null -ne $CancelState -and $CancelState.IsCancelled) {
                try { if (-not $proc.HasExited) { $proc.Kill() } } catch { }
                $CancelState.ThrowIfCancelled()
            }
            if ($proc.HasExited) {
                try { $proc.WaitForExit() } catch { }
                $errText = $drain.GetStderrText()
                throw "SSHトンネルの確立に失敗しました (終了コード $($proc.ExitCode)): $errText"
            }
            $probe = New-Object System.Net.Sockets.TcpClient
            try {
                $iar = $probe.BeginConnect('127.0.0.1', $localPort, $null, $null)
                if ($iar.AsyncWaitHandle.WaitOne(120)) {
                    $probe.EndConnect($iar)
                    $ready = $true
                    break
                }
            }
            catch { }
            finally { $probe.Close() }
            Start-Sleep -Milliseconds 80
        }

        if (-not $ready) {
            try { if (-not $proc.HasExited) { $proc.Kill() } } catch { }
            $errText = $drain.GetStderrText()
            $errDetail = if ([string]::IsNullOrWhiteSpace($errText)) { '' } else { " ($errText)" }
            throw "SSHトンネル待機タイムアウト (${TimeoutSeconds}秒): ${SshUser}@${SshHost}:${SshPort} -> ${RemoteDbHost}:${RemoteDbPort}$errDetail"
        }
    }
    catch {
        if ($null -ne $cancelReg) { try { $cancelReg.Dispose() } catch { } }
        try { if (-not $proc.HasExited) { $proc.Kill() } } catch { }
        try { $proc.Dispose() } catch { }
        throw
    }
    finally {
        if ($tempAskPass -and (Test-Path -LiteralPath $tempAskPass)) {
            Remove-Item -LiteralPath $tempAskPass -Force -ErrorAction SilentlyContinue
        }
    }

    return [PSCustomObject]@{
        Process            = $proc
        LocalPort          = $localPort
        RemoteHost         = $RemoteDbHost
        RemotePort         = $RemoteDbPort
        CancelRegistration = $cancelReg
        OutputDrain        = $drain
    }
}

function Stop-SshPortForward {
    param($Tunnel)
    if ($null -eq $Tunnel) { return }
    try {
        if ($null -ne $Tunnel.PSObject.Properties['CancelRegistration'] -and $null -ne $Tunnel.CancelRegistration) {
            $Tunnel.CancelRegistration.Dispose()
        }
        if ($null -ne $Tunnel.Process -and -not $Tunnel.Process.HasExited) {
            $Tunnel.Process.Kill()
            [void]$Tunnel.Process.WaitForExit(3000)
        }
        if ($null -ne $Tunnel.Process) {
            $Tunnel.Process.Dispose()
        }
    }
    catch { }
}

function Open-AuroraReadOnlySession {
    param(
        [Parameter(Mandatory = $true)][hashtable]$ConnConfig,
        [scriptblock]$LogCallback = $null,
        [AuroraAnalyzer.OperationCancelState]$CancelState = $null,
        [scriptblock]$PumpCallback = $null
    )

    function Write-ConnLog {
        param([string]$Msg, [int]$Pct = -1)
        if ($null -ne $LogCallback) { & $LogCallback $Msg $Pct }
        else { Write-Host $Msg }
    }

    if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }

    if ([string]::IsNullOrWhiteSpace($ConnConfig.ServerName)) {
        throw "接続先MySQLホスト名 (ServerName) が指定されていません。"
    }
    if ([string]::IsNullOrWhiteSpace($ConnConfig.UserName)) {
        throw "MySQLユーザー名 (UserName) が指定されていません。"
    }
    if ([string]::IsNullOrWhiteSpace($ConnConfig.Database)) {
        throw "対象データベース/スキーマ名 (Database) が指定されていません。"
    }

    $tunnel = $null
    $client = $null
    try {
        $connectHost = $ConnConfig.ServerName
        $effectivePort = [int]$ConnConfig.Port
        if ($effectivePort -le 0) { $effectivePort = 3306 }
        $connectPort = $effectivePort

        if ($ConnConfig.UseSshTunnel) {
            Write-ConnLog ("SSHトンネル確立中: {0}@{1}:{2} -> {3}:{4}" -f $ConnConfig.SshUser, $ConnConfig.SshHost, $ConnConfig.SshPort, $ConnConfig.ServerName, $connectPort) 5
            $tunnel = Start-SshPortForward `
                -SshHost $ConnConfig.SshHost `
                -SshPort ([int]$ConnConfig.SshPort) `
                -SshUser $ConnConfig.SshUser `
                -SshKeyPath ([string]$ConnConfig.SshKeyPath) `
                -SshPassphrase ([string]$ConnConfig.SshPassphrase) `
                -RemoteDbHost $ConnConfig.ServerName `
                -RemoteDbPort $connectPort `
                -CancelState $CancelState `
                -PumpCallback $PumpCallback
            $connectHost = '127.0.0.1'
            $connectPort = $tunnel.LocalPort
            Write-ConnLog ("SSHトンネル確立完了 (ローカルポート: 127.0.0.1:{0})" -f $connectPort) 12
        }

        if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }
        Write-ConnLog ("Aurora / MySQL接続中: {0}:{1} (DB={2}, SSL={3})" -f $ConnConfig.ServerName, $effectivePort, $ConnConfig.Database, $ConnConfig.SslMode) 15
        $client = New-Object AuroraAnalyzer.MySql84Client
        $connHandle = [AuroraAnalyzer.BackgroundTaskLauncher]::StartConnect(
            $client,
            $connectHost,
            $connectPort,
            $ConnConfig.ServerName,
            $ConnConfig.UserName,
            [string]$ConnConfig.Password,
            $ConnConfig.Database,
            [string]$ConnConfig.SslMode,
            [string]$ConnConfig.SslCaCert,
            60,
            $CancelState
        )
        [void](Wait-BackgroundTask -Handle $connHandle -LogCallback $LogCallback -PumpCallback $PumpCallback -CancelState $CancelState)
        Write-ConnLog ("接続成功: MySQL Server {0} (TLS暗号化={1})" -f $client.ServerVersion, $client.IsSslActive) 20

        # Ensure safe read-only session & utf8mb4
        if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }
        foreach ($sql in 'SET NAMES utf8mb4', 'SET SESSION TRANSACTION READ ONLY') {
            $setup = [AuroraAnalyzer.BackgroundTaskLauncher]::StartExecuteQuery($client, $sql, $CancelState)
            [void](Wait-BackgroundTask -Handle $setup -PumpCallback $PumpCallback -CancelState $CancelState)
        }

        return [PSCustomObject]@{
            Client        = $client
            Tunnel        = $tunnel
            EffectivePort = $effectivePort
        }
    }
    catch {
        if ($null -ne $client) { try { $client.Dispose() } catch { } }
        if ($null -ne $tunnel) { Stop-SshPortForward -Tunnel $tunnel }
        if ($null -ne $CancelState -and $CancelState.IsCancelled -and $_.Exception -isnot [System.OperationCanceledException]) {
            $CancelState.ThrowIfCancelled()
        }
        throw
    }
}

function Get-AuroraBaseTables {
    param(
        [Parameter(Mandatory = $true)]$Client,
        [string[]]$IncludeTables = @(),
        [string[]]$ExcludeTables = @(),
        [AuroraAnalyzer.OperationCancelState]$CancelState = $null,
        [scriptblock]$LogCallback = $null,
        [scriptblock]$PumpCallback = $null,
        [hashtable]$EstimatedRows = $null
    )

    if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }
    $sql = if ($null -eq $EstimatedRows) { "SHOW FULL TABLES WHERE Table_type = 'BASE TABLE'" }
           else { "SELECT TABLE_NAME, TABLE_ROWS FROM information_schema.tables WHERE TABLE_SCHEMA = DATABASE() AND TABLE_TYPE = 'BASE TABLE' ORDER BY TABLE_NAME" }
    $qHandle = [AuroraAnalyzer.BackgroundTaskLauncher]::StartExecuteQuery($Client, $sql, $CancelState)
    $tblRes = Wait-BackgroundTask -Handle $qHandle -LogCallback $LogCallback -PumpCallback $PumpCallback -CancelState $CancelState
    $tableNames = New-Object System.Collections.Generic.List[string]
    foreach ($r in $tblRes.Rows) {
        if ($r.Length -gt 0 -and -not [string]::IsNullOrWhiteSpace($r[0])) {
            $tName = $r[0]
            if ($IncludeTables.Count -gt 0 -and ($IncludeTables -notcontains $tName)) { continue }
            if ($ExcludeTables.Count -gt 0 -and ($ExcludeTables -contains $tName)) { continue }
            $tableNames.Add($tName)
            if ($null -ne $EstimatedRows) {
                $estimate = 0L
                [void][long]::TryParse($r[1], [ref]$estimate)
                $EstimatedRows[$tName] = $estimate
            }
        }
    }
    return $tableNames.ToArray()
}

function Invoke-AuroraBulkCsvExport {
    param(
        [Parameter(Mandatory = $true)][hashtable]$ConnConfig,
        [Parameter(Mandatory = $true)][string]$ExportFolder,
        [int]$MaxRowsPerTable = 0,
        [string[]]$IncludeTables = @(),
        [string[]]$ExcludeTables = @(),
        [scriptblock]$LogCallback = $null,
        [switch]$Resume,
        [AuroraAnalyzer.OperationCancelState]$CancelState = $null,
        [scriptblock]$PumpCallback = $null,
        [IO.FileStream]$FolderLock = $null,
        [int]$ExportEndPercent = 100,
        [AuroraAnalyzer.CsvExportAnalysis]$ExportAnalysis = $null
    )

    function Write-ExportLog {
        param([string]$Msg, [int]$Pct = -1, [int]$CsvPct = -1)
        if ($null -ne $LogCallback) { & $LogCallback $Msg $Pct $CsvPct }
        else { Write-Host $Msg }
    }

    function Write-CsvProgress([double]$Done, [string]$Detail, [switch]$Finished) {
        $csvPct = if ($Finished) { 100 } else { [int][Math]::Min(99, [Math]::Floor(100 * $Done / $tableNames.Count)) }
        $overall = 20 + [int][Math]::Floor(($ExportEndPercent - 20) * $csvPct / 100)
        Write-ExportLog ("CSV出力 {0}%  {1}" -f $csvPct, $Detail) $overall $csvPct
    }

    if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }

    if ([string]::IsNullOrWhiteSpace($ConnConfig.ServerName)) {
        throw "接続先MySQLホスト名 (ServerName) が指定されていません。"
    }
    if ([string]::IsNullOrWhiteSpace($ConnConfig.UserName)) {
        throw "MySQLユーザー名 (UserName) が指定されていません。"
    }
    if ([string]::IsNullOrWhiteSpace($ConnConfig.Database)) {
        throw "対象データベース/スキーマ名 (Database) が指定されていません。"
    }

    $ExportFolder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportFolder)
    $ownedLock = $null
    $checkpointStarted = $false
    $prevProgress = $null
    $completedMap = New-Object 'System.Collections.Generic.Dictionary[string, object]' ([System.StringComparer]::OrdinalIgnoreCase)
    $progressState = @{
        Database        = [string]$ConnConfig.Database
        ServerName      = [string]$ConnConfig.ServerName
        Port            = if ([int]$ConnConfig.Port -gt 0) { [int]$ConnConfig.Port } else { 3306 }
        MaxRowsPerTable = [int]$MaxRowsPerTable
        StartedAt       = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Status          = 'InProgress'
        TotalTables     = 0
        CompletedTables = $completedMap
    }

    $sessionConn = $null
    $exportedFiles = New-Object System.Collections.Generic.List[object]
    $tableNames = @()
    $newlyExportedCount = 0
    $skippedCount = 0

    try {
        if ($null -eq $FolderLock) { $ownedLock = Open-ExportFolderLock $ExportFolder }
        if ($Resume) { $prevProgress = Read-ExportProgress -ExportFolder $ExportFolder }
        if ($null -ne $prevProgress) {
            if ($prevProgress.ServerName -ne $progressState.ServerName -or $prevProgress.Database -cne $progressState.Database -or
                ($prevProgress.Port -ne 0 -and $prevProgress.Port -ne $progressState.Port) -or
                $prevProgress.MaxRowsPerTable -ne $MaxRowsPerTable) {
                throw '前回と接続先・データベース・行数上限が一致しません。再開には同じ設定を指定してください。'
            }
            $progressState.StartedAt = $prevProgress.StartedAt
        }
        $sessionConn = Open-AuroraReadOnlySession -ConnConfig $ConnConfig -LogCallback $LogCallback -CancelState $CancelState -PumpCallback $PumpCallback
        $client = $sessionConn.Client

        $estimatedRows = @{}
        $tableNames = @(Get-AuroraBaseTables -Client $client -IncludeTables $IncludeTables -ExcludeTables $ExcludeTables -CancelState $CancelState -LogCallback $LogCallback -PumpCallback $PumpCallback -EstimatedRows $estimatedRows)
        if ($tableNames.Count -eq 0) {
            throw "対象データベース '$($ConnConfig.Database)' にエクスポート対象のBASE TABLEが見つかりませんでした。"
        }

        $progressState.TotalTables = $tableNames.Count

        # Pre-evaluate which tables can be resumed/skipped when -Resume is active
        $skipPlan = @{}
        $fileNames = @{}
        $seenFiles = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($tbl in $tableNames) {
            if ($null -ne $PumpCallback) { & $PumpCallback }
            if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }
            $safeStem = ($tbl -replace '[\\/:*?"<>|]', '_').TrimEnd(' .')
            if ([string]::IsNullOrWhiteSpace($safeStem)) {
                throw "テーブル名を一意なWindowsファイル名として保存できません。 $tbl"
            }
            $safeFile = $safeStem + '.csv'
            if (-not $seenFiles.Add($safeFile) -or $safeFile -match '^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])\.') {
                throw "テーブル名を一意なWindowsファイル名として保存できません。 $tbl ($safeFile)"
            }
            $fileNames[$tbl] = $safeFile
            if (-not $Resume) { continue }
            $csvPath = Join-Path $ExportFolder $safeFile
            $fi = [IO.FileInfo]::new($csvPath)
            if (-not $fi.Exists -or $fi.Length -eq 0 -or [IO.File]::Exists($csvPath + '.part')) { continue }
            if ($null -ne $prevProgress) {
                $rec = $prevProgress.CompletedTables[$tbl]
                if ($null -eq $rec -or $rec.FileName -cne $safeFile -or $rec.FileSizeBytes -ne $fi.Length) { continue }
            }
            else {
                # Legacy folders have no manifest. Only this explicit -Resume case can adopt existing CSVs.
                $rec = @{ FileName = $safeFile; RowCount = -1L; FileSizeBytes = $fi.Length; CompletedAt = $fi.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss') }
            }
            $skipPlan[$tbl] = $rec
            $completedMap[$tbl] = $rec
        }
        if ($Resume) {
            $remCount = $tableNames.Count - $skipPlan.psbase.Count
            Write-CsvProgress $skipPlan.psbase.Count ("再開、出力済み {0}/{1} テーブル、残り {2} テーブル" -f $skipPlan.psbase.Count, $tableNames.Count, $remCount)
        }
        else {
            Write-CsvProgress 0 ("エクスポート対象テーブル数: {0} 件" -f $tableNames.Count)
        }

        Save-ExportProgress -ExportFolder $ExportFolder -Progress $progressState
        $checkpointStarted = $true

        $idx = 0
        foreach ($tbl in $tableNames) {
            if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }
            $idx++
            $safeFile = $fileNames[$tbl]
            $csvPath = Join-Path $ExportFolder $safeFile
            $done = $skipPlan.psbase.Count + $newlyExportedCount

            if ($Resume -and $skipPlan.ContainsKey($tbl)) {
                $sp = $skipPlan[$tbl]
                $skippedCount++
                $exportedFiles.Add([PSCustomObject]@{
                    TableName = $tbl
                    CsvPath   = $csvPath
                    RowCount  = $sp.RowCount
                    Skipped   = $true
                })
                $rowDesc = if ($sp.RowCount -ge 0) { ("{0:N0} 件 / " -f $sp.RowCount) } else { '' }
                Write-CsvProgress $done ("[{0}/{1}] スキップ: {2} -> {3} ({4}{5:N0} B)" -f $idx, $tableNames.Count, $tbl, $safeFile, $rowDesc, $sp.FileSizeBytes)
                continue
            }

            $escapedTbl = $tbl.Replace('`', '``')
            $sql = "SELECT * FROM ``$escapedTbl``"
            if ($MaxRowsPerTable -gt 0) {
                $sql = "$sql LIMIT $MaxRowsPerTable"
            }

            Write-CsvProgress $done ("[{0}/{1}] エクスポート中: {2} -> {3}" -f $idx, $tableNames.Count, $tbl, $safeFile)
            $estimate = [long]$estimatedRows[$tbl]
            if ($MaxRowsPerTable -gt 0 -and $estimate -gt $MaxRowsPerTable) { $estimate = $MaxRowsPerTable }
            $onRows = {
                param([long]$written)
                if ($estimate -gt 0) {
                    $fraction = [Math]::Min(0.99, $written / [double]$estimate)
                    Write-CsvProgress ($done + $fraction) ("(推定) {0}  {1:N0} 件出力 / 推定 {2:N0} 件" -f $tbl, $written, $estimate)
                }
                else {
                    Write-CsvProgress $done ("{0}  {1:N0} 件出力 (総件数不明、完了テーブル基準)" -f $tbl, $written)
                }
            }
            $expHandle = [AuroraAnalyzer.BackgroundTaskLauncher]::StartExportQueryToCsv($client, $tbl, $sql, $csvPath, $MaxRowsPerTable, $true, $CancelState, $ExportAnalysis)
            $rows = Wait-BackgroundTask -Handle $expHandle -LogCallback $LogCallback -PumpCallback $PumpCallback -RowProgressCallback $onRows -CancelState $CancelState
            $newlyExportedCount++

            $doneSize = [IO.FileInfo]::new($csvPath).Length
            $completedMap[$tbl] = @{
                FileName      = $safeFile
                RowCount      = [long]$rows
                FileSizeBytes = $doneSize
                CompletedAt   = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            }
            Save-ExportProgress -ExportFolder $ExportFolder -Progress $progressState

            $exportedFiles.Add([PSCustomObject]@{
                TableName = $tbl
                CsvPath   = $csvPath
                RowCount  = $rows
                Skipped   = $false
            })
            Write-CsvProgress ($done + 1) ("完了: {0} ({1:N0} 件)" -f $tbl, $rows)
        }

        if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }
        $progressState.Status = if ($null -ne $ExportAnalysis) { 'CsvCompleted' } else { 'Completed' }
        Save-ExportProgress -ExportFolder $ExportFolder -Progress $progressState

        Write-CsvProgress $tableNames.Count ("CSV一括エクスポート完了: {0} テーブル (新規出力 {1} / スキップ {2}) -> {3}" -f $exportedFiles.Count, $newlyExportedCount, $skippedCount, $ExportFolder) -Finished
        return $exportedFiles.ToArray()
    }
    catch {
        $isCancelled = ($_.Exception -is [System.OperationCanceledException]) -or ($null -ne $CancelState -and $CancelState.IsCancelled)
        if ($checkpointStarted) {
            $progressState.Status = 'Interrupted'
            try { Save-ExportProgress -ExportFolder $ExportFolder -Progress $progressState }
            catch { Write-Warning "中断状態の記録に失敗しました。最後に保存した進捗から再開できます。 $($_.Exception.Message)" }
        }
        if ($isCancelled) {
            Write-ExportLog ("CSVエクスポートを途中で中断しました (完了済み: {0}/{1} テーブル)。同じフォルダを指定し「出力済みCSVから続行」(-Resume)で再開できます。" -f $completedMap.psbase.Count, $progressState.TotalTables) -1
            if ($_.Exception -isnot [System.OperationCanceledException]) {
                throw (New-Object System.OperationCanceledException("ユーザーによって処理が中断されました。"))
            }
        }
        throw
    }
    finally {
        if ($null -ne $ownedLock) { $ownedLock.Dispose() }
        if ($null -ne $sessionConn) {
            if ($null -ne $sessionConn.Client) { $sessionConn.Client.Dispose() }
            if ($null -ne $sessionConn.Tunnel) { Stop-SshPortForward -Tunnel $sessionConn.Tunnel }
        }
    }
}

function Get-AuroraDirectDbMetadata {
    param(
        [Parameter(Mandatory = $true)]$Client,
        [AuroraAnalyzer.OperationCancelState]$CancelState = $null,
        [scriptblock]$LogCallback = $null,
        [scriptblock]$PumpCallback = $null
    )

    $srvMeta = New-Object AuroraAnalyzer.DbServerMetadata
    $tblMetaMap = New-Object 'System.Collections.Generic.Dictionary[string, AuroraAnalyzer.DbTableMetadata]' ([System.StringComparer]::OrdinalIgnoreCase)
    $colMetaMap = New-Object 'System.Collections.Generic.Dictionary[string, System.Collections.Generic.Dictionary[string, AuroraAnalyzer.DbColumnMetadata]]' ([System.StringComparer]::OrdinalIgnoreCase)

    function Invoke-OptionalMetaQuery([string]$Sql) {
        if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }
        try {
            $qh = [AuroraAnalyzer.BackgroundTaskLauncher]::StartExecuteQuery($Client, $Sql, $CancelState)
            return (Wait-BackgroundTask -Handle $qh -PumpCallback $PumpCallback -CancelState $CancelState)
        }
        catch [System.OperationCanceledException] { throw }
        catch { return $null }
    }

    function Get-ColIndexMap($QueryResult) {
        $m = @{}
        if ($null -eq $QueryResult -or $null -eq $QueryResult.Columns) { return $m }
        for ($i = 0; $i -lt $QueryResult.Columns.Count; $i++) {
            $cn = $QueryResult.Columns[$i].Name
            if (-not [string]::IsNullOrEmpty($cn)) { $m[$cn.ToUpperInvariant()] = $i }
        }
        return $m
    }

    function Get-RowStr([string[]]$Row, $Map, [string]$Col) {
        $k = $Col.ToUpperInvariant()
        if (-not $Map.ContainsKey($k)) { return '' }
        $idx = [int]$Map[$k]
        if ($idx -lt 0 -or $idx -ge $Row.Length -or $null -eq $Row[$idx]) { return '' }
        return [string]$Row[$idx]
    }

    function Get-RowLong([string[]]$Row, $Map, [string]$Col) {
        $s = Get-RowStr $Row $Map $Col
        $v = 0L
        if (-not [string]::IsNullOrWhiteSpace($s) -and [long]::TryParse($s.Trim(), [System.Globalization.NumberStyles]::Integer, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$v)) {
            return $v
        }
        return 0L
    }

    # 1. Capture SHOW GLOBAL STATUS FIRST (before SELECT * scans) for clean test workload delta
    $statusSql = "SHOW GLOBAL STATUS WHERE Variable_name IN ('Uptime','Questions','Queries','Com_select','Com_insert','Com_update','Com_delete','Com_replace','Com_commit','Com_rollback','Slow_queries','Innodb_rows_read','Innodb_rows_inserted','Innodb_rows_updated','Innodb_rows_deleted','Innodb_row_lock_waits','Innodb_row_lock_time','Innodb_row_lock_time_max','Innodb_buffer_pool_read_requests','Innodb_buffer_pool_reads','Innodb_buffer_pool_pages_data','Innodb_buffer_pool_pages_dirty','Innodb_buffer_pool_pages_total','Innodb_data_reads','Innodb_data_writes','Innodb_os_log_written','Created_tmp_tables','Created_tmp_disk_tables','Select_full_join','Select_scan','Sort_merge_passes','Threads_connected','Threads_running','Max_used_connections','Aborted_connects','Aborted_clients','Bytes_received','Bytes_sent','Handler_read_key','Handler_read_rnd_next')"
    $stRes = Invoke-OptionalMetaQuery $statusSql
    if ($null -ne $stRes -and $stRes.Columns.Count -ge 2 -and $stRes.Columns[0].Name -ieq 'Variable_name' -and $stRes.Columns[1].Name -ieq 'Value') {
        foreach ($r in $stRes.Rows) {
            if ($r.Length -lt 2 -or [string]::IsNullOrWhiteSpace($r[0])) { continue }
            $lv = 0L
            if ([long]::TryParse($r[1], [System.Globalization.NumberStyles]::Integer, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$lv)) {
                $srvMeta.StatusCounters[$r[0].Trim()] = $lv
            }
        }
        if ($srvMeta.StatusCounters.Count -gt 0) { $srvMeta.HasStatus = $true }
    }

    # 2. Server & Aurora environment variables
    $varSql = "SHOW GLOBAL VARIABLES WHERE Variable_name IN ('version','version_comment','aurora_version','innodb_buffer_pool_size','max_connections','transaction_isolation','innodb_page_size','character_set_server','time_zone')"
    $varRes = Invoke-OptionalMetaQuery $varSql
    if ($null -ne $varRes -and $varRes.Columns.Count -ge 2 -and $varRes.Columns[0].Name -ieq 'Variable_name' -and $varRes.Columns[1].Name -ieq 'Value') {
        foreach ($r in $varRes.Rows) {
            if ($r.Length -lt 2 -or [string]::IsNullOrWhiteSpace($r[0])) { continue }
            $srvMeta.Variables[$r[0].Trim()] = if ($null -ne $r[1]) { $r[1].Trim() } else { '' }
        }
        if ($srvMeta.Variables.Count -gt 0) { $srvMeta.HasVariables = $true }
    }

    # 3. information_schema.TABLES (physical InnoDB DATA/INDEX/FREE sizes, AUTO_INCREMENT, UPDATE_TIME)
    $tblSql = "SELECT TABLE_NAME, ENGINE, ROW_FORMAT, TABLE_ROWS, AVG_ROW_LENGTH, DATA_LENGTH, INDEX_LENGTH, DATA_FREE, AUTO_INCREMENT, CREATE_TIME, UPDATE_TIME, TABLE_COLLATION, TABLE_COMMENT FROM information_schema.TABLES WHERE TABLE_SCHEMA = DATABASE() AND TABLE_TYPE = 'BASE TABLE' ORDER BY TABLE_NAME"
    $tblRes = Invoke-OptionalMetaQuery $tblSql
    $tMap = Get-ColIndexMap $tblRes
    if ($tMap.ContainsKey('TABLE_NAME') -and $tMap.ContainsKey('ENGINE') -and $tMap.ContainsKey('DATA_LENGTH')) {
        foreach ($r in $tblRes.Rows) {
            $tn = (Get-RowStr $r $tMap 'TABLE_NAME').Trim()
            if ([string]::IsNullOrEmpty($tn)) { continue }
            $tm = New-Object AuroraAnalyzer.DbTableMetadata
            $tm.HasMetadata    = $true
            $tm.Engine         = Get-RowStr $r $tMap 'ENGINE'
            $tm.RowFormat      = Get-RowStr $r $tMap 'ROW_FORMAT'
            $tm.EstTableRows   = Get-RowLong $r $tMap 'TABLE_ROWS'
            $tm.AvgRowLength   = Get-RowLong $r $tMap 'AVG_ROW_LENGTH'
            $tm.DataLength     = Get-RowLong $r $tMap 'DATA_LENGTH'
            $tm.IndexLength    = Get-RowLong $r $tMap 'INDEX_LENGTH'
            $tm.DataFree       = Get-RowLong $r $tMap 'DATA_FREE'
            $aiRaw = Get-RowStr $r $tMap 'AUTO_INCREMENT'
            $aiVal = 0L
            if (-not [string]::IsNullOrWhiteSpace($aiRaw) -and [long]::TryParse($aiRaw.Trim(), [System.Globalization.NumberStyles]::Integer, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$aiVal)) {
                $tm.HasAutoIncrement = $true
                $tm.AutoIncrement    = $aiVal
            }
            $tm.CreateTime     = Get-RowStr $r $tMap 'CREATE_TIME'
            $tm.UpdateTime     = Get-RowStr $r $tMap 'UPDATE_TIME'
            $tm.TableCollation = Get-RowStr $r $tMap 'TABLE_COLLATION'
            $tm.TableComment   = Get-RowStr $r $tMap 'TABLE_COMMENT'
            $tblMetaMap[$tn] = $tm
        }
    }

    # 4. information_schema.COLUMNS (exact DDL types, nullability, PK/UK/MUL keys, defaults, extra)
    $colSql = "SELECT TABLE_NAME, COLUMN_NAME, ORDINAL_POSITION, COLUMN_TYPE, IS_NULLABLE, COLUMN_KEY, COLUMN_DEFAULT, EXTRA, CHARACTER_SET_NAME, COLLATION_NAME, COLUMN_COMMENT FROM information_schema.COLUMNS WHERE TABLE_SCHEMA = DATABASE() ORDER BY TABLE_NAME, ORDINAL_POSITION"
    $colRes = Invoke-OptionalMetaQuery $colSql
    $cMap = Get-ColIndexMap $colRes
    if ($cMap.ContainsKey('TABLE_NAME') -and $cMap.ContainsKey('COLUMN_NAME') -and $cMap.ContainsKey('COLUMN_TYPE')) {
        foreach ($r in $colRes.Rows) {
            $tn = (Get-RowStr $r $cMap 'TABLE_NAME').Trim()
            $cn = (Get-RowStr $r $cMap 'COLUMN_NAME').Trim()
            if ([string]::IsNullOrEmpty($tn) -or [string]::IsNullOrEmpty($cn)) { continue }
            $perTbl = $null
            if (-not $colMetaMap.TryGetValue($tn, [ref]$perTbl)) {
                $perTbl = New-Object 'System.Collections.Generic.Dictionary[string, AuroraAnalyzer.DbColumnMetadata]' ([System.StringComparer]::OrdinalIgnoreCase)
                $colMetaMap[$tn] = $perTbl
            }
            $cm = New-Object AuroraAnalyzer.DbColumnMetadata
            $cm.ColumnName       = $cn
            $cm.OrdinalPosition  = [int](Get-RowLong $r $cMap 'ORDINAL_POSITION')
            $cm.ColumnType       = Get-RowStr $r $cMap 'COLUMN_TYPE'
            $cm.IsNullable       = Get-RowStr $r $cMap 'IS_NULLABLE'
            $cm.ColumnKey        = Get-RowStr $r $cMap 'COLUMN_KEY'
            $cm.ColumnDefault    = Get-RowStr $r $cMap 'COLUMN_DEFAULT'
            $cm.Extra            = Get-RowStr $r $cMap 'EXTRA'
            $cm.CharacterSetName = Get-RowStr $r $cMap 'CHARACTER_SET_NAME'
            $cm.CollationName    = Get-RowStr $r $cMap 'COLLATION_NAME'
            $cm.ColumnComment    = Get-RowStr $r $cMap 'COLUMN_COMMENT'
            $perTbl[$cn] = $cm
        }
    }

    # 5. information_schema.STATISTICS (index definitions & optimizer cardinality)
    $idxSql = "SELECT TABLE_NAME, INDEX_NAME, NON_UNIQUE, SEQ_IN_INDEX, COLUMN_NAME, CARDINALITY, INDEX_TYPE FROM information_schema.STATISTICS WHERE TABLE_SCHEMA = DATABASE() ORDER BY TABLE_NAME, INDEX_NAME, SEQ_IN_INDEX"
    $idxRes = Invoke-OptionalMetaQuery $idxSql
    $iMap = Get-ColIndexMap $idxRes
    if ($iMap.ContainsKey('TABLE_NAME') -and $iMap.ContainsKey('INDEX_NAME') -and $iMap.ContainsKey('COLUMN_NAME')) {
        $idxAgg = @{}
        foreach ($r in $idxRes.Rows) {
            $tn = (Get-RowStr $r $iMap 'TABLE_NAME').Trim()
            $in = (Get-RowStr $r $iMap 'INDEX_NAME').Trim()
            $cn = (Get-RowStr $r $iMap 'COLUMN_NAME').Trim()
            if ([string]::IsNullOrEmpty($tn) -or [string]::IsNullOrEmpty($in)) { continue }
            $k = "$tn`0$in"
            if (-not $idxAgg.ContainsKey($k)) {
                $idxAgg[$k] = @{
                    TableName   = $tn
                    IndexName   = $in
                    NonUnique   = ((Get-RowLong $r $iMap 'NON_UNIQUE') -ne 0)
                    IndexType   = Get-RowStr $r $iMap 'INDEX_TYPE'
                    Cols        = (New-Object System.Collections.Generic.List[string])
                    Cardinality = Get-RowLong $r $iMap 'CARDINALITY'
                }
            }
            $idxAgg[$k].Cols.Add($cn)
            $cardVal = Get-RowLong $r $iMap 'CARDINALITY'
            if ($cardVal -gt $idxAgg[$k].Cardinality) { $idxAgg[$k].Cardinality = $cardVal }
        }
        foreach ($item in $idxAgg.Values) {
            $tm = $null
            if (-not $tblMetaMap.TryGetValue($item.TableName, [ref]$tm)) {
                $tm = New-Object AuroraAnalyzer.DbTableMetadata
                $tblMetaMap[$item.TableName] = $tm
            }
            $im = New-Object AuroraAnalyzer.DbIndexMetadata
            $im.IndexName   = $item.IndexName
            $im.NonUnique   = [bool]$item.NonUnique
            $im.IndexType   = $item.IndexType
            $im.Columns     = ($item.Cols.ToArray() -join ', ')
            $im.Cardinality = [long]$item.Cardinality
            $tm.Indexes.Add($im)
        }
    }

    # 6. information_schema.KEY_COLUMN_USAGE (declared foreign key constraints)
    $fkSql = "SELECT CONSTRAINT_NAME, TABLE_NAME, COLUMN_NAME, REFERENCED_TABLE_NAME, REFERENCED_COLUMN_NAME FROM information_schema.KEY_COLUMN_USAGE WHERE TABLE_SCHEMA = DATABASE() AND REFERENCED_TABLE_NAME IS NOT NULL ORDER BY TABLE_NAME, CONSTRAINT_NAME, ORDINAL_POSITION"
    $fkRes = Invoke-OptionalMetaQuery $fkSql
    $fMap = Get-ColIndexMap $fkRes
    if ($fMap.ContainsKey('CONSTRAINT_NAME') -and $fMap.ContainsKey('TABLE_NAME') -and $fMap.ContainsKey('REFERENCED_TABLE_NAME')) {
        foreach ($r in $fkRes.Rows) {
            $tn = (Get-RowStr $r $fMap 'TABLE_NAME').Trim()
            $refT = (Get-RowStr $r $fMap 'REFERENCED_TABLE_NAME').Trim()
            if ([string]::IsNullOrEmpty($tn) -or [string]::IsNullOrEmpty($refT)) { continue }
            $tm = $null
            if (-not $tblMetaMap.TryGetValue($tn, [ref]$tm)) {
                $tm = New-Object AuroraAnalyzer.DbTableMetadata
                $tblMetaMap[$tn] = $tm
            }
            $fkm = New-Object AuroraAnalyzer.DbForeignKeyMetadata
            $fkm.ConstraintName       = Get-RowStr $r $fMap 'CONSTRAINT_NAME'
            $fkm.ColumnName           = Get-RowStr $r $fMap 'COLUMN_NAME'
            $fkm.ReferencedTableName  = $refT
            $fkm.ReferencedColumnName = Get-RowStr $r $fMap 'REFERENCED_COLUMN_NAME'
            $tm.DeclaredFks.Add($fkm)
        }
    }

    return @{
        ServerMeta    = $srvMeta
        TableMetaMap  = $tblMetaMap
        ColumnMetaMap = $colMetaMap
    }
}

function Invoke-AuroraDirectAnalysis {
    param(
        [Parameter(Mandatory = $true)][hashtable]$ConnConfig,
        [string]$DbLabel = '',
        [int]$MaxRowsPerTable = 0,
        [string[]]$IncludeTables = @(),
        [string[]]$ExcludeTables = @(),
        [scriptblock]$LogCallback = $null,
        [AuroraAnalyzer.OperationCancelState]$CancelState = $null,
        [scriptblock]$PumpCallback = $null
    )

    function Write-DirectLog {
        param([string]$Msg, [int]$Pct = -1)
        if ($null -ne $LogCallback) { & $LogCallback $Msg $Pct }
        else { Write-Host $Msg }
    }

    if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }

    $sessionConn = $null
    $directSession = $null

    try {
        $sessionConn = Open-AuroraReadOnlySession -ConnConfig $ConnConfig -LogCallback $LogCallback -CancelState $CancelState -PumpCallback $PumpCallback
        $client = $sessionConn.Client

        $dbMeta = Get-AuroraDirectDbMetadata -Client $client -CancelState $CancelState -LogCallback $LogCallback -PumpCallback $PumpCallback

        $tableNames = @(Get-AuroraBaseTables -Client $client -IncludeTables $IncludeTables -ExcludeTables $ExcludeTables -CancelState $CancelState -LogCallback $LogCallback -PumpCallback $PumpCallback)
        if ($tableNames.Count -eq 0) {
            throw "対象データベース '$($ConnConfig.Database)' に分析対象のBASE TABLEが見つかりませんでした。"
        }

        $metaParts = New-Object System.Collections.Generic.List[string]
        if ($dbMeta.TableMetaMap.Count -gt 0) { $metaParts.Add("物理容量/DDL $($dbMeta.TableMetaMap.Count)表") }
        if ($dbMeta.ServerMeta.HasStatus) { $metaParts.Add("サーバー状態 $($dbMeta.ServerMeta.StatusCounters.Count)指標") }
        $metaSuffix = if ($metaParts.Count -gt 0) { " / DBメタ取得: " + ($metaParts -join ', ') } else { '' }
        Write-DirectLog ("DB直接分析対象テーブル数: {0} 件 (CSV出力なし・ストリーミング解析{1})" -f $tableNames.Count, $metaSuffix) 25

        $sourceDesc = "{0}:{1}/{2}" -f $ConnConfig.ServerName, $sessionConn.EffectivePort, $ConnConfig.Database
        $directSession = New-Object AuroraAnalyzer.DirectAnalysisSession($sourceDesc)
        $directSession.AttachDbMetadata($dbMeta.ServerMeta, $dbMeta.TableMetaMap, $dbMeta.ColumnMetaMap)

        $idx = 0
        foreach ($tbl in $tableNames) {
            if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }
            $idx++
            $escapedTbl = $tbl.Replace('`', '``')
            $sql = "SELECT * FROM ``$escapedTbl``"
            if ($MaxRowsPerTable -gt 0) {
                $sql = "$sql LIMIT $MaxRowsPerTable"
            }

            $pct = 25 + [int]([Math]::Floor(30.0 * $idx / $tableNames.Count))
            Write-DirectLog ("[{0}/{1}] DB直接分析中: {2}" -f $idx, $tableNames.Count, $tbl) $pct
            $anHandle = [AuroraAnalyzer.BackgroundTaskLauncher]::StartAnalyzeMySqlTable($directSession, $client, $tbl, $ConnConfig.Database, $sql, $MaxRowsPerTable, $CancelState)
            $ta = Wait-BackgroundTask -Handle $anHandle -LogCallback $LogCallback -PumpCallback $PumpCallback -CancelState $CancelState
            Write-DirectLog ("    完了: {0} ({1:N0} 件 / {2} 列)" -f $tbl, $ta.RowCount, $ta.ColumnCount) $pct
        }
    }
    finally {
        if ($null -ne $sessionConn) {
            if ($null -ne $sessionConn.Client) { $sessionConn.Client.Dispose() }
            if ($null -ne $sessionConn.Tunnel) { Stop-SshPortForward -Tunnel $sessionConn.Tunnel }
        }
    }

    if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }
    Write-DirectLog "全テーブル横断集計・リレーション推定・警告抽出中..." 56
    $summaryHandle = [AuroraAnalyzer.BackgroundTaskLauncher]::StartCompleteDirectAnalysis($directSession, $CancelState)
    $analysis = Wait-BackgroundTask -Handle $summaryHandle -PumpCallback $PumpCallback -CancelState $CancelState
    Write-DirectLog ("DB直接解析完了: {0} テーブル / 総レコード {1:N0} 件 / 検出警告 {2} 件 (高リスク {3} 件)" -f `
        $analysis.TableCount, $analysis.TotalRows, $analysis.TotalWarningCount, $analysis.HighWarningCount) 58
    return $analysis
}

# ==============================================================================
#  Module 4A: Excel COM Report Generator (Primary — Full Native Styling & Charts)
# ==============================================================================
function Test-WarningCellText {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
    $s = $Text.Trim()
    if ($s -match '^(正常|なし|通常|0|最適|良好|均衡|変化なし|変更なし|変更検出|継続警告|-)') { return $false }
    return ($s -match '^(高|中)$|高リスク|高NULL|要注意|空テーブル|全行NULL|全NULL|偏り|孤立|大容量|外れ値|不一致|定数|件数過小|キー数少|警告|注意|減少|消失|削除|スキーマ')
}

function Get-VisualTextWidth {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return 0.0 }
    $w = 0.0
    for ($i = 0; $i -lt $Text.Length; $i++) {
        if ([int][char]$Text[$i] -gt 0x7F) { $w += 1.85 }
        else { $w += 1.0 }
    }
    return $w
}

function Get-EstimatedRowHeight {
    param(
        [string[]]$Texts,
        [double[]]$ColWidths,
        [double]$MinHeight = 22.0,
        [double]$MaxHeight = 110.0
    )
    $maxLines = 1
    for ($i = 0; $i -lt $Texts.Count; $i++) {
        $t = $Texts[$i]
        if ([string]::IsNullOrEmpty($t)) { continue }
        $cw = if ($i -lt $ColWidths.Count -and $ColWidths[$i] -gt 0) { $ColWidths[$i] } else { 20.0 }
        $usable = [Math]::Max(8.0, ($cw - 2.0))
        $lines = 0
        foreach ($seg in ($t -split "`r`n|`r|`n")) {
            $vw = Get-VisualTextWidth $seg
            $lines += [Math]::Max(1, [int][Math]::Ceiling($vw / $usable))
        }
        if ($lines -gt $maxLines) { $maxLines = $lines }
    }
    if ($maxLines -le 1) { return $MinHeight }
    return [Math]::Max($MinHeight, [Math]::Min($MaxHeight, 9.0 + ($maxLines * 14.5)))
}

function Get-RowBelowPoints {
    param($Sheet, [int]$StartRow, [double]$TargetTopPt)
    $r = [Math]::Max(1, $StartRow)
    $limit = $r + 120
    while ($r -lt $limit -and $Sheet.Rows.Item($r).Top -lt $TargetTopPt) {
        $r++
    }
    return $r
}

function Initialize-ExcelSheetStyle {
    param($Sheet)
    $Sheet.Cells.Font.Name = $script:THEME.FontName
    $Sheet.Cells.Font.Size = 9.5
    $Sheet.Cells.Font.Color = ConvertTo-ExcelOleColor $script:THEME.TextHex
}

function Format-ExcelSheetCommon {
    param(
        $Sheet,
        [int]$FreezeRow = 0,
        [string]$AutoFilterRange = '',
        [bool]$Landscape = $true
    )

    if (-not [string]::IsNullOrEmpty($AutoFilterRange)) {
        try { [void]$Sheet.Range($AutoFilterRange).AutoFilter() } catch { }
    }

    try {
        $Sheet.Activate()
        $win = $Sheet.Application.ActiveWindow
        $win.DisplayGridlines = $false
        if ($FreezeRow -gt 0) {
            $win.SplitColumn = 0
            $win.SplitRow = $FreezeRow
            $win.FreezePanes = $true
        }
    }
    catch { }

    try {
        $ps = $Sheet.PageSetup
        $ps.Orientation = if ($Landscape) { 2 } else { 1 } # xlLandscape=2, xlPortrait=1
        $ps.Zoom = $false
        $ps.FitToPagesWide = 1
        $ps.FitToPagesTall = $false
        $ps.LeftMargin = 28
        $ps.RightMargin = 28
        $ps.TopMargin = 36
        $ps.BottomMargin = 36
    }
    catch { }
}

function Set-SheetTitleBanner {
    param(
        $Sheet,
        [string]$Title,
        [string]$Subtitle,
        [int]$EndCol
    )

    $navy = ConvertTo-ExcelOleColor $script:THEME.AccentHex
    $white = ConvertTo-ExcelOleColor $script:THEME.WhiteHex
    $textColor = ConvertTo-ExcelOleColor $script:THEME.TextHex
    $subBg = ConvertTo-ExcelOleColor $script:THEME.AccentLightHex

    $r1 = $Sheet.Range($Sheet.Cells.Item(1, 1), $Sheet.Cells.Item(1, $EndCol))
    $r1.Merge()
    [AuroraAnalyzer.ComHelper]::SetValue2($Sheet.Cells.Item(1, 1), $Title)
    $r1.Font.Name = $script:THEME.FontName
    $r1.Font.Size = 13.5
    $r1.Font.Bold = $true
    $r1.Font.Color = $white
    $r1.Interior.Color = $navy
    $r1.HorizontalAlignment = -4131 # xlLeft
    $r1.VerticalAlignment = -4108   # xlVAlignCenter
    $r1.IndentLevel = 1
    $Sheet.Rows.Item(1).RowHeight = 32

    $r2 = $Sheet.Range($Sheet.Cells.Item(2, 1), $Sheet.Cells.Item(2, $EndCol))
    $r2.Merge()
    [AuroraAnalyzer.ComHelper]::SetValue2($Sheet.Cells.Item(2, 1), $Subtitle)
    $r2.Font.Name = $script:THEME.FontName
    $r2.Font.Size = 9
    $r2.Font.Bold = $false
    $r2.Font.Color = $textColor
    $r2.Interior.Color = $subBg
    $r2.HorizontalAlignment = -4131
    $r2.VerticalAlignment = -4108
    $r2.IndentLevel = 1
    $r2.Borders.Item(9).LineStyle = 1 # xlEdgeBottom
    $r2.Borders.Item(9).Color = $navy
    $r2.Borders.Item(9).Weight = 2    # xlThin
    $Sheet.Rows.Item(2).RowHeight = 22
    $Sheet.Rows.Item(3).RowHeight = 10
}

function Set-SectionHeader {
    param(
        $Sheet,
        [int]$Row,
        [int]$StartCol,
        [int]$EndCol,
        [string]$Text
    )

    $navy = ConvertTo-ExcelOleColor $script:THEME.AccentHex
    $bg   = ConvertTo-ExcelOleColor $script:THEME.AccentLightHex

    $rng = $Sheet.Range($Sheet.Cells.Item($Row, $StartCol), $Sheet.Cells.Item($Row, $EndCol))
    $rng.Merge()
    [AuroraAnalyzer.ComHelper]::SetValue2($Sheet.Cells.Item($Row, $StartCol), $Text)
    $rng.Font.Name = $script:THEME.FontName
    $rng.Font.Size = 10.5
    $rng.Font.Bold = $true
    $rng.Font.Color = $navy
    $rng.Interior.Color = $bg
    $rng.HorizontalAlignment = -4131
    $rng.VerticalAlignment = -4108
    $rng.IndentLevel = 1
    $rng.Borders.Item(7).LineStyle = 1 # xlEdgeLeft accent bar
    $rng.Borders.Item(7).Color = $navy
    $rng.Borders.Item(7).Weight = 3    # xlMedium
    $rng.Borders.Item(9).LineStyle = 1 # xlEdgeBottom
    $rng.Borders.Item(9).Color = $navy
    $rng.Borders.Item(9).Weight = 2    # xlThin
    $Sheet.Rows.Item($Row).RowHeight = 25
}

function Write-StyledTable {
    param(
        $Sheet,
        [int]$StartRow,
        [int]$StartCol,
        [string[]]$Headers,
        [object[,]]$DataMatrix,
        [int]$RowCount,
        [hashtable]$ColumnFormats = @{},
        [int[]]$CenterCols = @(),
        [int[]]$RightCols = @(),
        [int[]]$WrapCols = @(),
        [int[]]$WarnCheckCols = @(),
        [scriptblock]$CheckCancellation = $null
    )

    $colCount = $Headers.Count
    $navy = ConvertTo-ExcelOleColor $script:THEME.AccentHex
    $white = ConvertTo-ExcelOleColor $script:THEME.WhiteHex
    $textColor = ConvertTo-ExcelOleColor $script:THEME.TextHex
    $zebra = ConvertTo-ExcelOleColor $script:THEME.ZebraHex
    $borderColor = ConvertTo-ExcelOleColor $script:THEME.BorderHex
    $darkBorderColor = ConvertTo-ExcelOleColor $script:THEME.BorderDarkHex
    $hdrBorderColor = ConvertTo-ExcelOleColor $script:THEME.HeaderBorderHex
    $warnColor = ConvertTo-ExcelOleColor $script:THEME.WarnHex
    $warnBg = ConvertTo-ExcelOleColor $script:THEME.WarnBgHex

    # Header row
    $hdrMatrix = New-Object 'object[,]' 1, $colCount
    for ($c = 0; $c -lt $colCount; $c++) { $hdrMatrix[0, $c] = $Headers[$c] }
    $hdrRng = $Sheet.Range($Sheet.Cells.Item($StartRow, $StartCol), $Sheet.Cells.Item($StartRow, $StartCol + $colCount - 1))
    $hdrRng.NumberFormat = '@'
    [AuroraAnalyzer.ComHelper]::SetValue2($hdrRng, $hdrMatrix)
    $hdrRng.Font.Name = $script:THEME.FontName
    $hdrRng.Font.Size = 9.5
    $hdrRng.Font.Bold = $true
    $hdrRng.Font.Color = $white
    $hdrRng.Interior.Color = $navy
    $hdrRng.HorizontalAlignment = -4108 # xlCenter
    $hdrRng.VerticalAlignment = -4108
    $hdrRng.WrapText = $true
    foreach ($edge in @(7, 8, 9, 10, 11)) {
        try {
            $hb = $hdrRng.Borders.Item($edge)
            $hb.LineStyle = 1
            $hb.Weight = 2
            $hb.Color = if ($edge -eq 11) { $hdrBorderColor } else { $navy }
        }
        catch { }
    }
    $Sheet.Rows.Item($StartRow).RowHeight = 28

    if ($RowCount -le 0) {
        $emptyRng = $Sheet.Range($Sheet.Cells.Item($StartRow + 1, $StartCol), $Sheet.Cells.Item($StartRow + 1, $StartCol + $colCount - 1))
        $emptyRng.Merge()
        [AuroraAnalyzer.ComHelper]::SetValue2($Sheet.Cells.Item($StartRow + 1, $StartCol), '該当データなし')
        $emptyRng.Font.Name = $script:THEME.FontName
        $emptyRng.Font.Size = 9.5
        $emptyRng.Font.Color = ConvertTo-ExcelOleColor $script:THEME.MutedHex
        $emptyRng.Interior.Color = $white
        $emptyRng.HorizontalAlignment = -4108
        $emptyRng.VerticalAlignment = -4108
        foreach ($edge in @(7, 8, 9, 10)) {
            try { $eb = $emptyRng.Borders.Item($edge); $eb.LineStyle = 1; $eb.Weight = 2; $eb.Color = $borderColor } catch { }
        }
        $Sheet.Rows.Item($StartRow + 1).RowHeight = 24
        return ($StartRow + 1)
    }

    $dataStart = $StartRow + 1
    $dataEnd = $StartRow + $RowCount
    $dataRng = $Sheet.Range($Sheet.Cells.Item($dataStart, $StartCol), $Sheet.Cells.Item($dataEnd, $StartCol + $colCount - 1))

    # Apply column number formats BEFORE SetValue2 so '@' text columns (times, dates, percentages) are never auto-coerced by Excel
    for ($c = 1; $c -le $colCount; $c++) {
        if ($ColumnFormats.ContainsKey($c)) {
            $colRng = $Sheet.Range($Sheet.Cells.Item($dataStart, $StartCol + $c - 1), $Sheet.Cells.Item($dataEnd, $StartCol + $c - 1))
            $colRng.NumberFormat = $ColumnFormats[$c]
        }
    }

    [AuroraAnalyzer.ComHelper]::SetValue2($dataRng, $DataMatrix)
    $dataRng.Font.Name = $script:THEME.FontName
    $dataRng.Font.Size = 9.5
    $dataRng.Font.Color = $textColor
    $dataRng.Interior.Color = $white
    $dataRng.VerticalAlignment = -4108
    $dataRng.HorizontalAlignment = -4131 # xlLeft

    # Borders across data range
    foreach ($edge in @(7, 8, 9, 10, 11, 12)) {
        try {
            $b = $dataRng.Borders.Item($edge)
            $b.LineStyle = 1
            $b.Weight = 2
            $b.Color = if ($edge -eq 9) { $darkBorderColor } else { $borderColor }
        }
        catch { }
    }

    # Apply column alignments & wrap
    $colWidthMap = @{}
    for ($c = 1; $c -le $colCount; $c++) {
        $absCol = $StartCol + $c - 1
        $colRng = $Sheet.Range($Sheet.Cells.Item($dataStart, $absCol), $Sheet.Cells.Item($dataEnd, $absCol))
        if ($CenterCols -contains $c) { $colRng.HorizontalAlignment = -4108 }
        elseif ($RightCols -contains $c) { $colRng.HorizontalAlignment = -4152 } # xlRight
        if ($WrapCols -contains $c) {
            $colRng.WrapText = $true
            try { $colWidthMap[$c] = [double]$Sheet.Columns.Item($absCol).ColumnWidth } catch { $colWidthMap[$c] = 24.0 }
        }
    }

    # Zebra striping + dynamic row height + Burgundy warning cell highlights (strictly monotone neutral background)
    for ($r = 0; $r -lt $RowCount; $r++) {
        if (($r % 16) -eq 0 -and $null -ne $CheckCancellation) { & $CheckCancellation }
        $excelRow = $dataStart + $r
        if ($WrapCols.Count -gt 0) {
            $wTexts = New-Object System.Collections.Generic.List[string]
            $wWidths = New-Object System.Collections.Generic.List[double]
            foreach ($wc in $WrapCols) {
                $colIdx = $wc - 1
                $wVal = [string]$DataMatrix[$r, $colIdx]
                $wTexts.Add($wVal)
                $wWidths.Add([double]$colWidthMap[$wc])
            }
            $Sheet.Rows.Item($excelRow).RowHeight = Get-EstimatedRowHeight -Texts $wTexts.ToArray() -ColWidths $wWidths.ToArray() -MinHeight 22.0 -MaxHeight 110.0
        }
        else {
            $Sheet.Rows.Item($excelRow).RowHeight = 22
        }

        if (($r % 2) -eq 1) {
            $rowRng = $Sheet.Range($Sheet.Cells.Item($excelRow, $StartCol), $Sheet.Cells.Item($excelRow, $StartCol + $colCount - 1))
            $rowRng.Interior.Color = $zebra
        }
        foreach ($wc in $WarnCheckCols) {
            $colIdx = $wc - 1
            $valStr = [string]$DataMatrix[$r, $colIdx]
            if (Test-WarningCellText -Text $valStr) {
                $cell = $Sheet.Cells.Item($excelRow, $StartCol + $wc - 1)
                $cell.Font.Color = $warnColor
                $cell.Font.Bold = $true
                $cell.Interior.Color = $warnBg
            }
        }
    }

    return $dataEnd
}

function Set-ExcelCellValue {
    param($Cell, $Value)
    # Invoke IDispatch::Value2 directly to bypass PowerShell 5.1 COM DLR type caching
    if ($null -eq $Value) {
        [AuroraAnalyzer.ComHelper]::SetValue2($Cell, '')
    }
    elseif ($Value -is [string]) {
        [AuroraAnalyzer.ComHelper]::SetValue2($Cell, [string]$Value)
    }
    else {
        [AuroraAnalyzer.ComHelper]::SetValue2($Cell, [double]$Value)
    }
}

function Set-ColumnWidths {
    param($Sheet, [double[]]$Widths)
    for ($i = 0; $i -lt $Widths.Count; $i++) {
        $Sheet.Columns.Item($i + 1).ColumnWidth = $Widths[$i]
    }
}

function ConvertTo-ExcelColLetter {
    param([Parameter(Mandatory = $true)][int]$ColIndex1Based)
    $dividend = $ColIndex1Based
    $colName = ''
    while ($dividend -gt 0) {
        $modulo = ($dividend - 1) % 26
        $colName = [char](65 + $modulo) + $colName
        $dividend = [int][Math]::Floor(($dividend - $modulo) / 26)
    }
    return $colName
}

function Get-RoleSummaryItems {
    param($Analysis)
    $groups = @{}
    foreach ($t in $Analysis.Tables) {
        $rName = if ([string]::IsNullOrWhiteSpace($t.InferredRole)) { 'その他' } else { $t.InferredRole }
        if (-not $groups.ContainsKey($rName)) {
            $groups[$rName] = [PSCustomObject]@{
                Role           = $rName
                TableCount     = 0
                TotalRows      = 0L
                TotalCols      = 0
                TotalFileBytes = 0L
                TotalInnoBytes = 0L
                WarningCount   = 0
                Examples       = (New-Object System.Collections.Generic.List[string])
            }
        }
        $g = $groups[$rName]
        $g.TableCount++
        $g.TotalRows += $t.RowCount
        $g.TotalCols += $t.ColumnCount
        $g.TotalFileBytes += $t.FileSizeBytes
        $g.TotalInnoBytes += $t.EstInnoDbBytes
        $g.WarningCount += $t.WarningCount
        if ($g.Examples.Count -lt 4) { $g.Examples.Add($t.TableName) }
    }
    return @($groups.Values | Sort-Object @{ Expression = { $_.TotalRows }; Descending = $true }, @{ Expression = { $_.TableCount }; Descending = $true }, @{ Expression = { $_.Role }; Ascending = $true })
}

function Get-TypeSummaryItems {
    param($AllCols)
    $groups = @{}
    foreach ($c in $AllCols) {
        $tName = if ([string]::IsNullOrWhiteSpace($c.InferredType)) { 'その他' } else { $c.InferredType }
        if (-not $groups.ContainsKey($tName)) {
            $groups[$tName] = [PSCustomObject]@{
                DataType     = $tName
                ColCount     = 0
                PkCount      = 0
                CatCount     = 0
                NullRateSum  = 0.0
                CardRateSum  = 0.0
                AvgByteSum   = 0.0
                NonEmptyCols = 0
                MaxByteLen   = 0
                HighSkewCnt  = 0
                AllNullCnt   = 0
                Examples     = (New-Object System.Collections.Generic.List[string])
            }
        }
        $g = $groups[$tName]
        $g.ColCount++
        if ($c.IsCandidatePk) { $g.PkCount++ }
        if ($c.IsCategorical) { $g.CatCount++ }
        $g.NullRateSum += $c.NullRate
        $g.CardRateSum += $c.CardinalityRate
        if ($c.NonNullCount -gt 0) {
            $g.AvgByteSum += $c.AvgByteLen
            $g.NonEmptyCols++
        }
        if ($c.MaxByteLen -gt $g.MaxByteLen) { $g.MaxByteLen = $c.MaxByteLen }
        if ($c.SkewRisk -like '高リスク*') { $g.HighSkewCnt++ }
        if ($c.TotalRows -gt 0 -and $c.NullRate -ge 0.999) { $g.AllNullCnt++ }
        if ($g.Examples.Count -lt 3) { $g.Examples.Add("$($c.TableName).$($c.ColumnName)") }
    }
    return @($groups.Values | Sort-Object @{ Expression = { $_.ColCount }; Descending = $true }, @{ Expression = { $_.DataType }; Ascending = $true })
}

function Add-StyledExcelChart {
    param(
        $Sheet,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height,
        [int]$ChartType,          # 51=xlColumnClustered, 57=xlBarClustered, -4120=xlDoughnut, 4=xlLineMarkers
        [string]$Title,
        [string]$CategoryAddress,
        [string[]]$ValueAddresses,
        [string[]]$SeriesNames = @(),
        [string[]]$SeriesColors = @(),
        [bool]$ShowLegend = $false,
        [bool]$ShowDataLabels = $false,
        [bool]$ReverseCategories = $false,
        [bool]$ColorNegativeBurgundy = $false,
        [int]$GapWidth = 75
    )

    $co = $Sheet.ChartObjects().Add($Left, $Top, $Width, $Height)
    $ch = $co.Chart
    $ch.ChartType = $ChartType

    # Clear default series if any
    while ($ch.SeriesCollection().Count -gt 0) {
        $ch.SeriesCollection().Item(1).Delete()
    }

    $palette = @(if (@($SeriesColors).Count -gt 0) { $SeriesColors } else { $script:THEME.ChartPalette })
    $textOle = ConvertTo-ExcelOleColor $script:THEME.TextHex
    $navyOle = ConvertTo-ExcelOleColor $script:THEME.AccentHex
    $warnOle = ConvertTo-ExcelOleColor $script:THEME.WarnHex
    $borderOle = ConvertTo-ExcelOleColor $script:THEME.BorderHex
    $gridOle = ConvertTo-ExcelOleColor $script:THEME.GridlineHex
    $whiteOle = ConvertTo-ExcelOleColor $script:THEME.WhiteHex

    $catRng = $Sheet.Range($CategoryAddress)

    for ($i = 0; $i -lt $ValueAddresses.Count; $i++) {
        $valRng = $Sheet.Range($ValueAddresses[$i])
        $ser = $ch.SeriesCollection().NewSeries()
        $ser.XValues = $catRng
        $ser.Values = $valRng
        if ($i -lt $SeriesNames.Count -and -not [string]::IsNullOrEmpty($SeriesNames[$i])) {
            $ser.Name = $SeriesNames[$i]
        }

        $hexCol = $palette[$i % $palette.Count]
        $oleCol = ConvertTo-ExcelOleColor $hexCol
        try {
            if ($ChartType -eq -4120) {
                # Doughnut: color each slice in Navy + Monotone Gray palette with clean white separators
                $pts = $ser.Points()
                for ($p = 1; $p -le $pts.Count; $p++) {
                    $pHex = $script:THEME.ChartPalette[($p - 1) % $script:THEME.ChartPalette.Count]
                    $pts.Item($p).Format.Fill.ForeColor.RGB = ConvertTo-ExcelOleColor $pHex
                    $pts.Item($p).Format.Line.ForeColor.RGB = $whiteOle
                    $pts.Item($p).Format.Line.Weight = 1.5
                }
            }
            elseif ($ChartType -eq 4) {
                $ser.Format.Line.ForeColor.RGB = $oleCol
                $ser.Format.Line.Weight = 2.25
                $ser.MarkerBackgroundColor = $oleCol
                $ser.MarkerForegroundColor = $oleCol
            }
            else {
                $ser.Format.Fill.ForeColor.RGB = $oleCol
                $ser.Format.Line.Visible = 0
                if ($ColorNegativeBurgundy) {
                    $pts = $ser.Points()
                    $rawVals = $valRng.Value2
                    for ($p = 1; $p -le $pts.Count; $p++) {
                        $pv = if ($rawVals -is [System.Array]) { $rawVals[$p, 1] } else { $rawVals }
                        $numVal = 0.0
                        if ($null -ne $pv -and [double]::TryParse([string]$pv, [ref]$numVal) -and $numVal -lt -1e-9) {
                            $pts.Item($p).Format.Fill.ForeColor.RGB = $warnOle
                        }
                    }
                }
            }
        }
        catch { }

        if ($ShowDataLabels) {
            try {
                $ser.HasDataLabels = $true
                $dl = $ser.DataLabels()
                $dl.Font.Name = $script:THEME.FontName
                $dl.Font.Size = 8
                $dl.Font.Color = $textOle
                if ($ChartType -eq -4120) {
                    $dl.ShowPercentage = $true
                    $dl.ShowValue = $false
                    $dl.ShowCategoryName = $false
                    $dl.Font.Bold = $true
                    # White pill behind doughnut percentages guarantees high contrast over dark Navy slices
                    $dl.Format.Fill.Visible = -1
                    $dl.Format.Fill.ForeColor.RGB = $whiteOle
                    $dl.Format.Line.Visible = -1
                    $dl.Format.Line.ForeColor.RGB = $borderOle
                    $dl.Format.Line.Weight = 0.5
                }
                elseif ($ChartType -eq 51 -or $ChartType -eq 57) {
                    try { $dl.Position = 2 } catch { } # xlLabelPositionOutsideEnd
                }
            }
            catch { }
        }
    }

    # Chart Area & Plot Area Monotone Styling
    try {
        $ch.ChartArea.Format.Fill.ForeColor.RGB = $whiteOle
        $ch.ChartArea.Format.Line.ForeColor.RGB = $borderOle
        $ch.ChartArea.Format.Line.Weight = 0.75
        $ch.ChartArea.Font.Name = $script:THEME.FontName
        $ch.ChartArea.Font.Color = $textOle
        $ch.PlotArea.Format.Fill.ForeColor.RGB = $whiteOle
        $ch.PlotArea.Format.Line.Visible = 0
    }
    catch { }

    # Title
    $ch.HasTitle = $true
    $ch.ChartTitle.Text = $Title
    $ch.ChartTitle.Font.Name = $script:THEME.FontName
    $ch.ChartTitle.Font.Size = 10.5
    $ch.ChartTitle.Font.Bold = $true
    $ch.ChartTitle.Font.Color = $navyOle

    # Legend
    $ch.HasLegend = $ShowLegend
    if ($ShowLegend) {
        try {
            $ch.Legend.Position = -4107 # xlLegendPositionBottom
            $ch.Legend.Font.Name = $script:THEME.FontName
            $ch.Legend.Font.Size = 8
            $ch.Legend.Font.Color = $textOle
        }
        catch { }
    }

    # Axes styling (not applicable to Doughnut)
    if ($ChartType -ne -4120) {
        try {
            $catAxis = $ch.Axes(1) # xlCategory
            $catAxis.TickLabels.Font.Name = $script:THEME.FontName
            $catAxis.TickLabels.Font.Size = 8.5
            $catAxis.TickLabels.Font.Color = $textOle
            $catAxis.Format.Line.ForeColor.RGB = $borderOle
            $catAxis.TickLabelPosition = -4134 # xlTickLabelPositionLow (prevents overlap with negative bars)
            if ($ReverseCategories) {
                $catAxis.ReversePlotOrder = $true
            }
        }
        catch { }

        try {
            $valAxis = $ch.Axes(2) # xlValue
            $valAxis.TickLabels.Font.Name = $script:THEME.FontName
            $valAxis.TickLabels.Font.Size = 8.5
            $valAxis.TickLabels.Font.Color = $textOle
            try {
                $curAxisFmt = [string]$valAxis.TickLabels.NumberFormat
                if ($curAxisFmt -eq '0.0%') {
                    $valAxis.TickLabels.NumberFormatLinked = $false
                    $valAxis.TickLabels.NumberFormat = '0%'
                }
            }
            catch { }
            $valAxis.HasMajorGridlines = $true
            $valAxis.MajorGridlines.Format.Line.ForeColor.RGB = $gridOle
            $valAxis.Format.Line.Visible = 0
        }
        catch { }

        try {
            if ($ChartType -eq 51 -or $ChartType -eq 57) {
                $ch.ChartGroups(1).GapWidth = $GapWidth
            }
        }
        catch { }
    }

    return $co
}

function Export-AnalysisToExcelCom {
    param(
        [Parameter(Mandatory = $true)]$Analysis,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [string]$DbLabel = '',
        [switch]$DetailedAnalysis,
        [scriptblock]$LogCallback = $null,
        [AuroraAnalyzer.OperationCancelState]$CancelState = $null,
        [scriptblock]$PumpCallback = $null
    )

    function Write-XlsLog {
        param([string]$Msg, [int]$Pct = -1)
        if ($null -ne $LogCallback) { & $LogCallback $Msg $Pct }
        if ($Pct -lt 100 -and $null -ne $CancelState) { $CancelState.ThrowIfCancelled() }
    }

    if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }

    $OutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
    $stagingPath = New-AnalysisStagingPath $OutputPath
    $checkReport = {
        if ($null -ne $PumpCallback) { & $PumpCallback }
        if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }
    }

    $excel = $null
    $excelPid = 0
    $wb = $null
    $sheetsToRelease = New-Object System.Collections.Generic.List[object]

    try {
        $modeSheetText = if ($DetailedAnalysis) { '詳細分析8シート' } else { '標準3シート' }
        Write-XlsLog "Excelレポート生成開始 (Excel COM / $modeSheetText / Meiryo UI モノトーン仕様)..." 60
        try { $excel = New-Object -ComObject Excel.Application }
        catch [System.Runtime.InteropServices.COMException] {
            throw [AuroraAnalyzer.ExcelActivationException]::new($_.Exception)
        }
        try { $excelPid = [AuroraAnalyzer.ComHelper]::GetWindowProcessId([int]$excel.Hwnd) } catch { $excelPid = 0 }
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false
        try { $excel.EnableEvents = $false } catch { }
        $targetSheetCount = if ($DetailedAnalysis) { 8 } else { 3 }
        $excel.SheetsInNewWorkbook = $targetSheetCount
        $wb = $excel.Workbooks.Add()
        while ($wb.Worksheets.Count -lt $targetSheetCount) {
            [void]$wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
        }
        while ($wb.Worksheets.Count -gt $targetSheetCount) {
            $wb.Worksheets.Item($wb.Worksheets.Count).Delete()
        }
        $excel.Calculation = -4135 # xlCalculationManual

        $wsSummary = $wb.Worksheets.Item(1); $wsSummary.Name = '01_サマリー'
        $wsTables  = $wb.Worksheets.Item(2); $wsTables.Name  = '02_テーブル一覧'
        $wsCols    = $wb.Worksheets.Item(3); $wsCols.Name    = '03_カラム分析'
        $initSheets = @($wsSummary, $wsTables, $wsCols)
        $wsDist = $null; $wsNum = $null; $wsTime = $null; $wsRel = $null; $wsGuide = $null
        if ($DetailedAnalysis) {
            $wsDist  = $wb.Worksheets.Item(4); $wsDist.Name  = '04_分布・割合分析'
            $wsNum   = $wb.Worksheets.Item(5); $wsNum.Name   = '05_数値統計・ヒストグラム'
            $wsTime  = $wb.Worksheets.Item(6); $wsTime.Name  = '06_時系列分析'
            $wsRel   = $wb.Worksheets.Item(7); $wsRel.Name   = '07_リレーション・ファンアウト'
            $wsGuide = $wb.Worksheets.Item(8); $wsGuide.Name = '08_性能試験ガイド・警告'
            $initSheets = @($wsSummary, $wsTables, $wsCols, $wsDist, $wsNum, $wsTime, $wsRel, $wsGuide)
        }

        foreach ($s in $initSheets) {
            $sheetsToRelease.Add($s)
            Initialize-ExcelSheetStyle -Sheet $s
        }

        $isDirectDb = [bool]$Analysis.IsDirectDb
        $dbLabelPart = if ([string]::IsNullOrWhiteSpace($DbLabel)) { '' } else { "  |  A5M2登録DB: $DbLabel" }
        $subInfo = if ($isDirectDb) {
            "分析日時: {0}  |  接続先DB: {1}{2}  |  取得方式: DB直接分析(CSV未出力)  |  出力構成: {3}  |  ターゲット: Aurora Serverless v2 (MySQL 8.4.7)" -f `
                $Analysis.AnalyzedAt.ToString('yyyy/MM/dd HH:mm:ss'), $Analysis.InputFolder, $dbLabelPart, $modeSheetText
        }
        else {
            "分析日時: {0}  |  対象フォルダ: {1}{2}  |  出力構成: {3}  |  ターゲット: Aurora Serverless v2 (MySQL 8.4.7)" -f `
                $Analysis.AnalyzedAt.ToString('yyyy/MM/dd HH:mm:ss'), $Analysis.InputFolder, $dbLabelPart, $modeSheetText
        }
        $sizeKpiLabel = if ($isDirectDb) { 'CSV換算サイズ (MB)' } else { '総CSVサイズ (MB)' }
        $sizeColHeader = if ($isDirectDb) { 'CSV換算サイズ(KB)' } else { 'CSVサイズ(KB)' }
        $srcColHeader = if ($isDirectDb) { '取得元 (DB.テーブル)' } else { 'CSVファイル名' }
        $tblSheetTitle = if ($isDirectDb) { 'テーブル別詳細一覧 (DB直接分析プロファイル・CSV未出力)' } else { 'テーブル別詳細一覧 (全CSVファイルプロファイル)' }

        $navyOle       = ConvertTo-ExcelOleColor $script:THEME.AccentHex
        $textOle       = ConvertTo-ExcelOleColor $script:THEME.TextHex
        $mutedOle      = ConvertTo-ExcelOleColor $script:THEME.MutedHex
        $warnOle       = ConvertTo-ExcelOleColor $script:THEME.WarnHex
        $warnBgOle     = ConvertTo-ExcelOleColor $script:THEME.WarnBgHex
        $cardBgOle     = ConvertTo-ExcelOleColor $script:THEME.CardBgHex
        $borderOle     = ConvertTo-ExcelOleColor $script:THEME.BorderHex
        $darkBorderOle = ConvertTo-ExcelOleColor $script:THEME.BorderDarkHex
        $zebraOle      = ConvertTo-ExcelOleColor $script:THEME.ZebraHex
        $totalBgOle    = ConvertTo-ExcelOleColor $script:THEME.TotalBgHex
        $whiteOle      = ConvertTo-ExcelOleColor $script:THEME.WhiteHex

        $allCols = New-Object System.Collections.Generic.List[object]
        foreach ($t in $Analysis.Tables) {
            foreach ($c in $t.Columns) { $allCols.Add($c) }
        }
        $roleSummary = @(Get-RoleSummaryItems -Analysis $Analysis)
        $typeSummary = @(Get-TypeSummaryItems -AllCols $allCols)

        function Write-AnalysisTotalRow {
            param($Sheet, [int]$Row, [int]$StartCol, [object[]]$Values, [hashtable]$Formats = @{}, [int[]]$CenterCols = @(), [int[]]$RightCols = @(), [int[]]$WarnPositiveCols = @())
            $cnt = $Values.Count
            $rng = $Sheet.Range($Sheet.Cells.Item($Row, $StartCol), $Sheet.Cells.Item($Row, $StartCol + $cnt - 1))
            $rng.Font.Name = $script:THEME.FontName
            $rng.Font.Size = 9.5
            $rng.Font.Bold = $true
            $rng.Font.Color = $navyOle
            $rng.Interior.Color = $totalBgOle
            $rng.VerticalAlignment = -4108
            $rng.HorizontalAlignment = -4131
            foreach ($edge in @(7, 8, 9, 10, 11)) {
                try {
                    $b = $rng.Borders.Item($edge)
                    $b.LineStyle = 1
                    $b.Weight = 2
                    $b.Color = if ($edge -eq 8 -or $edge -eq 9) { $navyOle } else { $borderOle }
                }
                catch { }
            }
            for ($c = 0; $c -lt $cnt; $c++) {
                $col1 = $c + 1
                $cell = $Sheet.Cells.Item($Row, $StartCol + $c)
                if ($Formats.ContainsKey($col1)) { $cell.NumberFormat = $Formats[$col1] }
                Set-ExcelCellValue -Cell $cell -Value $Values[$c]
                if ($CenterCols -contains $col1) { $cell.HorizontalAlignment = -4108 }
                elseif ($RightCols -contains $col1) { $cell.HorizontalAlignment = -4152 }
                if (($WarnPositiveCols -contains $col1) -and ($Values[$c] -is [double] -or $Values[$c] -is [int] -or $Values[$c] -is [long]) -and [double]$Values[$c] -gt 0) {
                    $cell.Font.Color = $warnOle
                }
            }
            $Sheet.Rows.Item($Row).RowHeight = 25
        }

        function Write-KpiCardBand {
            param($Sheet, [int]$LabelRow, [object[]]$Kpis)
            $Sheet.Rows.Item($LabelRow).RowHeight = 20
            $Sheet.Rows.Item($LabelRow + 1).RowHeight = 30
            foreach ($kpi in $Kpis) {
                $lblRng = $Sheet.Range($Sheet.Cells.Item($LabelRow, $kpi.Col1), $Sheet.Cells.Item($LabelRow, $kpi.Col2))
                $lblRng.Merge()
                [AuroraAnalyzer.ComHelper]::SetValue2($Sheet.Cells.Item($LabelRow, $kpi.Col1), $kpi.Label)
                $lblRng.Font.Name = $script:THEME.FontName
                $lblRng.Font.Size = 8.5
                $lblRng.Font.Bold = $true
                $lblRng.Font.Color = if ($kpi.Warn) { $warnOle } else { $mutedOle }
                $lblRng.Interior.Color = $cardBgOle
                $lblRng.HorizontalAlignment = -4108
                $lblRng.VerticalAlignment = -4108

                $valRng = $Sheet.Range($Sheet.Cells.Item($LabelRow + 1, $kpi.Col1), $Sheet.Cells.Item($LabelRow + 1, $kpi.Col2))
                $valRng.Merge()
                $valRng.NumberFormat = $kpi.Fmt
                Set-ExcelCellValue -Cell $Sheet.Cells.Item($LabelRow + 1, $kpi.Col1) -Value $kpi.Value
                $valRng.Font.Name = $script:THEME.FontName
                $valRng.Font.Size = if ($kpi.ContainsKey('FontSize')) { $kpi.FontSize } else { 13.0 }
                $valRng.Font.Bold = $true
                $valRng.Font.Color = if ($kpi.Warn) { $warnOle } else { $navyOle }
                $valRng.Interior.Color = $cardBgOle
                $valRng.HorizontalAlignment = -4108
                $valRng.VerticalAlignment = -4108

                $box = $Sheet.Range($Sheet.Cells.Item($LabelRow, $kpi.Col1), $Sheet.Cells.Item($LabelRow + 1, $kpi.Col2))
                $box.Borders.Item(8).LineStyle = 1; $box.Borders.Item(8).Color = if ($kpi.Warn) { $warnOle } else { $navyOle }; $box.Borders.Item(8).Weight = 3
                foreach ($e in @(7, 9, 10)) { $box.Borders.Item($e).LineStyle = 1; $box.Borders.Item($e).Color = $borderOle; $box.Borders.Item($e).Weight = 2 }
            }
        }

        # ======================================================================
        #  SHEET 1: 01_サマリー (Comprehensive KPI Dashboard, Breakdowns & Charts)
        # ======================================================================
        Write-XlsLog "シート作成中: 01_サマリー..." 64
        Set-ColumnWidths -Sheet $wsSummary -Widths @(6, 24, 17, 13, 11, 8, 15, 11, 14, 11, 8, 32)
        Set-SheetTitleBanner -Sheet $wsSummary -Title 'Aurora Serverless v2 (MySQL 8.4.7) 性能試験用データ分析レポート — サマリー' -Subtitle $subInfo -EndCol 12

        # KPI Cards Row 1 (Primary Volume & Schema KPIs: Rows 4-5)
        $kpiRow1 = @(
            @{ Col1 = 1;  Col2 = 2;  Label = '総テーブル数';           Value = [double]$Analysis.TableCount;                                Fmt = '#,##0" テーブル"'; Warn = $false; FontSize = 13.5 },
            @{ Col1 = 3;  Col2 = 4;  Label = '総レコード数';           Value = [double]$Analysis.TotalRows;                                 Fmt = '#,##0" 件"';       Warn = ($Analysis.TotalRows -eq 0); FontSize = 13.5 },
            @{ Col1 = 5;  Col2 = 6;  Label = $sizeKpiLabel;            Value = [Math]::Round($Analysis.TotalFileBytes / 1MB, 2);            Fmt = '#,##0.00" MB"';    Warn = $false; FontSize = 13.5 },
            @{ Col1 = 7;  Col2 = 8;  Label = '推定InnoDB容量 (MB)';    Value = [Math]::Round($Analysis.TotalEstInnoDbBytes / 1MB, 2);       Fmt = '#,##0.00" MB"';    Warn = $false; FontSize = 13.5 },
            @{ Col1 = 9;  Col2 = 10; Label = '総カラム数 / PK候補';    Value = ("{0} 列 / PK {1} 列" -f $Analysis.TotalColumns, $Analysis.CandidatePkColumnCount); Fmt = '@'; Warn = $false; FontSize = 12.0 },
            @{ Col1 = 11; Col2 = 12; Label = '検出警告 (高/中/低)';    Value = ("高 {0} / 中 {1} / 低 {2} (計 {3})" -f $Analysis.HighWarningCount, $Analysis.MediumWarningCount, $Analysis.LowWarningCount, $Analysis.TotalWarningCount); Fmt = '@'; Warn = ($Analysis.TotalWarningCount -gt 0); FontSize = 11.5 }
        )
        Write-KpiCardBand -Sheet $wsSummary -LabelRow 4 -Kpis $kpiRow1
        $wsSummary.Rows.Item(6).RowHeight = 8

        # KPI Cards Row 2 (Quality, Skew, Relation & ACU KPIs: Rows 7-8)
        $kpiRow2 = @(
            @{ Col1 = 1;  Col2 = 2;  Label = '全体セルNULL率';            Value = [double]$Analysis.OverallAvgNullRate; Fmt = '0.0%'; Warn = ($Analysis.OverallAvgNullRate -ge 0.30); FontSize = 13.0 },
            @{ Col1 = 3;  Col2 = 4;  Label = '平均行長 / 最大行長';       Value = ("{0:#,##0.0} B / {1:#,##0} B" -f $Analysis.OverallAvgRowBytes, $Analysis.OverallMaxRowBytes); Fmt = '@'; Warn = $false; FontSize = 11.0 },
            @{ Col1 = 5;  Col2 = 6;  Label = '偏り高リスク / 大容量列';   Value = ("偏り {0} 列 / 大容量 {1} 列" -f $Analysis.HighSkewColumnCount, $Analysis.LargeColumnCount); Fmt = '@'; Warn = ($Analysis.HighSkewColumnCount -gt 0 -or $Analysis.LargeColumnCount -gt 0); FontSize = 11.0 },
            @{ Col1 = 7;  Col2 = 8;  Label = '検出リレーション / 孤立行'; Value = ("{0} 件 / 孤立 {1:#,##0} 行" -f $Analysis.Relationships.Count, $Analysis.TotalOrphanRows); Fmt = '@'; Warn = ($Analysis.TotalOrphanRows -gt 0); FontSize = 11.5 },
            @{ Col1 = 9;  Col2 = 10; Label = '空テーブル / 書式異常行';   Value = ("空 {0} 表 / 異常 {1:#,##0} 行" -f $Analysis.EmptyTableCount, $Analysis.RaggedRowTotal); Fmt = '@'; Warn = ($Analysis.EmptyTableCount -gt 0 -or $Analysis.RaggedRowTotal -gt 0); FontSize = 11.5 },
            @{ Col1 = 11; Col2 = 12; Label = '推定ACU目安 (現状 / 10倍)'; Value = ("{0:0.0} ACU / {1:0.0} ACU" -f $Analysis.EstMinAcuCurrent, $Analysis.EstMinAcu10x); Fmt = '@'; Warn = $false; FontSize = 11.5 }
        )
        Write-KpiCardBand -Sheet $wsSummary -LabelRow 7 -Kpis $kpiRow2
        $wsSummary.Rows.Item(9).RowHeight = 10

        # Section 1: Top Tables Summary Table (Rows 10..)
        Set-SectionHeader -Sheet $wsSummary -Row 10 -StartCol 1 -EndCol 12 -Text '1. テーブル別レコード数・容量構成比 TOP15 (性能試験データ規模)'
        $topTables = @($Analysis.Tables | Select-Object -First 15)
        $mTop = New-Object 'object[,]' ([Math]::Max(1, $topTables.Count)), 12
        for ($i = 0; $i -lt $topTables.Count; $i++) {
            $t = $topTables[$i]
            $mTop[$i, 0]  = $i + 1
            $mTop[$i, 1]  = $t.TableName
            $mTop[$i, 2]  = $t.InferredRole
            $mTop[$i, 3]  = [double]$t.RowCount
            $mTop[$i, 4]  = [double]$t.RowShare
            $mTop[$i, 5]  = [double]$t.ColumnCount
            $mTop[$i, 6]  = [Math]::Round($t.FileSizeBytes / 1024.0, 2)
            $mTop[$i, 7]  = [double]$t.SizeShare
            $mTop[$i, 8]  = [Math]::Round($t.EstInnoDbBytes / 1024.0, 1)
            $mTop[$i, 9]  = [double]$t.AvgNullRate
            $mTop[$i, 10] = [double]$t.WarningCount
            $mTop[$i, 11] = if ([string]::IsNullOrEmpty($t.WarningSummary)) { '正常' } else { $t.WarningSummary }
        }
        $sumEndRow = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsSummary -StartRow 11 -StartCol 1 `
            -Headers @('順位', 'テーブル名', '推定役割', 'レコード数', '件数構成比', '列数', $sizeColHeader, '容量構成比', '推定InnoDB(KB)', '平均NULL率', '警告数', '健全性・注意点') `
            -DataMatrix $mTop -RowCount $topTables.Count `
            -ColumnFormats @{ 1 = '#,##0'; 2 = '@'; 4 = '#,##0'; 5 = '0.0%'; 6 = '#,##0'; 7 = '#,##0.00'; 8 = '0.0%'; 9 = '#,##0.0'; 10 = '0.0%'; 11 = '#,##0'; 12 = '@' } `
            -CenterCols @(1, 3, 6, 11) -RightCols @(4, 5, 7, 8, 9, 10) -WrapCols @(12) -WarnCheckCols @(12)

        for ($i = 0; $i -lt $topTables.Count; $i++) {
            if ($topTables[$i].WarningCount -gt 0) {
                $wc = $wsSummary.Cells.Item(12 + $i, 11)
                $wc.Font.Color = $warnOle; $wc.Font.Bold = $true; $wc.Interior.Color = $warnBgOle
            }
        }

        $sumTotalRow = $sumEndRow + 1
        $totRowLabel = if ($Analysis.TableCount -gt 15) { "全$($Analysis.TableCount)テーブル合計" } else { "全体合計 ($($Analysis.TableCount)テーブル)" }
        Write-AnalysisTotalRow -Sheet $wsSummary -Row $sumTotalRow -StartCol 1 `
            -Values @('計', $totRowLabel, '-', [double]$Analysis.TotalRows, 1.0, [double]$Analysis.TotalColumns, [Math]::Round($Analysis.TotalFileBytes / 1024.0, 2), 1.0, [Math]::Round($Analysis.TotalEstInnoDbBytes / 1024.0, 1), [double]$Analysis.OverallAvgNullRate, [double]$Analysis.TotalWarningCount, ("高 {0} / 中 {1} / 低 {2}" -f $Analysis.HighWarningCount, $Analysis.MediumWarningCount, $Analysis.LowWarningCount)) `
            -Formats @{ 4 = '#,##0'; 5 = '0.0%'; 6 = '#,##0'; 7 = '#,##0.00'; 8 = '0.0%'; 9 = '#,##0.0'; 10 = '0.0%'; 11 = '#,##0'; 12 = '@' } `
            -CenterCols @(1, 3, 6, 11) -RightCols @(4, 5, 7, 8, 9, 10) -WarnPositiveCols @(11)

        # Charts on Summary Sheet (placed cleanly below total row; data rows are 12..topEndRowIdx)
        $chartTopRow = $sumTotalRow + 2
        $chartTopPt = $wsSummary.Rows.Item($chartTopRow).Top
        $chartHeight = 250.0
        if ($topTables.Count -ge 1) {
            $topEndRowIdx = 11 + $topTables.Count
            [void](Add-StyledExcelChart -Sheet $wsSummary -Left 6 -Top $chartTopPt -Width 460 -Height $chartHeight `
                -ChartType 57 -Title 'テーブル別レコード数 TOP15' `
                -CategoryAddress "B12:B$topEndRowIdx" -ValueAddresses @("D12:D$topEndRowIdx") `
                -SeriesNames @('レコード数') -SeriesColors @($script:THEME.AccentHex) -ReverseCategories $true -ShowDataLabels $true)

            $pieCount = [Math]::Min(10, $topTables.Count)
            $pieEndRowIdx = 11 + $pieCount
            [void](Add-StyledExcelChart -Sheet $wsSummary -Left 480 -Top $chartTopPt -Width 440 -Height $chartHeight `
                -ChartType -4120 -Title 'テーブル別データ容量構成比 TOP10 (KB)' `
                -CategoryAddress "B12:B$pieEndRowIdx" -ValueAddresses @("G12:G$pieEndRowIdx") `
                -SeriesNames @('容量(KB)') -ShowLegend $true -ShowDataLabels $true)
        }

        # Section 2: Role Breakdown Summary
        $curRow = if ($topTables.Count -ge 1) { Get-RowBelowPoints -Sheet $wsSummary -StartRow ($chartTopRow + 1) -TargetTopPt ($chartTopPt + $chartHeight + 16) } else { $chartTopRow }
        Set-SectionHeader -Sheet $wsSummary -Row $curRow -StartCol 1 -EndCol 12 -Text '2. 推定役割別テーブル構成サマリー (データ分類・規模内訳)'
        $curRow++
        $mRole = New-Object 'object[,]' ([Math]::Max(1, $roleSummary.Count)), 12
        for ($i = 0; $i -lt $roleSummary.Count; $i++) {
            $rg = $roleSummary[$i]
            $tblShare = if ($Analysis.TableCount -gt 0) { [double]$rg.TableCount / $Analysis.TableCount } else { 0.0 }
            $rowShare = if ($Analysis.TotalRows -gt 0) { [double]$rg.TotalRows / $Analysis.TotalRows } else { 0.0 }
            $szShare  = if ($Analysis.TotalFileBytes -gt 0) { [double]$rg.TotalFileBytes / $Analysis.TotalFileBytes } else { 0.0 }
            $mRole[$i, 0]  = $i + 1
            $mRole[$i, 1]  = $rg.Role
            $mRole[$i, 2]  = [double]$rg.TableCount
            $mRole[$i, 3]  = $tblShare
            $mRole[$i, 4]  = [double]$rg.TotalRows
            $mRole[$i, 5]  = $rowShare
            $mRole[$i, 6]  = [double]$rg.TotalCols
            $mRole[$i, 7]  = [Math]::Round($rg.TotalFileBytes / 1024.0, 2)
            $mRole[$i, 8]  = $szShare
            $mRole[$i, 9]  = [Math]::Round($rg.TotalInnoBytes / 1024.0, 1)
            $mRole[$i, 10] = [double]$rg.WarningCount
            $mRole[$i, 11] = ($rg.Examples -join ', ') + $(if ($rg.TableCount -gt $rg.Examples.Count) { " 他$($rg.TableCount - $rg.Examples.Count)" } else { '' })
        }
        $roleEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsSummary -StartRow $curRow -StartCol 1 `
            -Headers @('No', '推定役割', 'テーブル数', 'テーブル比', '合計レコード数', '件数構成比', '合計列数', $sizeColHeader, '容量構成比', '推定InnoDB(KB)', '警告数', '該当テーブル例 (上位)') `
            -DataMatrix $mRole -RowCount $roleSummary.Count `
            -ColumnFormats @{ 1 = '#,##0'; 2 = '@'; 3 = '#,##0'; 4 = '0.0%'; 5 = '#,##0'; 6 = '0.0%'; 7 = '#,##0'; 8 = '#,##0.00'; 9 = '0.0%'; 10 = '#,##0.0'; 11 = '#,##0'; 12 = '@' } `
            -CenterCols @(1, 3, 7, 11) -RightCols @(4, 5, 6, 8, 9, 10) -WrapCols @(12) -WarnCheckCols @(2)
        Write-AnalysisTotalRow -Sheet $wsSummary -Row ($roleEnd + 1) -StartCol 1 `
            -Values @('計', '全役割合計', [double]$Analysis.TableCount, 1.0, [double]$Analysis.TotalRows, 1.0, [double]$Analysis.TotalColumns, [Math]::Round($Analysis.TotalFileBytes / 1024.0, 2), 1.0, [Math]::Round($Analysis.TotalEstInnoDbBytes / 1024.0, 1), [double]$Analysis.TotalWarningCount, "空テーブル $($Analysis.EmptyTableCount) 件") `
            -Formats @{ 3 = '#,##0'; 4 = '0.0%'; 5 = '#,##0'; 6 = '0.0%'; 7 = '#,##0'; 8 = '#,##0.00'; 9 = '0.0%'; 10 = '#,##0.0'; 11 = '#,##0'; 12 = '@' } `
            -CenterCols @(1, 3, 7, 11) -RightCols @(4, 5, 6, 8, 9, 10) -WarnPositiveCols @(11)

        # Section 3: Column Data Type & Quality Breakdown
        $curRow = $roleEnd + 3
        Set-SectionHeader -Sheet $wsSummary -Row $curRow -StartCol 1 -EndCol 12 -Text '3. データ型別カラム構成 ＆ 品質サマリー (全カラム内訳)'
        $curRow++
        $totColsDenom = [Math]::Max(1, $allCols.Count)
        $mType = New-Object 'object[,]' ([Math]::Max(1, $typeSummary.Count)), 12
        for ($i = 0; $i -lt $typeSummary.Count; $i++) {
            $tg = $typeSummary[$i]
            $avgNull = if ($tg.ColCount -gt 0) { $tg.NullRateSum / $tg.ColCount } else { 0.0 }
            $avgCard = if ($tg.ColCount -gt 0) { $tg.CardRateSum / $tg.ColCount } else { 0.0 }
            $avgB    = if ($tg.NonEmptyCols -gt 0) { [Math]::Round($tg.AvgByteSum / $tg.NonEmptyCols, 1) } else { 0.0 }
            $riskTxt = if ($tg.HighSkewCnt -gt 0 -or $tg.AllNullCnt -gt 0) { "偏り:$($tg.HighSkewCnt) / 全NULL:$($tg.AllNullCnt)" } else { '正常' }
            $mType[$i, 0]  = $i + 1
            $mType[$i, 1]  = $tg.DataType
            $mType[$i, 2]  = [double]$tg.ColCount
            $mType[$i, 3]  = [double]$tg.ColCount / $totColsDenom
            $mType[$i, 4]  = [double]$tg.PkCount
            $mType[$i, 5]  = [double]$tg.CatCount
            $mType[$i, 6]  = $avgNull
            $mType[$i, 7]  = $avgCard
            $mType[$i, 8]  = $avgB
            $mType[$i, 9]  = [double]$tg.MaxByteLen
            $mType[$i, 10] = $riskTxt
            $mType[$i, 11] = ($tg.Examples -join ', ') + $(if ($tg.ColCount -gt $tg.Examples.Count) { " 他$($tg.ColCount - $tg.Examples.Count)" } else { '' })
        }
        $typeEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsSummary -StartRow $curRow -StartCol 1 `
            -Headers @('No', '推定データ型', 'カラム数', '構成比', 'PK候補', '区分列', '平均NULL率', '平均ユニーク率', '平均長(B)', '最大長(B)', '偏り/全NULL', '代表カラム例') `
            -DataMatrix $mType -RowCount $typeSummary.Count `
            -ColumnFormats @{ 1 = '#,##0'; 2 = '@'; 3 = '#,##0'; 4 = '0.0%'; 5 = '#,##0'; 6 = '#,##0'; 7 = '0.0%'; 8 = '0.0%'; 9 = '#,##0.0'; 10 = '#,##0'; 11 = '@'; 12 = '@' } `
            -CenterCols @(1, 3, 5, 6, 11) -RightCols @(4, 7, 8, 9, 10) -WrapCols @(11, 12) -WarnCheckCols @(11)

        # Section 4: Top Warnings & Risk Summary (merged cleanly across cols 1..12)
        $curRow = $typeEnd + 2
        $topWarns = @($Analysis.Warnings | Select-Object -First 20)
        $warnSecTitle = if ($Analysis.Warnings.Count -gt 20) {
            "4. 検出された主要警告・リスクサマリー (重要度順 上位20件 / 全$($Analysis.Warnings.Count)件)"
        }
        else {
            "4. 検出された主要警告・リスクサマリー (全$($Analysis.Warnings.Count)件)"
        }
        Set-SectionHeader -Sheet $wsSummary -Row $curRow -StartCol 1 -EndCol 12 -Text $warnSecTitle
        $curRow++
        # Header row for Section 4 (Col 1, 2, 3, 4, 5-7, 8-12)
        $wSpans = @(
            @{ C1 = 1; C2 = 1; Title = '重要度' },
            @{ C1 = 2; C2 = 2; Title = 'カテゴリ' },
            @{ C1 = 3; C2 = 3; Title = 'テーブル名' },
            @{ C1 = 4; C2 = 4; Title = 'カラム名' },
            @{ C1 = 5; C2 = 7; Title = '検出値・指標' },
            @{ C1 = 8; C2 = 12; Title = 'Aurora / 性能試験への影響 ＆ 推奨アクション' }
        )
        $wsSummary.Rows.Item($curRow).RowHeight = 26
        foreach ($sp in $wSpans) {
            $rH = $wsSummary.Range($wsSummary.Cells.Item($curRow, $sp.C1), $wsSummary.Cells.Item($curRow, $sp.C2))
            if ($sp.C2 -gt $sp.C1) { $rH.Merge() }
            [AuroraAnalyzer.ComHelper]::SetValue2($wsSummary.Cells.Item($curRow, $sp.C1), $sp.Title)
            $rH.Font.Name = $script:THEME.FontName; $rH.Font.Size = 9.5; $rH.Font.Bold = $true; $rH.Font.Color = $whiteOle
            $rH.Interior.Color = $navyOle; $rH.HorizontalAlignment = -4108; $rH.VerticalAlignment = -4108
            foreach ($edge in @(7, 8, 9, 10)) { try { $hb = $rH.Borders.Item($edge); $hb.LineStyle = 1; $hb.Weight = 2; $hb.Color = $navyOle } catch { } }
        }
        $curRow++
        if ($topWarns.Count -eq 0) {
            $rEmpty = $wsSummary.Range($wsSummary.Cells.Item($curRow, 1), $wsSummary.Cells.Item($curRow, 12))
            $rEmpty.Merge()
            [AuroraAnalyzer.ComHelper]::SetValue2($wsSummary.Cells.Item($curRow, 1), '警告・注意項目は検出されませんでした (良好)')
            $rEmpty.Font.Name = $script:THEME.FontName; $rEmpty.Font.Size = 9.5; $rEmpty.Font.Color = $mutedOle
            $rEmpty.Interior.Color = $whiteOle; $rEmpty.HorizontalAlignment = -4108; $rEmpty.VerticalAlignment = -4108
            foreach ($edge in @(7, 8, 9, 10)) { try { $eb = $rEmpty.Borders.Item($edge); $eb.LineStyle = 1; $eb.Weight = 2; $eb.Color = $borderOle } catch { } }
            $wsSummary.Rows.Item($curRow).RowHeight = 24
            $curRow++
        }
        else {
            for ($i = 0; $i -lt $topWarns.Count; $i++) {
                $w = $topWarns[$i]
                $bgCol = if (($i % 2) -eq 1) { $zebraOle } else { $whiteOle }
                $isHiMed = ($w.Severity -eq '高' -or $w.Severity -eq '中')
                $actionText = "【影響】$($w.PerfImpact)  【対策】$($w.Recommendation)"
                $vals = @($w.Severity, $w.Category, $w.TableName, $w.ColumnName, $w.MetricValue, $actionText)
                for ($k = 0; $k -lt $wSpans.Count; $k++) {
                    $sp = $wSpans[$k]
                    $rng = $wsSummary.Range($wsSummary.Cells.Item($curRow, $sp.C1), $wsSummary.Cells.Item($curRow, $sp.C2))
                    if ($sp.C2 -gt $sp.C1) { $rng.Merge() }
                    $rng.NumberFormat = '@'
                    [AuroraAnalyzer.ComHelper]::SetValue2($wsSummary.Cells.Item($curRow, $sp.C1), $vals[$k])
                    $rng.Font.Name = $script:THEME.FontName; $rng.Font.Size = 9.5; $rng.VerticalAlignment = -4108
                    $rng.WrapText = ($k -ge 4)
                    $rng.HorizontalAlignment = if ($k -le 1) { -4108 } else { -4131 }
                    if ($isHiMed -and $k -le 1) {
                        $rng.Font.Color = $warnOle; $rng.Font.Bold = $true; $rng.Interior.Color = $warnBgOle
                    }
                    else {
                        $rng.Font.Color = $textOle; $rng.Interior.Color = $bgCol
                    }
                    foreach ($edge in @(7, 8, 9, 10)) { try { $b = $rng.Borders.Item($edge); $b.LineStyle = 1; $b.Weight = 2; $b.Color = $borderOle } catch { } }
                }
                $wsSummary.Rows.Item($curRow).RowHeight = Get-EstimatedRowHeight -Texts @($w.MetricValue, $actionText) -ColWidths @(34.0, 76.0) -MinHeight 24.0 -MaxHeight 85.0
                $curRow++
            }
        }

        # Section 5: Relations & JMeter Readiness Highlights
        $curRow++
        Set-SectionHeader -Sheet $wsSummary -Row $curRow -StartCol 1 -EndCol 12 -Text '5. テーブル間リレーション概要 ＆ JMeterパラメータ化適性ハイライト'
        $curRow++
        $topRels = @($Analysis.Relationships | Select-Object -First 10)
        $rSpans = @(
            @{ C1 = 1;  C2 = 1;  Title = 'No' },
            @{ C1 = 2;  C2 = 3;  Title = '親テーブル.PK → 子テーブル.FK' },
            @{ C1 = 4;  C2 = 4;  Title = '参照一致率' },
            @{ C1 = 5;  C2 = 5;  Title = '孤立行数' },
            @{ C1 = 6;  C2 = 7;  Title = '平均 / P95 (子/親)' },
            @{ C1 = 8;  C2 = 8;  Title = '最大(子/親)' },
            @{ C1 = 9;  C2 = 10; Title = 'リスク評価' },
            @{ C1 = 11; C2 = 12; Title = '所見・スーパーユーザー情報' }
        )
        $wsSummary.Rows.Item($curRow).RowHeight = 25
        foreach ($sp in $rSpans) {
            $rH = $wsSummary.Range($wsSummary.Cells.Item($curRow, $sp.C1), $wsSummary.Cells.Item($curRow, $sp.C2))
            if ($sp.C2 -gt $sp.C1) { $rH.Merge() }
            [AuroraAnalyzer.ComHelper]::SetValue2($wsSummary.Cells.Item($curRow, $sp.C1), $sp.Title)
            $rH.Font.Name = $script:THEME.FontName; $rH.Font.Size = 9.5; $rH.Font.Bold = $true; $rH.Font.Color = $whiteOle
            $rH.Interior.Color = $navyOle; $rH.HorizontalAlignment = -4108; $rH.VerticalAlignment = -4108
            foreach ($edge in @(7, 8, 9, 10)) { try { $hb = $rH.Borders.Item($edge); $hb.LineStyle = 1; $hb.Weight = 2; $hb.Color = $navyOle } catch { } }
        }
        $curRow++
        if ($topRels.Count -eq 0) {
            $rEmpty = $wsSummary.Range($wsSummary.Cells.Item($curRow, 1), $wsSummary.Cells.Item($curRow, 12))
            $rEmpty.Merge()
            [AuroraAnalyzer.ComHelper]::SetValue2($wsSummary.Cells.Item($curRow, 1), '自動推定されたテーブル間FKリレーションはありません')
            $rEmpty.Font.Name = $script:THEME.FontName; $rEmpty.Font.Size = 9.5; $rEmpty.Font.Color = $mutedOle
            $rEmpty.Interior.Color = $whiteOle; $rEmpty.HorizontalAlignment = -4108; $rEmpty.VerticalAlignment = -4108
            foreach ($edge in @(7, 8, 9, 10)) { try { $eb = $rEmpty.Borders.Item($edge); $eb.LineStyle = 1; $eb.Weight = 2; $eb.Color = $borderOle } catch { } }
            $wsSummary.Rows.Item($curRow).RowHeight = 24
            $curRow++
        }
        else {
            for ($i = 0; $i -lt $topRels.Count; $i++) {
                $rl = $topRels[$i]
                $bgCol = if (($i % 2) -eq 1) { $zebraOle } else { $whiteOle }
                $relPair = "$($rl.ParentTable).$($rl.ParentColumn) → $($rl.ChildTable).$($rl.ChildColumn)"
                $avgP95  = ("平均 {0:0.00} / P95 {1:0.0}" -f $rl.FanOutAvg, $rl.FanOutP95)
                $rVals = @([double]($i + 1), $relPair, [double]$rl.MatchRate, [double]$rl.OrphanRows, $avgP95, [double]$rl.FanOutMax, $rl.RiskLevel, $rl.Note)
                $rFmts = @('#,##0', '@', '0.0%', '#,##0', '@', '#,##0', '@', '@')
                for ($k = 0; $k -lt $rSpans.Count; $k++) {
                    $sp = $rSpans[$k]
                    $rng = $wsSummary.Range($wsSummary.Cells.Item($curRow, $sp.C1), $wsSummary.Cells.Item($curRow, $sp.C2))
                    if ($sp.C2 -gt $sp.C1) { $rng.Merge() }
                    $rng.NumberFormat = $rFmts[$k]
                    Set-ExcelCellValue -Cell $wsSummary.Cells.Item($curRow, $sp.C1) -Value $rVals[$k]
                    $rng.Font.Name = $script:THEME.FontName; $rng.Font.Size = 9.5; $rng.VerticalAlignment = -4108
                    $rng.WrapText = ($k -eq 7)
                    $rng.HorizontalAlignment = if ($k -in @(0, 4, 6)) { -4108 } elseif ($k -in @(2, 3, 5)) { -4152 } else { -4131 }
                    $isW = ($k -eq 3 -and $rl.OrphanRows -gt 0) -or (($k -eq 6 -or $k -eq 7) -and (Test-WarningCellText -Text ([string]$rVals[$k])))
                    if ($isW) { $rng.Font.Color = $warnOle; $rng.Font.Bold = $true; $rng.Interior.Color = $warnBgOle }
                    else { $rng.Font.Color = $textOle; $rng.Interior.Color = $bgCol }
                    foreach ($edge in @(7, 8, 9, 10)) { try { $b = $rng.Borders.Item($edge); $b.LineStyle = 1; $b.Weight = 2; $b.Color = $borderOle } catch { } }
                }
                $wsSummary.Rows.Item($curRow).RowHeight = Get-EstimatedRowHeight -Texts @($rl.Note) -ColWidths @(40.0) -MinHeight 22.0 -MaxHeight 65.0
                $curRow++
            }
        }

        # Top JMeter Candidates compact table right under relations
        $curRow++
        $topJcs = @($Analysis.JMeterCandidates | Select-Object -First 8)
        $jSpans = @(
            @{ C1 = 1; C2 = 1;  Title = 'No' },
            @{ C1 = 2; C2 = 2;  Title = 'テーブル名' },
            @{ C1 = 3; C2 = 3;  Title = 'キーカラム' },
            @{ C1 = 4; C2 = 4;  Title = 'データ型' },
            @{ C1 = 5; C2 = 5;  Title = '一意キー数' },
            @{ C1 = 6; C2 = 7;  Title = 'JMeter並列適性' },
            @{ C1 = 8; C2 = 12; Title = '推奨設定・活用指針' }
        )
        $wsSummary.Rows.Item($curRow).RowHeight = 25
        foreach ($sp in $jSpans) {
            $rH = $wsSummary.Range($wsSummary.Cells.Item($curRow, $sp.C1), $wsSummary.Cells.Item($curRow, $sp.C2))
            if ($sp.C2 -gt $sp.C1) { $rH.Merge() }
            [AuroraAnalyzer.ComHelper]::SetValue2($wsSummary.Cells.Item($curRow, $sp.C1), $sp.Title)
            $rH.Font.Name = $script:THEME.FontName; $rH.Font.Size = 9.5; $rH.Font.Bold = $true; $rH.Font.Color = $whiteOle
            $rH.Interior.Color = $navyOle; $rH.HorizontalAlignment = -4108; $rH.VerticalAlignment = -4108
            foreach ($edge in @(7, 8, 9, 10)) { try { $hb = $rH.Borders.Item($edge); $hb.LineStyle = 1; $hb.Weight = 2; $hb.Color = $navyOle } catch { } }
        }
        $curRow++
        if ($topJcs.Count -eq 0) {
            $rEmpty = $wsSummary.Range($wsSummary.Cells.Item($curRow, 1), $wsSummary.Cells.Item($curRow, 12))
            $rEmpty.Merge()
            [AuroraAnalyzer.ComHelper]::SetValue2($wsSummary.Cells.Item($curRow, 1), '該当するJMeter一意キー候補 (10件以上) はありません')
            $rEmpty.Font.Name = $script:THEME.FontName; $rEmpty.Font.Size = 9.5; $rEmpty.Font.Color = $mutedOle
            $rEmpty.Interior.Color = $whiteOle; $rEmpty.HorizontalAlignment = -4108; $rEmpty.VerticalAlignment = -4108
            foreach ($edge in @(7, 8, 9, 10)) { try { $eb = $rEmpty.Borders.Item($edge); $eb.LineStyle = 1; $eb.Weight = 2; $eb.Color = $borderOle } catch { } }
            $wsSummary.Rows.Item($curRow).RowHeight = 24
            $curRow++
        }
        else {
            for ($i = 0; $i -lt $topJcs.Count; $i++) {
                $jc = $topJcs[$i]
                $bgCol = if (($i % 2) -eq 1) { $zebraOle } else { $whiteOle }
                $jVals = @([double]($i + 1), $jc.TableName, $jc.KeyColumn, $jc.DataType, [double]$jc.UniqueCount, $jc.Suitability, $jc.RecommendedUsage)
                $jFmts = @('#,##0', '@', '@', '@', '#,##0', '@', '@')
                for ($k = 0; $k -lt $jSpans.Count; $k++) {
                    $sp = $jSpans[$k]
                    $rng = $wsSummary.Range($wsSummary.Cells.Item($curRow, $sp.C1), $wsSummary.Cells.Item($curRow, $sp.C2))
                    if ($sp.C2 -gt $sp.C1) { $rng.Merge() }
                    $rng.NumberFormat = $jFmts[$k]
                    Set-ExcelCellValue -Cell $wsSummary.Cells.Item($curRow, $sp.C1) -Value $jVals[$k]
                    $rng.Font.Name = $script:THEME.FontName; $rng.Font.Size = 9.5; $rng.VerticalAlignment = -4108
                    $rng.WrapText = ($k -eq 6)
                    $rng.HorizontalAlignment = if ($k -in @(0, 3, 5)) { -4108 } elseif ($k -eq 4) { -4152 } else { -4131 }
                    if ($k -eq 5 -and $jc.Suitability -like '最適*') {
                        $rng.Font.Color = $navyOle; $rng.Font.Bold = $true; $rng.Interior.Color = $bgCol
                    }
                    elseif ($k -eq 5 -and (Test-WarningCellText -Text $jc.Suitability)) {
                        $rng.Font.Color = $warnOle; $rng.Font.Bold = $true; $rng.Interior.Color = $warnBgOle
                    }
                    else {
                        $rng.Font.Color = $textOle; $rng.Interior.Color = $bgCol
                    }
                    foreach ($edge in @(7, 8, 9, 10)) { try { $b = $rng.Borders.Item($edge); $b.LineStyle = 1; $b.Weight = 2; $b.Color = $borderOle } catch { } }
                }
                $wsSummary.Rows.Item($curRow).RowHeight = Get-EstimatedRowHeight -Texts @($jc.RecommendedUsage) -ColWidths @(76.0) -MinHeight 22.0 -MaxHeight 65.0
                $curRow++
            }
        }

        # Section 6: Aurora Insights & Key Findings
        $insRow = $curRow + 1
        Set-SectionHeader -Sheet $wsSummary -Row $insRow -StartCol 1 -EndCol 12 -Text '6. Aurora Serverless v2 (MySQL 8.4.7) 性能試験向け総合所見'
        $insRow++
        foreach ($ins in $Analysis.AuroraInsights) {
            $rIns = $wsSummary.Range($wsSummary.Cells.Item($insRow, 1), $wsSummary.Cells.Item($insRow, 12))
            $rIns.Merge()
            $rIns.NumberFormat = '@'
            [AuroraAnalyzer.ComHelper]::SetValue2($wsSummary.Cells.Item($insRow, 1), $ins)
            $rIns.Font.Name = $script:THEME.FontName
            $rIns.Font.Size = 9.5
            $rIns.Font.Color = $textOle
            $rIns.WrapText = $true
            $rIns.HorizontalAlignment = -4131
            $rIns.VerticalAlignment = -4108
            $rIns.IndentLevel = 1
            $rIns.Interior.Color = $zebraOle
            $rIns.Borders.Item(7).LineStyle = 1; $rIns.Borders.Item(7).Color = $navyOle; $rIns.Borders.Item(7).Weight = 3
            $rIns.Borders.Item(9).LineStyle = 1; $rIns.Borders.Item(9).Color = $borderOle; $rIns.Borders.Item(9).Weight = 2
            $wsSummary.Rows.Item($insRow).RowHeight = Get-EstimatedRowHeight -Texts @($ins) -ColWidths @(165.0) -MinHeight 26.0 -MaxHeight 85.0
            $insRow++
        }

        if (-not $DetailedAnalysis) {
            $rNote = $wsSummary.Range($wsSummary.Cells.Item($insRow + 1, 1), $wsSummary.Cells.Item($insRow + 1, 12))
            $rNote.Merge()
            $rNote.NumberFormat = '@'
            [AuroraAnalyzer.ComHelper]::SetValue2($wsSummary.Cells.Item($insRow + 1, 1), '※ カテゴリ分布・数値統計ヒストグラム・時系列分析・リレーション詳細・性能試験ガイド全件シートを出力するには、ダイアログの「詳細分析」にチェックを入れて実行してください。')
            $rNote.Font.Name = $script:THEME.FontName; $rNote.Font.Size = 8.5; $rNote.Font.Color = $mutedOle
            $rNote.HorizontalAlignment = -4131; $rNote.VerticalAlignment = -4108; $rNote.IndentLevel = 1
            $wsSummary.Rows.Item($insRow + 1).RowHeight = 20
        }

        Format-ExcelSheetCommon -Sheet $wsSummary -FreezeRow 0 -Landscape $true

        # ======================================================================
        #  SHEET 2: 02_テーブル一覧 (All Tables Complete Detailed Metrics - 29 cols)
        # ======================================================================
        $pctTbl = if ($DetailedAnalysis) { 68 } else { 76 }
        Write-XlsLog "シート作成中: 02_テーブル一覧..." $pctTbl
        $tblWidths = @(5, 22, 24, 15, 16, 12, 10, 7, 24, 15, 10, 11, 11, 11, 14, 13, 12, 10, 11, 22, 11, 18, 22, 20, 24, 28, 22, 8, 32)
        Set-ColumnWidths -Sheet $wsTables -Widths $tblWidths
        Set-SheetTitleBanner -Sheet $wsTables -Title $tblSheetTitle -Subtitle $subInfo -EndCol $tblWidths.Count

        $allTables = @($Analysis.Tables)
        $mTbl = New-Object 'object[,]' ([Math]::Max(1, $allTables.Count)), 29
        $minRowNonEmpty = [long]::MaxValue
        for ($i = 0; $i -lt $allTables.Count; $i++) {
            $t = $allTables[$i]
            if ($t.RowCount -gt 0 -and $t.MinRowBytes -lt $minRowNonEmpty) { $minRowNonEmpty = $t.MinRowBytes }
            $mTbl[$i, 0]  = $i + 1
            $mTbl[$i, 1]  = $t.TableName
            $mTbl[$i, 2]  = $t.FileName
            $mTbl[$i, 3]  = $t.EncodingName
            $mTbl[$i, 4]  = $t.InferredRole
            $mTbl[$i, 5]  = [double]$t.RowCount
            $mTbl[$i, 6]  = [double]$t.RowShare
            $mTbl[$i, 7]  = [double]$t.ColumnCount
            $mTbl[$i, 8]  = if ([string]::IsNullOrEmpty($t.ColumnTypeBreakdown)) { '-' } else { $t.ColumnTypeBreakdown }
            $mTbl[$i, 9]  = [Math]::Round($t.FileSizeBytes / 1024.0, 2)
            $mTbl[$i, 10] = [double]$t.SizeShare
            $mTbl[$i, 11] = [double]$t.MinRowBytes
            $mTbl[$i, 12] = [Math]::Round($t.AvgRowBytes, 1)
            $mTbl[$i, 13] = [double]$t.MaxRowBytes
            $mTbl[$i, 14] = [Math]::Round($t.EstInnoDbBytes / 1024.0, 1)
            $mTbl[$i, 15] = [double]$t.EstInnoDbPages
            $mTbl[$i, 16] = [Math]::Round($t.RowsPerInnoDbPage, 1)
            $mTbl[$i, 17] = [double]$t.AvgNullRate
            $mTbl[$i, 18] = [double]$t.AvgCardinalityRate
            $mTbl[$i, 19] = if ([string]::IsNullOrEmpty($t.QualityRiskBreakdown)) { '問題なし' } else { $t.QualityRiskBreakdown }
            $mTbl[$i, 20] = [double]$t.RaggedRowCount
            $mTbl[$i, 21] = if ([string]::IsNullOrEmpty($t.CandidatePks)) { '-' } else { $t.CandidatePks }
            $mTbl[$i, 22] = if ([string]::IsNullOrEmpty($t.CategoricalColumns)) { '-' } else { $t.CategoricalColumns }
            $mTbl[$i, 23] = if ([string]::IsNullOrEmpty($t.DateColumns)) { '-' } else { $t.DateColumns }
            $mTbl[$i, 24] = if ([string]::IsNullOrEmpty($t.PrimaryDateRange)) { '-' } else { $t.PrimaryDateRange }
            $mTbl[$i, 25] = if ([string]::IsNullOrEmpty($t.RelationSummary)) { '-' } else { $t.RelationSummary }
            $mTbl[$i, 26] = if ([string]::IsNullOrEmpty($t.JMeterReadiness)) { '対象外' } else { $t.JMeterReadiness }
            $mTbl[$i, 27] = [double]$t.WarningCount
            $mTbl[$i, 28] = if ([string]::IsNullOrEmpty($t.WarningSummary)) { '正常' } else { $t.WarningSummary }
        }
        $tblHeaders = @('No', 'テーブル名', $srcColHeader, '文字コード', '推定役割', 'レコード数', '件数構成比', '列数', 'データ型構成', $sizeColHeader, '容量構成比', '最小行長(B)', '平均行長(B)', '最大行長(B)', '推定InnoDB(KB)', '推定InnoDBページ数', '推定行/ページ', '平均NULL率', '平均ユニーク率', '品質リスク列内訳', '列数不一致行', 'PK/UK候補', 'カテゴリ区分カラム', '日時カラム', '日時データ期間', 'リレーション概要', 'JMeter並列適性', '警告数', '警告・注意点')
        $tblEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsTables -StartRow 4 -StartCol 1 `
            -Headers $tblHeaders -DataMatrix $mTbl -RowCount $allTables.Count `
            -ColumnFormats @{ 1 = '#,##0'; 2 = '@'; 3 = '@'; 6 = '#,##0'; 7 = '0.0%'; 8 = '#,##0'; 9 = '@'; 10 = '#,##0.00'; 11 = '0.0%'; 12 = '#,##0'; 13 = '#,##0.0'; 14 = '#,##0'; 15 = '#,##0.0'; 16 = '#,##0'; 17 = '#,##0.0'; 18 = '0.0%'; 19 = '0.0%'; 20 = '@'; 21 = '#,##0'; 22 = '@'; 23 = '@'; 24 = '@'; 25 = '@'; 26 = '@'; 27 = '@'; 28 = '#,##0'; 29 = '@' } `
            -CenterCols @(1, 4, 5, 8, 25, 27) -RightCols @(6, 7, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 21, 28) -WrapCols @(9, 20, 22, 23, 24, 26, 29) -WarnCheckCols @(20, 27, 29)

        for ($i = 0; $i -lt $allTables.Count; $i++) {
            $rIdx = 5 + $i
            if ($allTables[$i].RaggedRowCount -gt 0) {
                $rc = $wsTables.Cells.Item($rIdx, 21); $rc.Font.Color = $warnOle; $rc.Font.Bold = $true; $rc.Interior.Color = $warnBgOle
            }
            if ($allTables[$i].WarningCount -gt 0) {
                $wc = $wsTables.Cells.Item($rIdx, 28); $wc.Font.Color = $warnOle; $wc.Font.Bold = $true; $wc.Interior.Color = $warnBgOle
            }
            if ($allTables[$i].JMeterReadiness -like '最適*') {
                $jc = $wsTables.Cells.Item($rIdx, 27); $jc.Font.Color = $navyOle; $jc.Font.Bold = $true
            }
        }

        # Overall summary row separated by 1 blank row (keeps AutoFilter & ParseTablesSheet cleanly bounded)
        $totMinRowB = if ($minRowNonEmpty -eq [long]::MaxValue) { 0L } else { $minRowNonEmpty }
        $totRowsPerPg = if ($Analysis.TotalEstInnoDbPages -gt 0) { [Math]::Round($Analysis.TotalRows / [double]$Analysis.TotalEstInnoDbPages, 1) } else { 0.0 }
        $avgTblCard = if ($allTables.Count -gt 0) { ($allTables | Measure-Object -Property AvgCardinalityRate -Average).Average } else { 0.0 }
        $tblSumRow = $tblEnd + 2
        Write-AnalysisTotalRow -Sheet $wsTables -Row $tblSumRow -StartCol 1 `
            -Values @('合計', "全 $($Analysis.TableCount) テーブル集計", '-', '-', "空 $($Analysis.EmptyTableCount) 表", [double]$Analysis.TotalRows, 1.0, [double]$Analysis.TotalColumns, "PK候補 $($Analysis.CandidatePkColumnCount) 列", [Math]::Round($Analysis.TotalFileBytes / 1024.0, 2), 1.0, [double]$totMinRowB, [Math]::Round($Analysis.OverallAvgRowBytes, 1), [double]$Analysis.OverallMaxRowBytes, [Math]::Round($Analysis.TotalEstInnoDbBytes / 1024.0, 1), [double]$Analysis.TotalEstInnoDbPages, $totRowsPerPg, [double]$Analysis.OverallAvgNullRate, [double]$avgTblCard, "全NULL:$($Analysis.AllNullColumnCount) / 偏り:$($Analysis.HighSkewColumnCount) / 大容量:$($Analysis.LargeColumnCount)", [double]$Analysis.RaggedRowTotal, '-', '-', '-', '-', "関連 $($Analysis.Relationships.Count) 件 / 孤立 $($Analysis.TotalOrphanRows)", "候補 $($Analysis.JMeterCandidates.Count) 列", [double]$Analysis.TotalWarningCount, "高 $($Analysis.HighWarningCount) / 中 $($Analysis.MediumWarningCount) / 低 $($Analysis.LowWarningCount)") `
            -Formats @{ 6 = '#,##0'; 7 = '0.0%'; 8 = '#,##0'; 10 = '#,##0.00'; 11 = '0.0%'; 12 = '#,##0'; 13 = '#,##0.0'; 14 = '#,##0'; 15 = '#,##0.0'; 16 = '#,##0'; 17 = '#,##0.0'; 18 = '0.0%'; 19 = '0.0%'; 21 = '#,##0'; 28 = '#,##0' } `
            -CenterCols @(1, 4, 5, 8, 25, 27) -RightCols @(6, 7, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 21, 28) -WarnPositiveCols @(21, 28)

        $tblLastColL = ConvertTo-ExcelColLetter $tblWidths.Count
        Format-ExcelSheetCommon -Sheet $wsTables -FreezeRow 4 -AutoFilterRange "A4:${tblLastColL}$tblEnd" -Landscape $true

        # ======================================================================
        #  SHEET 3: 03_カラム分析 (All Columns Full Detailed Profile - 31 cols + 3 Charts)
        # ======================================================================
        $pctCol = if ($DetailedAnalysis) { 73 } else { 88 }
        Write-XlsLog "シート作成中: 03_カラム分析..." $pctCol
        $colWidths = @(22, 20, 6, 16, 14, 10, 10, 10, 9, 11, 10, 16, 9, 10, 10, 10, 17, 22, 10, 11, 11, 11, 9, 11, 20, 20, 20, 36, 44, 30, 30)
        Set-ColumnWidths -Sheet $wsCols -Widths $colWidths
        Set-SheetTitleBanner -Sheet $wsCols -Title '全カラム詳細プロファイル (選択性・NULL率・偏り・バイト長・型別統計・上位値分布・関連)' -Subtitle $subInfo -EndCol $colWidths.Count

        $mCol = New-Object 'object[,]' ([Math]::Max(1, $allCols.Count)), 31
        for ($i = 0; $i -lt $allCols.Count; $i++) {
            $c = $allCols[$i]
            $charLenTxt = if ($c.NonNullCount -gt 0) { "{0} / {1:0.0} / {2}" -f $c.MinCharLen, $c.AvgCharLen, $c.MaxCharLen } else { '-' }
            $mCol[$i, 0]  = $c.TableName
            $mCol[$i, 1]  = $c.ColumnName
            $mCol[$i, 2]  = [double]($c.ColumnIndex + 1)
            $mCol[$i, 3]  = $c.InferredType
            $mCol[$i, 4]  = if ([string]::IsNullOrEmpty($c.KeyRole)) { '一般' } else { $c.KeyRole }
            $mCol[$i, 5]  = [double]$c.TotalRows
            $mCol[$i, 6]  = [double]$c.NonNullCount
            $mCol[$i, 7]  = [double]$c.NullCount
            $mCol[$i, 8]  = [double]$c.NullRate
            $mCol[$i, 9]  = [double]$c.DistinctCount
            $mCol[$i, 10] = [double]$c.CardinalityRate
            $mCol[$i, 11] = $c.SelectivityRating
            $mCol[$i, 12] = [double]$c.MinByteLen
            $mCol[$i, 13] = [Math]::Round($c.AvgByteLen, 1)
            $mCol[$i, 14] = [double]$c.P95ByteLen
            $mCol[$i, 15] = [double]$c.MaxByteLen
            $mCol[$i, 16] = $charLenTxt
            $mCol[$i, 17] = if ($c.Top1Value.Length -gt 35) { $c.Top1Value.Substring(0, 32) + '...' } else { $c.Top1Value }
            $mCol[$i, 18] = [double]$c.Top1Count
            $mCol[$i, 19] = [double]$c.Top1Share
            $mCol[$i, 20] = [double]$c.Top3Share
            $mCol[$i, 21] = [double]$c.Top5Share
            $mCol[$i, 22] = [Math]::Round($c.GiniCoefficient, 3)
            $mCol[$i, 23] = [Math]::Round($c.NormalizedEntropy, 3)
            $mCol[$i, 24] = $c.SkewRisk
            $mCol[$i, 25] = if ([string]::IsNullOrEmpty($c.MinValueText)) { '-' } else { $c.MinValueText }
            $mCol[$i, 26] = if ([string]::IsNullOrEmpty($c.MaxValueText)) { '-' } else { $c.MaxValueText }
            $mCol[$i, 27] = if ([string]::IsNullOrEmpty($c.TopValuesSummary)) { '-' } else { $c.TopValuesSummary }
            $mCol[$i, 28] = if ([string]::IsNullOrEmpty($c.DetailStatsSummary)) { '-' } else { $c.DetailStatsSummary }
            $mCol[$i, 29] = if ([string]::IsNullOrEmpty($c.RelationSummary)) { '-' } else { $c.RelationSummary }
            $mCol[$i, 30] = if ([string]::IsNullOrEmpty($c.AlertNote)) { '正常' } else { $c.AlertNote }
        }
        $colHeaders = @('テーブル名', 'カラム名', '列順', '推定データ型', 'キー・役割', '総行数', '有効件数', 'NULL件数', 'NULL率', 'ユニーク数', 'ユニーク率', 'INDEX選択性評価', '最小長(B)', '平均長(B)', 'P95長(B)', '最大長(B)', '文字長(最小/平均/最大)', '最頻値(Top1)', 'Top1件数', 'Top1占有率', 'Top3占有率', 'Top5占有率', 'ジニ係数', '正規化エントロピー', '偏りリスク', '最小値 / 最古日時', '最大値 / 最新日時', '上位値分布 (Top5)', '型別詳細統計 (数値・日時・文字特性)', 'リレーション推定', '警告・注意点')
        $colEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsCols -StartRow 4 -StartCol 1 `
            -Headers $colHeaders -DataMatrix $mCol -RowCount $allCols.Count `
            -ColumnFormats @{ 1 = '@'; 2 = '@'; 3 = '#,##0'; 4 = '@'; 5 = '@'; 6 = '#,##0'; 7 = '#,##0'; 8 = '#,##0'; 9 = '0.0%'; 10 = '#,##0'; 11 = '0.0%'; 12 = '@'; 13 = '#,##0'; 14 = '#,##0.0'; 15 = '#,##0'; 16 = '#,##0'; 17 = '@'; 18 = '@'; 19 = '#,##0'; 20 = '0.0%'; 21 = '0.0%'; 22 = '0.0%'; 23 = '0.000'; 24 = '0.000'; 25 = '@'; 26 = '@'; 27 = '@'; 28 = '@'; 29 = '@'; 30 = '@'; 31 = '@' } `
            -CenterCols @(3, 4, 5, 12, 17, 25, 26, 27) -RightCols @(6, 7, 8, 9, 10, 11, 13, 14, 15, 16, 19, 20, 21, 22, 23, 24) -WrapCols @(28, 29, 30, 31) -WarnCheckCols @(25, 31)

        # Mini summary tables & charts for Null Rate Worst, Skew Worst & Largest Payload columns
        $worstNull = @($allCols | Where-Object { $_.TotalRows -gt 0 -and $_.NullRate -gt 0 } | Sort-Object NullRate -Descending | Select-Object -First 12)
        $worstSkew = @($allCols | Where-Object { $_.NonNullCount -ge 10 -and $_.DistinctCount -ge 2 -and -not $_.IsCandidatePk } | Sort-Object GiniCoefficient, Top1Share -Descending | Select-Object -First 12)
        $worstSize = @($allCols | Where-Object { $_.NonNullCount -gt 0 -and $_.MaxByteLen -gt 0 } | Sort-Object MaxByteLen, P95ByteLen -Descending | Select-Object -First 12)

        $auxStart = $colEnd + 3
        if ($worstNull.Count -gt 0 -or $worstSkew.Count -gt 0 -or $worstSize.Count -gt 0) {
            Set-SectionHeader -Sheet $wsCols -Row ($auxStart - 1) -StartCol 1 -EndCol $colWidths.Count -Text 'NULL率・データ偏り (ホットキー)・最大バイト長 ワーストランキング'
        }
        $nextAuxRow = $auxStart
        if ($worstNull.Count -gt 0) {
            $mWn = New-Object 'object[,]' $worstNull.Count, 5
            for ($i = 0; $i -lt $worstNull.Count; $i++) {
                $mWn[$i, 0] = $worstNull[$i].TableName + '.' + $worstNull[$i].ColumnName
                $mWn[$i, 1] = $worstNull[$i].InferredType
                $mWn[$i, 2] = [double]$worstNull[$i].NullCount
                $mWn[$i, 3] = [double]$worstNull[$i].TotalRows
                $mWn[$i, 4] = [double]$worstNull[$i].NullRate
            }
            $wnEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsCols -StartRow $auxStart -StartCol 1 `
                -Headers @('対象カラム (NULL率ワースト)', '推定データ型', 'NULL件数', '総行数', 'NULL率') -DataMatrix $mWn -RowCount $worstNull.Count `
                -ColumnFormats @{ 1 = '@'; 2 = '@'; 3 = '#,##0'; 4 = '#,##0'; 5 = '0.0%' } -CenterCols @(2) -RightCols @(3, 4, 5)

            $wnChartTop = $wsCols.Rows.Item($auxStart).Top
            $wnChartLeft = $wsCols.Columns.Item(7).Left
            $wnDataStart = $auxStart + 1
            [void](Add-StyledExcelChart -Sheet $wsCols -Left $wnChartLeft -Top $wnChartTop -Width 460 -Height 235 `
                -ChartType 57 -Title 'NULL率ワースト TOP12 カラム' `
                -CategoryAddress "A${wnDataStart}:A${wnEnd}" `
                -ValueAddresses @("E${wnDataStart}:E${wnEnd}") `
                -SeriesNames @('NULL率') -SeriesColors @($script:THEME.WarnHex) -ReverseCategories $true -ShowDataLabels $true)
            $nextAuxRow = [Math]::Max(($wnEnd + 3), (Get-RowBelowPoints -Sheet $wsCols -StartRow ($auxStart + 1) -TargetTopPt ($wnChartTop + 250)))
        }
        if ($worstSkew.Count -gt 0) {
            $skStart = $nextAuxRow
            $mWs = New-Object 'object[,]' $worstSkew.Count, 5
            for ($i = 0; $i -lt $worstSkew.Count; $i++) {
                $mWs[$i, 0] = $worstSkew[$i].TableName + '.' + $worstSkew[$i].ColumnName
                $mWs[$i, 1] = if ($worstSkew[$i].Top1Value.Length -gt 22) { $worstSkew[$i].Top1Value.Substring(0, 19) + '...' } else { $worstSkew[$i].Top1Value }
                $mWs[$i, 2] = [double]$worstSkew[$i].Top1Share
                $mWs[$i, 3] = [double]$worstSkew[$i].GiniCoefficient
                $mWs[$i, 4] = $worstSkew[$i].SkewRisk
            }
            $wsEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsCols -StartRow $skStart -StartCol 1 `
                -Headers @('対象カラム (データ偏り上位)', '最頻値(Top1)', 'Top1占有率', 'ジニ係数', '偏りリスク') -DataMatrix $mWs -RowCount $worstSkew.Count `
                -ColumnFormats @{ 1 = '@'; 2 = '@'; 3 = '0.0%'; 4 = '0.000'; 5 = '@' } -CenterCols @(5) -RightCols @(3, 4) -WarnCheckCols @(5)

            $wsChartTop = $wsCols.Rows.Item($skStart).Top
            $wsChartLeft = $wsCols.Columns.Item(7).Left
            $wsDataStart = $skStart + 1
            [void](Add-StyledExcelChart -Sheet $wsCols -Left $wsChartLeft -Top $wsChartTop -Width 460 -Height 250 `
                -ChartType 57 -Title 'データ偏り上位 TOP12 (Top1占有率 & ジニ係数)' `
                -CategoryAddress "A${wsDataStart}:A${wsEnd}" `
                -ValueAddresses @("C${wsDataStart}:C${wsEnd}", "D${wsDataStart}:D${wsEnd}") `
                -SeriesNames @('Top1占有率', 'ジニ係数') -SeriesColors @($script:THEME.AccentHex, $script:THEME.WarnHex) `
                -ShowLegend $true -ReverseCategories $true)
            $nextAuxRow = [Math]::Max(($wsEnd + 3), (Get-RowBelowPoints -Sheet $wsCols -StartRow ($skStart + 1) -TargetTopPt ($wsChartTop + 265)))
        }
        if ($worstSize.Count -gt 0) {
            $szStart = $nextAuxRow
            $mWz = New-Object 'object[,]' $worstSize.Count, 5
            for ($i = 0; $i -lt $worstSize.Count; $i++) {
                $mWz[$i, 0] = $worstSize[$i].TableName + '.' + $worstSize[$i].ColumnName
                $mWz[$i, 1] = $worstSize[$i].InferredType
                $mWz[$i, 2] = [Math]::Round($worstSize[$i].AvgByteLen, 1)
                $mWz[$i, 3] = [double]$worstSize[$i].P95ByteLen
                $mWz[$i, 4] = [double]$worstSize[$i].MaxByteLen
            }
            $wzEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsCols -StartRow $szStart -StartCol 1 `
                -Headers @('対象カラム (最大バイト長上位)', '推定データ型', '平均長(B)', 'P95長(B)', '最大長(B)') -DataMatrix $mWz -RowCount $worstSize.Count `
                -ColumnFormats @{ 1 = '@'; 2 = '@'; 3 = '#,##0.0'; 4 = '#,##0'; 5 = '#,##0' } -CenterCols @(2) -RightCols @(3, 4, 5)

            $wzChartTop = $wsCols.Rows.Item($szStart).Top
            $wzChartLeft = $wsCols.Columns.Item(7).Left
            $wzDataStart = $szStart + 1
            [void](Add-StyledExcelChart -Sheet $wsCols -Left $wzChartLeft -Top $wzChartTop -Width 460 -Height 245 `
                -ChartType 57 -Title '最大バイト長上位 TOP12 (P95長 & 最大長)' `
                -CategoryAddress "A${wzDataStart}:A${wzEnd}" `
                -ValueAddresses @("D${wzDataStart}:D${wzEnd}", "E${wzDataStart}:E${wzEnd}") `
                -SeriesNames @('P95長(B)', '最大長(B)') -SeriesColors @('#737373', $script:THEME.AccentHex) `
                -ShowLegend $true -ReverseCategories $true)
        }

        $colLastColL = ConvertTo-ExcelColLetter $colWidths.Count
        Format-ExcelSheetCommon -Sheet $wsCols -FreezeRow 4 -AutoFilterRange "A4:${colLastColL}$colEnd" -Landscape $true

        if ($DetailedAnalysis) {
        # ======================================================================
        #  SHEET 4: 04_分布・割合分析 (Categorical / Status / Enum Distributions)
        # ======================================================================
        Write-XlsLog "シート作成中: 04_分布・割合分析..." 78
        Set-ColumnWidths -Sheet $wsDist -Widths @(20, 20, 8, 6, 26, 12, 11, 11, 20, 3)
        Set-SheetTitleBanner -Sheet $wsDist -Title 'カテゴリ・ステータス・区分値の分布と割合 (JMeterシナリオ重み付け設計用)' -Subtitle $subInfo -EndCol 9

        $catColsAll = @($allCols | Where-Object { $_.IsCategorical } | Sort-Object NonNullCount -Descending)
        $distRows = New-Object System.Collections.Generic.List[object]
        $chartGroups = New-Object System.Collections.Generic.List[object]

        $curDataRow = 5
        foreach ($cc in $catColsAll) {
            $groupStart = $curDataRow
            $rk = 0
            foreach ($tv in $cc.TopValues) {
                $rk++
                $weightHint = if ($tv.Share -ge 0.50) { '主要パス (高頻度)' } elseif ($tv.Share -ge 0.15) { '標準パス' } elseif ($tv.Share -ge 0.03) { '副次パス' } else { '低頻度・エッジケース' }
                $distRows.Add([PSCustomObject]@{
                    TableName = $cc.TableName
                    ColName   = $cc.ColumnName
                    Distinct  = $cc.DistinctCount
                    Rank      = $rk
                    Value     = $tv.Value
                    Count     = $tv.Count
                    Share     = $tv.Share
                    CumShare  = $tv.CumulativeShare
                    Hint      = $weightHint
                })
                $curDataRow++
            }
            $groupEnd = $curDataRow - 1
            if ($chartGroups.Count -lt 5 -and ($groupEnd -ge $groupStart)) {
                $chartGroups.Add([PSCustomObject]@{
                    Title    = ("{0}.{1} 構成割合" -f $cc.TableName, $cc.ColumnName)
                    StartRow = $groupStart
                    EndRow   = [Math]::Min($groupEnd, $groupStart + 9)
                })
            }
        }

        $mDist = New-Object 'object[,]' ([Math]::Max(1, $distRows.Count)), 9
        for ($i = 0; $i -lt $distRows.Count; $i++) {
            $dr = $distRows[$i]
            $mDist[$i, 0] = $dr.TableName
            $mDist[$i, 1] = $dr.ColName
            $mDist[$i, 2] = [double]$dr.Distinct
            $mDist[$i, 3] = [double]$dr.Rank
            $mDist[$i, 4] = $dr.Value
            $mDist[$i, 5] = [double]$dr.Count
            $mDist[$i, 6] = [double]$dr.Share
            $mDist[$i, 7] = [double]$dr.CumShare
            $mDist[$i, 8] = $dr.Hint
        }
        $distEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsDist -StartRow 4 -StartCol 1 `
            -Headers @('テーブル名', 'カラム名', '区分数', '順位', 'カテゴリ値', '件数', '割合(%)', '累積割合(%)', '性能試験シナリオ区分') `
            -DataMatrix $mDist -RowCount $distRows.Count `
            -ColumnFormats @{ 1 = '@'; 2 = '@'; 3 = '#,##0'; 4 = '#,##0'; 5 = '@'; 6 = '#,##0'; 7 = '0.0%'; 8 = '0.0%' } `
            -CenterCols @(3, 4, 9) -RightCols @(6, 7, 8) -WrapCols @(5)

        # Place top categorical distribution charts neatly in column K (Right pane)
        $distChartLeft = $wsDist.Columns.Item(11).Left
        $distChartTop = $wsDist.Rows.Item(4).Top
        foreach ($cg in $chartGroups) {
            [void](Add-StyledExcelChart -Sheet $wsDist -Left $distChartLeft -Top $distChartTop -Width 400 -Height 210 `
                -ChartType 57 -Title $cg.Title `
                -CategoryAddress ("E{0}:E{1}" -f $cg.StartRow, $cg.EndRow) `
                -ValueAddresses @("G{0}:G{1}" -f $cg.StartRow, $cg.EndRow) `
                -SeriesNames @('構成比') -SeriesColors @($script:THEME.AccentHex) -ReverseCategories $true -ShowDataLabels $true)
            $distChartTop += 225
        }

        Format-ExcelSheetCommon -Sheet $wsDist -FreezeRow 4 -AutoFilterRange "A4:I$distEnd" -Landscape $true

        # ======================================================================
        #  SHEET 5: 05_数値統計・ヒストグラム (Numeric Percentiles & Histograms)
        # ======================================================================
        Write-XlsLog "シート作成中: 05_数値統計・ヒストグラム..." 83
        Set-ColumnWidths -Sheet $wsNum -Widths @(20, 22, 11, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 9, 9, 9)
        Set-SheetTitleBanner -Sheet $wsNum -Title '数値カラム統計量 (パーセンタイル・外れ値) ＆ 度数分布ヒストグラム' -Subtitle $subInfo -EndCol 16

        Set-SectionHeader -Sheet $wsNum -Row 4 -StartCol 1 -EndCol 16 -Text '1. 数値カラム パーセンタイル・統計量一覧'
        $numColsAll = @($allCols | Where-Object { $_.IsNumeric -and $_.NonNullCount -gt 0 } | Sort-Object NonNullCount -Descending)
        $mNum = New-Object 'object[,]' ([Math]::Max(1, $numColsAll.Count)), 16
        for ($i = 0; $i -lt $numColsAll.Count; $i++) {
            $nc = $numColsAll[$i]
            $mNum[$i, 0]  = $nc.TableName
            $mNum[$i, 1]  = $nc.ColumnName
            $mNum[$i, 2]  = [double]$nc.NonNullCount
            $mNum[$i, 3]  = [Math]::Round($nc.NumMin, 2)
            $mNum[$i, 4]  = [Math]::Round($nc.NumP25, 2)
            $mNum[$i, 5]  = [Math]::Round($nc.NumMedian, 2)
            $mNum[$i, 6]  = [Math]::Round($nc.NumMean, 2)
            $mNum[$i, 7]  = [Math]::Round($nc.NumP75, 2)
            $mNum[$i, 8]  = [Math]::Round($nc.NumP90, 2)
            $mNum[$i, 9]  = [Math]::Round($nc.NumP95, 2)
            $mNum[$i, 10] = [Math]::Round($nc.NumP99, 2)
            $mNum[$i, 11] = [Math]::Round($nc.NumMax, 2)
            $mNum[$i, 12] = [Math]::Round($nc.NumStdDev, 2)
            $mNum[$i, 13] = [double]$nc.ZeroRate
            $mNum[$i, 14] = [double]$nc.NegativeRate
            $mNum[$i, 15] = [double]$nc.OutlierRate
        }
        $numEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsNum -StartRow 5 -StartCol 1 `
            -Headers @('テーブル名', 'カラム名', '有効件数', '最小値', 'P25(Q1)', '中央値(P50)', '平均値', 'P75(Q3)', 'P90', 'P95', 'P99', '最大値', '標準偏差', 'ゼロ率', '負値率', '外れ値率') `
            -DataMatrix $mNum -RowCount $numColsAll.Count `
            -ColumnFormats @{ 1 = '@'; 2 = '@'; 3 = '#,##0'; 4 = '#,##0.00'; 5 = '#,##0.00'; 6 = '#,##0.00'; 7 = '#,##0.00'; 8 = '#,##0.00'; 9 = '#,##0.00'; 10 = '#,##0.00'; 11 = '#,##0.00'; 12 = '#,##0.00'; 13 = '#,##0.00'; 14 = '0.0%'; 15 = '0.0%'; 16 = '0.0%' } `
            -RightCols @(3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)

        # Histograms for top non-PK numeric columns
        $histCols = @($numColsAll | Where-Object { $_.Histogram.Count -gt 0 } | Sort-Object OutlierRate, NonNullCount -Descending | Select-Object -First 4)
        $hRow = $numEnd + 2
        if ($histCols.Count -gt 0) {
            Set-SectionHeader -Sheet $wsNum -Row $hRow -StartCol 1 -EndCol 16 -Text '2. 主要数値カラム 度数分布ヒストグラム (性能試験データレンジ設計)'
            $hRow++
            foreach ($hc in $histCols) {
                $mH = New-Object 'object[,]' $hc.Histogram.Count, 5
                for ($b = 0; $b -lt $hc.Histogram.Count; $b++) {
                    $bin = $hc.Histogram[$b]
                    $mH[$b, 0] = $hc.TableName + '.' + $hc.ColumnName
                    $mH[$b, 1] = $bin.Label
                    $mH[$b, 2] = [double]$bin.Count
                    $mH[$b, 3] = [double]$bin.Share
                    $mH[$b, 4] = [double]$bin.CumulativeShare
                }
                $hEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsNum -StartRow $hRow -StartCol 1 `
                    -Headers @('対象カラム', '区間レンジ', '推定件数', '割合(%)', '累積割合(%)') `
                    -DataMatrix $mH -RowCount $hc.Histogram.Count `
                    -ColumnFormats @{ 1 = '@'; 2 = '@'; 3 = '#,##0'; 4 = '0.0%'; 5 = '0.0%' } `
                    -RightCols @(3, 4, 5)

                $hChartTop = $wsNum.Rows.Item($hRow).Top
                $hChartLeft = $wsNum.Columns.Item(7).Left
                $hDataStart = $hRow + 1
                [void](Add-StyledExcelChart -Sheet $wsNum -Left $hChartLeft -Top $hChartTop -Width 480 -Height 225 `
                    -ChartType 51 -Title ("ヒストグラム: {0}.{1}" -f $hc.TableName, $hc.ColumnName) `
                    -CategoryAddress "B${hDataStart}:B${hEnd}" `
                    -ValueAddresses @("C${hDataStart}:C${hEnd}") `
                    -SeriesNames @('件数') -SeriesColors @($script:THEME.AccentHex) -GapWidth 25)

                $hRow = [Math]::Max(($hEnd + 3), (Get-RowBelowPoints -Sheet $wsNum -StartRow ($hRow + 1) -TargetTopPt ($hChartTop + 240)))
            }
        }

        Format-ExcelSheetCommon -Sheet $wsNum -FreezeRow 5 -Landscape $true

        # ======================================================================
        #  SHEET 6: 06_時系列分析 (Date/Time Spread, Monthly, Hourly, Day-of-Week)
        # ======================================================================
        Write-XlsLog "シート作成中: 06_時系列分析..." 88
        Set-ColumnWidths -Sheet $wsTime -Widths @(20, 20, 12, 22, 22, 11, 12, 14, 13, 13)
        Set-SheetTitleBanner -Sheet $wsTime -Title '時系列・日時分布分析 (月別増加推移・時間帯別ピーク・曜日別割合)' -Subtitle $subInfo -EndCol 10

        Set-SectionHeader -Sheet $wsTime -Row 4 -StartCol 1 -EndCol 10 -Text '1. 日時カラム別 期間・日次ピーク・バースト倍率一覧'
        $dtColsAll = @($allCols | Where-Object { $_.IsDateTime -and $_.NonNullCount -gt 0 } | Sort-Object NonNullCount -Descending)
        $mDt = New-Object 'object[,]' ([Math]::Max(1, $dtColsAll.Count)), 10
        for ($i = 0; $i -lt $dtColsAll.Count; $i++) {
            $dc = $dtColsAll[$i]
            $mDt[$i, 0] = $dc.TableName
            $mDt[$i, 1] = $dc.ColumnName
            $mDt[$i, 2] = [double]$dc.NonNullCount
            $mDt[$i, 3] = if ($dc.MinDate) { $dc.MinDate.ToString('yyyy/MM/dd HH:mm:ss') } else { '-' }
            $mDt[$i, 4] = if ($dc.MaxDate) { $dc.MaxDate.ToString('yyyy/MM/dd HH:mm:ss') } else { '-' }
            $mDt[$i, 5] = [Math]::Round($dc.DateSpanDays, 1)
            $mDt[$i, 6] = [Math]::Round($dc.AvgPerDay, 1)
            $mDt[$i, 7] = $dc.PeakDay
            $mDt[$i, 8] = [double]$dc.PeakDayCount
            $mDt[$i, 9] = [Math]::Round($dc.PeakBurstRatio, 2)
        }
        $dtEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsTime -StartRow 5 -StartCol 1 `
            -Headers @('テーブル名', '日時カラム', '有効件数', '最古日時', '最新日時', '期間(日)', '日平均件数', 'ピーク日', 'ピーク日件数', 'ピーク倍率') `
            -DataMatrix $mDt -RowCount $dtColsAll.Count `
            -ColumnFormats @{ 1 = '@'; 2 = '@'; 3 = '#,##0'; 4 = '@'; 5 = '@'; 6 = '#,##0.0'; 7 = '#,##0.0'; 8 = '@'; 9 = '#,##0'; 10 = '#,##0.00"倍"' } `
            -CenterCols @(4, 5, 8) -RightCols @(3, 6, 7, 9, 10)

        $tsSrcLabel = if ([string]::IsNullOrEmpty($Analysis.PrimaryTimeSeriesSource)) { '対象なし' } else { $Analysis.PrimaryTimeSeriesSource }

        # Monthly Distribution
        $mRow = $dtEnd + 2
        Set-SectionHeader -Sheet $wsTime -Row $mRow -StartCol 1 -EndCol 10 -Text ("2. 月別レコード推移 (代表日時軸: {0})" -f $tsSrcLabel)
        $mRow++
        $monthKeys = @($Analysis.GlobalMonthlyCounts.Keys)
        $monthTotal = 0L
        foreach ($mk in $monthKeys) { $monthTotal += $Analysis.GlobalMonthlyCounts[$mk] }
        $mMon = New-Object 'object[,]' ([Math]::Max(1, $monthKeys.Count)), 4
        $mCum = 0L
        for ($i = 0; $i -lt $monthKeys.Count; $i++) {
            $mk = $monthKeys[$i]
            $cnt = $Analysis.GlobalMonthlyCounts[$mk]
            $mCum += $cnt
            $mMon[$i, 0] = $mk
            $mMon[$i, 1] = [double]$cnt
            $mMon[$i, 2] = if ($monthTotal -gt 0) { [double]$cnt / $monthTotal } else { 0.0 }
            $mMon[$i, 3] = if ($monthTotal -gt 0) { [double]$mCum / $monthTotal } else { 0.0 }
        }
        $monEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsTime -StartRow $mRow -StartCol 1 `
            -Headers @('年月', 'レコード数', '構成比(%)', '累積構成比(%)') `
            -DataMatrix $mMon -RowCount $monthKeys.Count `
            -ColumnFormats @{ 1 = '@'; 2 = '#,##0'; 3 = '0.0%'; 4 = '0.0%' } `
            -CenterCols @(1) -RightCols @(2, 3, 4)

        $monChartTop = $wsTime.Rows.Item($mRow).Top
        if ($monthKeys.Count -ge 1) {
            $monDataStart = $mRow + 1
            [void](Add-StyledExcelChart -Sheet $wsTime -Left $wsTime.Columns.Item(6).Left -Top $monChartTop -Width 440 -Height 225 `
                -ChartType 51 -Title ("月別レコード増加推移 ({0})" -f $tsSrcLabel) `
                -CategoryAddress "A${monDataStart}:A${monEnd}" `
                -ValueAddresses @("B${monDataStart}:B${monEnd}") `
                -SeriesNames @('件数') -SeriesColors @($script:THEME.AccentHex) -ShowDataLabels $true)
        }

        # Hourly Distribution (00..23)
        $hrRow = [Math]::Max(($monEnd + 3), (Get-RowBelowPoints -Sheet $wsTime -StartRow ($mRow + 1) -TargetTopPt ($monChartTop + 240)))
        Set-SectionHeader -Sheet $wsTime -Row $hrRow -StartCol 1 -EndCol 10 -Text ("3. 時間帯別 (00時〜23時) 発生分布・割合 (JMeterピーク負荷モデル用 / {0})" -f $tsSrcLabel)
        $hrRow++
        $hrTotal = 0L
        for ($h = 0; $h -lt 24; $h++) { $hrTotal += $Analysis.GlobalHourlyCounts[$h] }
        $mHr = New-Object 'object[,]' 24, 3
        for ($h = 0; $h -lt 24; $h++) {
            $cnt = $Analysis.GlobalHourlyCounts[$h]
            $mHr[$h, 0] = ('{0:00}:00' -f $h)
            $mHr[$h, 1] = [double]$cnt
            $mHr[$h, 2] = if ($hrTotal -gt 0) { [double]$cnt / $hrTotal } else { 0.0 }
        }
        $hrEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsTime -StartRow $hrRow -StartCol 1 `
            -Headers @('時間帯', '件数', '構成比(%)') `
            -DataMatrix $mHr -RowCount 24 `
            -ColumnFormats @{ 1 = '@'; 2 = '#,##0'; 3 = '0.0%' } `
            -CenterCols @(1) -RightCols @(2, 3)

        if ($hrTotal -gt 0) {
            $hrDataStart = $hrRow + 1
            [void](Add-StyledExcelChart -Sheet $wsTime -Left $wsTime.Columns.Item(5).Left -Top $wsTime.Rows.Item($hrRow).Top -Width 490 -Height 255 `
                -ChartType 51 -Title ("時間帯別 (00〜23時) トラフィック/レコード分布 ({0})" -f $tsSrcLabel) `
                -CategoryAddress "A${hrDataStart}:A${hrEnd}" `
                -ValueAddresses @("B${hrDataStart}:B${hrEnd}") `
                -SeriesNames @('件数') -SeriesColors @($script:THEME.AccentHex) -GapWidth 40)
        }

        # Day-of-Week Distribution (Mon..Sun)
        $dowRow = $hrEnd + 2
        Set-SectionHeader -Sheet $wsTime -Row $dowRow -StartCol 1 -EndCol 10 -Text ("4. 曜日別 (月〜日) 発生分布・割合 ({0})" -f $tsSrcLabel)
        $dowRow++
        $dowNames = @('月曜', '火曜', '水曜', '木曜', '金曜', '土曜', '日曜')
        $dowTotal = 0L
        for ($d = 0; $d -lt 7; $d++) { $dowTotal += $Analysis.GlobalDayOfWeekCounts[$d] }
        $mDow = New-Object 'object[,]' 7, 3
        for ($d = 0; $d -lt 7; $d++) {
            $cnt = $Analysis.GlobalDayOfWeekCounts[$d]
            $mDow[$d, 0] = $dowNames[$d]
            $mDow[$d, 1] = [double]$cnt
            $mDow[$d, 2] = if ($dowTotal -gt 0) { [double]$cnt / $dowTotal } else { 0.0 }
        }
        $dowEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsTime -StartRow $dowRow -StartCol 1 `
            -Headers @('曜日', '件数', '構成比(%)') `
            -DataMatrix $mDow -RowCount 7 `
            -ColumnFormats @{ 1 = '@'; 2 = '#,##0'; 3 = '0.0%' } `
            -CenterCols @(1) -RightCols @(2, 3)

        if ($dowTotal -gt 0) {
            $dowDataStart = $dowRow + 1
            [void](Add-StyledExcelChart -Sheet $wsTime -Left $wsTime.Columns.Item(5).Left -Top $wsTime.Rows.Item($dowRow).Top -Width 490 -Height 215 `
                -ChartType 51 -Title ("曜日別レコード分布 ({0})" -f $tsSrcLabel) `
                -CategoryAddress "A${dowDataStart}:A${dowEnd}" `
                -ValueAddresses @("B${dowDataStart}:B${dowEnd}") `
                -SeriesNames @('件数') -SeriesColors @('#4D4D4D') -ShowDataLabels $true)
        }

        Format-ExcelSheetCommon -Sheet $wsTime -FreezeRow 5 -Landscape $true

        # ======================================================================
        #  SHEET 7: 07_リレーション・ファンアウト (Cross-Table FK & 1:N Fan-Out)
        # ======================================================================
        Write-XlsLog "シート作成中: 07_リレーション・ファンアウト..." 92
        Set-ColumnWidths -Sheet $wsRel -Widths @(24, 24, 11, 11, 11, 10, 10, 11, 10, 10, 10, 11, 24, 38)
        Set-SheetTitleBanner -Sheet $wsRel -Title 'テーブル間リレーション自動推定 ＆ 1:Nファンアウト倍率・孤立レコード分析' -Subtitle $subInfo -EndCol 14

        $rels = @($Analysis.Relationships)
        $mRel = New-Object 'object[,]' ([Math]::Max(1, $rels.Count)), 14
        for ($i = 0; $i -lt $rels.Count; $i++) {
            $rl = $rels[$i]
            $mRel[$i, 0]  = $rl.ParentTable + '.' + $rl.ParentColumn
            $mRel[$i, 1]  = $rl.ChildTable + '.' + $rl.ChildColumn
            $mRel[$i, 2]  = [double]$rl.ParentDistinctCount
            $mRel[$i, 3]  = [double]$rl.ChildNonNullRows
            $mRel[$i, 4]  = [double]$rl.MatchRate
            $mRel[$i, 5]  = [double]$rl.OrphanRows
            $mRel[$i, 6]  = [double]$rl.FanOutMin
            $mRel[$i, 7]  = [Math]::Round($rl.FanOutAvg, 2)
            $mRel[$i, 8]  = [Math]::Round($rl.FanOutP50, 1)
            $mRel[$i, 9]  = [Math]::Round($rl.FanOutP95, 1)
            $mRel[$i, 10] = [Math]::Round($rl.FanOutP99, 1)
            $mRel[$i, 11] = [double]$rl.FanOutMax
            $mRel[$i, 12] = $rl.RiskLevel
            $mRel[$i, 13] = $rl.Note
        }
        $relEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsRel -StartRow 4 -StartCol 1 `
            -Headers @('親テーブル.PK', '子テーブル.FK', '親キー数', '子有効行数', '参照一致率', '孤立行数', '最小(子/親)', '平均(子/親)', 'P50(子/親)', 'P95(子/親)', 'P99(子/親)', '最大(子/親)', 'リスク評価', '所見・スーパーユーザー情報') `
            -DataMatrix $mRel -RowCount $rels.Count `
            -ColumnFormats @{ 1 = '@'; 2 = '@'; 3 = '#,##0'; 4 = '#,##0'; 5 = '0.0%'; 6 = '#,##0'; 7 = '#,##0'; 8 = '#,##0.00'; 9 = '#,##0.0'; 10 = '#,##0.0'; 11 = '#,##0.0'; 12 = '#,##0'; 14 = '@' } `
            -CenterCols @(13) -RightCols @(3, 4, 5, 6, 7, 8, 9, 10, 11, 12) -WrapCols @(14) -WarnCheckCols @(13, 14)

        # Highlight non-zero OrphanRows in Burgundy on neutral background
        for ($i = 0; $i -lt $rels.Count; $i++) {
            if ($rels[$i].OrphanRows -gt 0) {
                $oc = $wsRel.Cells.Item(5 + $i, 6)
                $oc.Font.Color = $warnOle
                $oc.Font.Bold = $true
                $oc.Interior.Color = $warnBgOle
            }
        }

        if ($rels.Count -ge 1) {
            $relChartCount = [Math]::Min(12, $rels.Count)
            $relChartEnd = 4 + $relChartCount
            $relChartTop = $wsRel.Rows.Item($relEnd + 2).Top
            [void](Add-StyledExcelChart -Sheet $wsRel -Left 6 -Top $relChartTop -Width 660 -Height 265 `
                -ChartType 57 -Title '1:Nファンアウト倍率比較 (平均 vs P95 vs 最大 / スーパーユーザー偏り可視化)' `
                -CategoryAddress "B5:B${relChartEnd}" `
                -ValueAddresses @("H5:H${relChartEnd}", "J5:J${relChartEnd}", "L5:L${relChartEnd}") `
                -SeriesNames @('平均 (子/親)', 'P95 (子/親)', '最大 (子/親)') `
                -SeriesColors @('#737373', $script:THEME.AccentHex, $script:THEME.WarnHex) `
                -ShowLegend $true -ReverseCategories $true)
        }

        Format-ExcelSheetCommon -Sheet $wsRel -FreezeRow 4 -AutoFilterRange "A4:N$relEnd" -Landscape $true

        # ======================================================================
        #  SHEET 8: 08_性能試験ガイド・警告 (Warnings + JMeter Dataset Readiness)
        # ======================================================================
        Write-XlsLog "シート作成中: 08_性能試験ガイド・警告..." 96
        Set-ColumnWidths -Sheet $wsGuide -Widths @(10, 18, 22, 18, 28, 42, 46)
        Set-SheetTitleBanner -Sheet $wsGuide -Title '性能試験設計ガイド・検出警告一覧・JMeterパラメータ化候補' -Subtitle $subInfo -EndCol 7

        Set-SectionHeader -Sheet $wsGuide -Row 4 -StartCol 1 -EndCol 7 -Text '1. 検出された警告・注意項目一覧 (バーガンディ強調: 性能試験前に対処推奨)'
        $warns = @($Analysis.Warnings)
        $mWarn = New-Object 'object[,]' ([Math]::Max(1, $warns.Count)), 7
        for ($i = 0; $i -lt $warns.Count; $i++) {
            $w = $warns[$i]
            $mWarn[$i, 0] = $w.Severity
            $mWarn[$i, 1] = $w.Category
            $mWarn[$i, 2] = $w.TableName
            $mWarn[$i, 3] = $w.ColumnName
            $mWarn[$i, 4] = $w.MetricValue
            $mWarn[$i, 5] = $w.PerfImpact
            $mWarn[$i, 6] = $w.Recommendation
        }
        $warnEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsGuide -StartRow 5 -StartCol 1 `
            -Headers @('重要度', 'カテゴリ', 'テーブル名', 'カラム名', '検出値・指標', 'Aurora / 性能試験への影響', '推奨アクション') `
            -DataMatrix $mWarn -RowCount $warns.Count `
            -ColumnFormats @{ 1 = '@'; 2 = '@'; 3 = '@'; 4 = '@'; 5 = '@'; 6 = '@'; 7 = '@' } `
            -CenterCols @(1, 2) -WrapCols @(5, 6, 7) -WarnCheckCols @(1)

        # Highlight 高 / 中 warning severity & category in Burgundy (neutral background, zero same-family tint)
        for ($i = 0; $i -lt $warns.Count; $i++) {
            $rIdx = 6 + $i
            if ($warns[$i].Severity -eq '高' -or $warns[$i].Severity -eq '中') {
                $cSev = $wsGuide.Cells.Item($rIdx, 1)
                $cSev.Font.Color = $warnOle
                $cSev.Font.Bold = $true
                $cSev.Interior.Color = $warnBgOle
                $cCat = $wsGuide.Cells.Item($rIdx, 2)
                $cCat.Font.Color = $warnOle
                $cCat.Font.Bold = $true
            }
        }

        $jmStart = $warnEnd + 2
        Set-SectionHeader -Sheet $wsGuide -Row $jmStart -StartCol 1 -EndCol 7 -Text '2. JMeter CSV Data Set パラメータ化候補 (一意キー・並列実行適性)'
        $jmStart++
        $jcs = @($Analysis.JMeterCandidates)
        $mJm = New-Object 'object[,]' ([Math]::Max(1, $jcs.Count)), 7
        for ($i = 0; $i -lt $jcs.Count; $i++) {
            $jc = $jcs[$i]
            $mJm[$i, 0] = $jc.TableName
            $mJm[$i, 1] = $jc.KeyColumn
            $mJm[$i, 2] = $jc.DataType
            $mJm[$i, 3] = [double]$jc.UniqueCount
            $mJm[$i, 4] = [double]$jc.TotalRows
            $mJm[$i, 5] = $jc.Suitability
            $mJm[$i, 6] = $jc.RecommendedUsage
        }
        $jmEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsGuide -StartRow $jmStart -StartCol 1 `
            -Headers @('テーブル名', 'キーカラム', 'データ型', '一意キー数', '総行数', 'JMeter並列適性', '推奨設定・活用指針') `
            -DataMatrix $mJm -RowCount $jcs.Count `
            -ColumnFormats @{ 1 = '@'; 2 = '@'; 4 = '#,##0'; 5 = '#,##0'; 6 = '@'; 7 = '@' } `
            -CenterCols @(3, 6) -RightCols @(4, 5) -WrapCols @(7) -WarnCheckCols @(6)

        for ($i = 0; $i -lt $jcs.Count; $i++) {
            $cSuit = $wsGuide.Cells.Item($jmStart + 1 + $i, 6)
            if ($jcs[$i].Suitability -like '最適*') {
                $cSuit.Font.Color = $navyOle
                $cSuit.Font.Bold = $true
            }
        }

        Format-ExcelSheetCommon -Sheet $wsGuide -FreezeRow 5 -Landscape $true
        } # end if ($DetailedAnalysis)

        # Embed lossless full-fidelity snapshot via standard OOXML CustomXMLPart (zero visible sheet clutter)
        $snapshotXml = [AuroraAnalyzer.AnalysisSnapshotCodec]::SerializeToXml($Analysis, $DbLabel)
        if (-not [string]::IsNullOrEmpty($snapshotXml)) {
            [void]$wb.CustomXMLParts.Add($snapshotXml)
        }

        # Activate Summary sheet before save
        $wsSummary.Activate()
        $excel.Calculation = -4105 # xlCalculationAutomatic
        $excel.ScreenUpdating = $true

        Write-XlsLog "Excelファイル保存中: $OutputPath" 98
        $wb.SaveAs($stagingPath, 51) # 51 = xlOpenXMLWorkbook (.xlsx)
        $wb.Close($false)
        $excel.Quit()
        if ($null -ne $PumpCallback) { & $PumpCallback }
        if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }
        [AuroraAnalyzer.AtomicFile]::Publish($stagingPath, $OutputPath)
        Write-XlsLog "Excelレポート出力完了: $OutputPath" 100
    }
    finally {
        foreach ($s in $sheetsToRelease) {
            if ($null -ne $s) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($s) } catch { } }
        }
        if ($null -ne $wb) {
            try { $wb.Close($false) } catch { }
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } catch { }
        }
        if ($null -ne $excel) {
            try { $excel.Quit() } catch { }
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch { }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()

        if ($excelPid -gt 0) {
            try {
                $ep = [System.Diagnostics.Process]::GetProcessById($excelPid)
                if ($null -ne $ep -and -not $ep.HasExited) {
                    if (-not $ep.WaitForExit(1500)) {
                        $ep.Kill()
                        [void]$ep.WaitForExit(2000)
                    }
                }
            }
            catch { }
        }
        [IO.File]::Delete($stagingPath)
    }
}

# ==============================================================================
#  Module 4B: Pure .NET OpenXML (.xlsx) Fallback Generator (Zero Excel Dependency)
# ==============================================================================
function Export-AnalysisToOpenXml {
    param(
        [Parameter(Mandatory = $true)]$Analysis,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [string]$DbLabel = '',
        [switch]$DetailedAnalysis,
        [scriptblock]$LogCallback = $null,
        [AuroraAnalyzer.OperationCancelState]$CancelState = $null,
        [scriptblock]$PumpCallback = $null
    )

    if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    function XmlEsc([string]$s) {
        return [AuroraAnalyzer.XmlText]::Escape($s)
    }

    function ColLetter([int]$colIndex1Based) {
        return (ConvertTo-ExcelColLetter $colIndex1Based)
    }

    function Build-SheetXml {
        param(
            [string]$Title,
            [string]$Subtitle,
            [string[]]$Headers,
            [object[]]$Rows,
            [double[]]$ColWidths,
            [hashtable]$NumColStyles = @{},
            [int[]]$CenterCols = @(),
            [int[]]$WrapCols = @(),
            [int[]]$WarnCols = @(),
            [int[]]$WarnPositiveIntCols = @(),
            [int[]]$NavyOptimalCols = @(),
            [bool]$HighlightCatOnWarnCol1 = $false,
            [object[]]$SummaryFooterRow = $null,
            [hashtable]$SummaryFooterStyles = @{},
            [int[]]$SummaryFooterWarnCols = @(),
            [bool]$IncludeChartDrawing = $false
        )
        # Style IDs (paired odd/even for clean edge-to-edge monotone zebra striping):
        # 0: Default | 1: Banner Title (#FFF on #1A2B4C) | 2: Subtitle (#333 on #F2F2F2 + Navy bottom)
        # 3: Table Header (#FFF bold on #1A2B4C)
        # 4/5: Text Left (White / #F7F7F7)
        # 6/7: Int #,##0 (White / #F7F7F7)
        # 8/9: Dec2 #,##0.00 (White / #F7F7F7)
        # 10/11: Pct 0.0% (White / #F7F7F7)
        # 12/13: Dec1 #,##0.0 (White / #F7F7F7)
        # 14/15: Dec3 0.000 (White / #F7F7F7)
        # 16/17: Text Center (White / #F7F7F7)
        # 18/19: Text Wrap Left (White / #F7F7F7)
        # 20: Warning Text Left (#800020 bold on #F5F5F5, wrap)
        # 21: Warning Text Center (#800020 bold on #F5F5F5, wrap)
        # 22: Warning Int (#800020 bold on #F5F5F5, #,##0)
        # 23: Empty placeholder (#595959 center on White)

        $sb = New-Object System.Text.StringBuilder 32768
        [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
        [void]$sb.Append('<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>')
        [void]$sb.Append('<sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="4" topLeftCell="A5" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
        [void]$sb.Append('<sheetFormatPr defaultRowHeight="22"/>')
        [void]$sb.Append('<cols>')
        for ($i = 0; $i -lt $ColWidths.Count; $i++) {
            [void]$sb.AppendFormat([System.Globalization.CultureInfo]::InvariantCulture, '<col min="{0}" max="{0}" width="{1}" customWidth="1"/>', ($i + 1), $ColWidths[$i])
        }
        [void]$sb.Append('</cols><sheetData>')

        $endL = ColLetter $Headers.Count
        # Row 1: Title Banner (styled across all merged cells)
        [void]$sb.Append('<row r="1" ht="32" customHeight="1">')
        [void]$sb.AppendFormat('<c r="A1" t="inlineStr" s="1"><is><t>{0}</t></is></c>', (XmlEsc $Title))
        for ($c = 1; $c -lt $Headers.Count; $c++) {
            [void]$sb.AppendFormat('<c r="{0}1" s="1"/>', (ColLetter ($c + 1)))
        }
        [void]$sb.Append('</row>')

        # Row 2: Subtitle (styled across all merged cells)
        [void]$sb.Append('<row r="2" ht="22" customHeight="1">')
        [void]$sb.AppendFormat('<c r="A2" t="inlineStr" s="2"><is><t>{0}</t></is></c>', (XmlEsc $Subtitle))
        for ($c = 1; $c -lt $Headers.Count; $c++) {
            [void]$sb.AppendFormat('<c r="{0}2" s="2"/>', (ColLetter ($c + 1)))
        }
        [void]$sb.Append('</row>')
        [void]$sb.Append('<row r="3" ht="10" customHeight="1"/>')

        # Row 4: Table Header
        [void]$sb.Append('<row r="4" ht="28" customHeight="1">')
        for ($c = 0; $c -lt $Headers.Count; $c++) {
            $ref = (ColLetter ($c + 1)) + '4'
            [void]$sb.AppendFormat('<c r="{0}" t="inlineStr" s="3"><is><t>{1}</t></is></c>', $ref, (XmlEsc $Headers[$c]))
        }
        [void]$sb.Append('</row>')

        $merges = New-Object System.Collections.Generic.List[string]
        $merges.Add("A1:${endL}1")
        $merges.Add("A2:${endL}2")

        $rNum = 5
        if ($Rows.Count -eq 0) {
            [void]$sb.Append('<row r="5" ht="24" customHeight="1">')
            [void]$sb.Append('<c r="A5" t="inlineStr" s="23"><is><t>該当データなし</t></is></c>')
            for ($c = 1; $c -lt $Headers.Count; $c++) {
                [void]$sb.AppendFormat('<c r="{0}5" s="23"/>', (ColLetter ($c + 1)))
            }
            [void]$sb.Append('</row>')
            $merges.Add("A5:${endL}5")
            $rNum = 6
        }
        else {
            for ($r = 0; $r -lt $Rows.Count; $r++) {
                if (($r % 32) -eq 0) {
                    if ($null -ne $PumpCallback) { & $PumpCallback }
                    if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }
                }
                $rowVals = $Rows[$r]
                $isZebra = (($r % 2) -eq 1)
                $ht = 22.0
                if ($WrapCols.Count -gt 0) {
                    $wTexts = New-Object System.Collections.Generic.List[string]
                    $wWidths = New-Object System.Collections.Generic.List[double]
                    foreach ($wc in $WrapCols) {
                        $wTexts.Add([string]$rowVals[$wc - 1])
                        $wWidths.Add($(if (($wc - 1) -lt $ColWidths.Count) { [double]$ColWidths[$wc - 1] } else { 24.0 }))
                    }
                    $ht = Get-EstimatedRowHeight -Texts $wTexts.ToArray() -ColWidths $wWidths.ToArray() -MinHeight 22.0 -MaxHeight 110.0
                }
                [void]$sb.AppendFormat([System.Globalization.CultureInfo]::InvariantCulture, '<row r="{0}" ht="{1:0.#}" customHeight="1">', $rNum, $ht)

                for ($c = 0; $c -lt $Headers.Count; $c++) {
                    $ref = (ColLetter ($c + 1)) + $rNum
                    $v = $rowVals[$c]
                    $col1 = $c + 1
                    $isWarn = $false
                    if ($WarnCols -contains $col1) {
                        if (Test-WarningCellText -Text ([string]$v)) { $isWarn = $true }
                    }
                    if (-not $isWarn -and $HighlightCatOnWarnCol1 -and $col1 -eq 2 -and $rowVals.Count -gt 0) {
                        if (Test-WarningCellText -Text ([string]$rowVals[0])) { $isWarn = $true }
                    }
                    if (-not $isWarn -and ($WarnPositiveIntCols -contains $col1) -and ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal])) {
                        if ([double]$v -gt 0) { $isWarn = $true }
                    }

                    if ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal]) {
                        if ($isWarn) {
                            $st = 22
                        }
                        elseif ($NumColStyles.ContainsKey($col1)) {
                            $baseSt = [int]$NumColStyles[$col1]
                            $st = if ($isZebra) { $baseSt + 1 } else { $baseSt }
                        }
                        else {
                            $st = if ($isZebra) { 5 } else { 4 }
                        }
                        [void]$sb.AppendFormat([System.Globalization.CultureInfo]::InvariantCulture, '<c r="{0}" s="{1}"><v>{2}</v></c>', $ref, $st, $v)
                    }
                    else {
                        $sv = [string]$v
                        if ($isWarn) {
                            $st = if ($CenterCols -contains $col1) { 21 } else { 20 }
                        }
                        elseif (($NavyOptimalCols -contains $col1) -and $sv -like '最適*') {
                            $st = if ($isZebra) { 38 } else { 37 }
                        }
                        elseif ($WrapCols -contains $col1) {
                            $st = if ($isZebra) { 19 } else { 18 }
                        }
                        elseif ($CenterCols -contains $col1) {
                            $st = if ($isZebra) { 17 } else { 16 }
                        }
                        else {
                            $st = if ($isZebra) { 5 } else { 4 }
                        }
                        [void]$sb.AppendFormat('<c r="{0}" t="inlineStr" s="{1}"><is><t>{2}</t></is></c>', $ref, $st, (XmlEsc $sv))
                    }
                }
                [void]$sb.Append('</row>')
                $rNum++
            }
        }

        $filterLast = [Math]::Max(4, $rNum - 1)
        if ($null -ne $SummaryFooterRow -and $Rows.Count -gt 0) {
            $rNum++
            [void]$sb.AppendFormat('<row r="{0}" ht="25" customHeight="1">', $rNum)
            for ($c = 0; $c -lt $Headers.Count; $c++) {
                $ref = (ColLetter ($c + 1)) + $rNum
                $v = $SummaryFooterRow[$c]
                $col1 = $c + 1
                if ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal]) {
                    $st = if ($SummaryFooterStyles.ContainsKey($col1)) { [int]$SummaryFooterStyles[$col1] } else { 30 }
                    if (($SummaryFooterWarnCols -contains $col1) -and [double]$v -gt 0) { $st = 35 }
                    [void]$sb.AppendFormat([System.Globalization.CultureInfo]::InvariantCulture, '<c r="{0}" s="{1}"><v>{2}</v></c>', $ref, $st, $v)
                }
                else {
                    $st = if ($CenterCols -contains $col1) { 34 } else { 29 }
                    [void]$sb.AppendFormat('<c r="{0}" t="inlineStr" s="{1}"><is><t>{2}</t></is></c>', $ref, $st, (XmlEsc ([string]$v)))
                }
            }
            [void]$sb.Append('</row>')
            $rNum++
        }

        [void]$sb.Append('</sheetData>')
        [void]$sb.AppendFormat('<autoFilter ref="A4:{0}{1}"/>', $endL, $filterLast)
        [void]$sb.AppendFormat('<mergeCells count="{0}">', $merges.Count)
        foreach ($mg in $merges) { [void]$sb.AppendFormat('<mergeCell ref="{0}"/>', $mg) }
        [void]$sb.Append('</mergeCells>')
        [void]$sb.Append('<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>')
        [void]$sb.Append('<pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/>')
        if ($IncludeChartDrawing) {
            [void]$sb.Append('<drawing r:id="rId1"/>')
        }
        [void]$sb.Append('</worksheet>')
        return $sb.ToString()
    }

    function Build-SummarySheetXml {
        param(
            [string]$Title,
            [string]$Subtitle,
            $AnalysisObj,
            $RoleSummaryItems,
            $TypeSummaryItems,
            [string]$SizeKpiLbl,
            [string]$SizeColHdr,
            [bool]$IsDetailed,
            [ref]$OutTopStartRow,
            [ref]$OutTopEndRow,
            [ref]$OutLastRow
        )

        $colW = @(6, 24, 17, 13, 11, 8, 15, 11, 14, 11, 8, 32)
        $endL = ColLetter 12
        $sb = New-Object System.Text.StringBuilder 65536
        $merges = New-Object System.Collections.Generic.List[string]
        [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
        [void]$sb.Append('<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>')
        [void]$sb.Append('<sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>')
        [void]$sb.Append('<sheetFormatPr defaultRowHeight="22"/>')
        [void]$sb.Append('<cols>')
        for ($i = 0; $i -lt $colW.Count; $i++) {
            [void]$sb.AppendFormat([System.Globalization.CultureInfo]::InvariantCulture, '<col min="{0}" max="{0}" width="{1}" customWidth="1"/>', ($i + 1), $colW[$i])
        }
        [void]$sb.Append('</cols><sheetData>')

        [void]$sb.Append('<row r="1" ht="32" customHeight="1">')
        [void]$sb.AppendFormat('<c r="A1" t="inlineStr" s="1"><is><t>{0}</t></is></c>', (XmlEsc $Title))
        for ($c = 1; $c -lt 12; $c++) { [void]$sb.AppendFormat('<c r="{0}1" s="1"/>', (ColLetter ($c + 1))) }
        [void]$sb.Append('</row>')
        $merges.Add("A1:${endL}1")

        [void]$sb.Append('<row r="2" ht="22" customHeight="1">')
        [void]$sb.AppendFormat('<c r="A2" t="inlineStr" s="2"><is><t>{0}</t></is></c>', (XmlEsc $Subtitle))
        for ($c = 1; $c -lt 12; $c++) { [void]$sb.AppendFormat('<c r="{0}2" s="2"/>', (ColLetter ($c + 1))) }
        [void]$sb.Append('</row>')
        $merges.Add("A2:${endL}2")
        [void]$sb.Append('<row r="3" ht="10" customHeight="1"/>')

        function Add-XmlKpiBand([int]$rLbl, [object[]]$kpis) {
            [void]$sb.AppendFormat('<row r="{0}" ht="20" customHeight="1">', $rLbl)
            foreach ($k in $kpis) {
                $st = if ($k.Warn) { 27 } else { 25 }
                $c1L = ColLetter $k.Col1; $c2L = ColLetter $k.Col2
                [void]$sb.AppendFormat('<c r="{0}{1}" t="inlineStr" s="{2}"><is><t>{3}</t></is></c><c r="{4}{1}" s="{2}"/>', $c1L, $rLbl, $st, (XmlEsc $k.Label), $c2L)
                $merges.Add("${c1L}${rLbl}:${c2L}${rLbl}")
            }
            [void]$sb.Append('</row>')
            $rVal = $rLbl + 1
            [void]$sb.AppendFormat('<row r="{0}" ht="30" customHeight="1">', $rVal)
            foreach ($k in $kpis) {
                $st = if ($k.Warn) { 28 } else { 26 }
                $c1L = ColLetter $k.Col1; $c2L = ColLetter $k.Col2
                [void]$sb.AppendFormat('<c r="{0}{1}" t="inlineStr" s="{2}"><is><t>{3}</t></is></c><c r="{4}{1}" s="{2}"/>', $c1L, $rVal, $st, (XmlEsc ([string]$k.Display)), $c2L)
                $merges.Add("${c1L}${rVal}:${c2L}${rVal}")
            }
            [void]$sb.Append('</row>')
        }

        function Add-XmlSecHdr([int]$r, [string]$txt) {
            [void]$sb.AppendFormat('<row r="{0}" ht="25" customHeight="1">', $r)
            [void]$sb.AppendFormat('<c r="A{0}" t="inlineStr" s="24"><is><t>{1}</t></is></c>', $r, (XmlEsc $txt))
            for ($c = 1; $c -lt 12; $c++) { [void]$sb.AppendFormat('<c r="{0}{1}" s="24"/>', (ColLetter ($c + 1)), $r) }
            [void]$sb.Append('</row>')
            $merges.Add("A${r}:${endL}${r}")
        }

        function Add-XmlTblHdr([int]$r, [string[]]$hdrs) {
            [void]$sb.AppendFormat('<row r="{0}" ht="26" customHeight="1">', $r)
            for ($c = 0; $c -lt $hdrs.Count; $c++) {
                [void]$sb.AppendFormat('<c r="{0}{1}" t="inlineStr" s="3"><is><t>{2}</t></is></c>', (ColLetter ($c + 1)), $r, (XmlEsc $hdrs[$c]))
            }
            [void]$sb.Append('</row>')
        }

        function Add-XmlDataRow([int]$r, [object[]]$vals, [bool]$isZebra, [hashtable]$numSt, [int[]]$cCols, [int[]]$wCols, [int[]]$warnCols, [int[]]$warnPosCols) {
            $ht = 22.0
            if ($wCols.Count -gt 0) {
                $wt = @(); $ww = @()
                foreach ($wc in $wCols) { $wt += [string]$vals[$wc - 1]; $ww += [double]$colW[$wc - 1] }
                $ht = Get-EstimatedRowHeight -Texts $wt -ColWidths $ww -MinHeight 22.0 -MaxHeight 90.0
            }
            [void]$sb.AppendFormat([System.Globalization.CultureInfo]::InvariantCulture, '<row r="{0}" ht="{1:0.#}" customHeight="1">', $r, $ht)
            for ($c = 0; $c -lt $vals.Count; $c++) {
                $col1 = $c + 1; $ref = (ColLetter $col1) + $r; $v = $vals[$c]
                $isW = (($warnCols -contains $col1) -and (Test-WarningCellText -Text ([string]$v))) -or
                       (($warnPosCols -contains $col1) -and ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal]) -and [double]$v -gt 0)
                if ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal]) {
                    $st = if ($isW) { 22 } elseif ($numSt.ContainsKey($col1)) { $b = [int]$numSt[$col1]; if ($isZebra) { $b + 1 } else { $b } } else { if ($isZebra) { 5 } else { 4 } }
                    [void]$sb.AppendFormat([System.Globalization.CultureInfo]::InvariantCulture, '<c r="{0}" s="{1}"><v>{2}</v></c>', $ref, $st, $v)
                }
                else {
                    $st = if ($isW) { if ($cCols -contains $col1) { 21 } else { 20 } } elseif ($wCols -contains $col1) { if ($isZebra) { 19 } else { 18 } } elseif ($cCols -contains $col1) { if ($isZebra) { 17 } else { 16 } } else { if ($isZebra) { 5 } else { 4 } }
                    [void]$sb.AppendFormat('<c r="{0}" t="inlineStr" s="{1}"><is><t>{2}</t></is></c>', $ref, $st, (XmlEsc ([string]$v)))
                }
            }
            [void]$sb.Append('</row>')
        }

        function Add-XmlTotalRow([int]$r, [object[]]$vals, [hashtable]$totSt, [int[]]$cCols, [int[]]$warnPosCols) {
            [void]$sb.AppendFormat('<row r="{0}" ht="25" customHeight="1">', $r)
            for ($c = 0; $c -lt $vals.Count; $c++) {
                $col1 = $c + 1; $ref = (ColLetter $col1) + $r; $v = $vals[$c]
                if ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal]) {
                    $st = if (($warnPosCols -contains $col1) -and [double]$v -gt 0) { 35 } elseif ($totSt.ContainsKey($col1)) { [int]$totSt[$col1] } else { 30 }
                    [void]$sb.AppendFormat([System.Globalization.CultureInfo]::InvariantCulture, '<c r="{0}" s="{1}"><v>{2}</v></c>', $ref, $st, $v)
                }
                else {
                    $st = if ($cCols -contains $col1) { 34 } else { 29 }
                    [void]$sb.AppendFormat('<c r="{0}" t="inlineStr" s="{1}"><is><t>{2}</t></is></c>', $ref, $st, (XmlEsc ([string]$v)))
                }
            }
            [void]$sb.Append('</row>')
        }

        function Add-XmlMergedRow([int]$r, [object[]]$spans, [object[]]$vals, [int[]]$styles, [double]$ht) {
            [void]$sb.AppendFormat([System.Globalization.CultureInfo]::InvariantCulture, '<row r="{0}" ht="{1:0.#}" customHeight="1">', $r, $ht)
            for ($k = 0; $k -lt $spans.Count; $k++) {
                $sp = $spans[$k]; $v = $vals[$k]; $st = $styles[$k]
                $c1L = ColLetter $sp.C1; $c2L = ColLetter $sp.C2
                if ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal]) {
                    [void]$sb.AppendFormat([System.Globalization.CultureInfo]::InvariantCulture, '<c r="{0}{1}" s="{2}"><v>{3}</v></c>', $c1L, $r, $st, $v)
                }
                else {
                    [void]$sb.AppendFormat('<c r="{0}{1}" t="inlineStr" s="{2}"><is><t>{3}</t></is></c>', $c1L, $r, $st, (XmlEsc ([string]$v)))
                }
                for ($cc = $sp.C1 + 1; $cc -le $sp.C2; $cc++) {
                    [void]$sb.AppendFormat('<c r="{0}{1}" s="{2}"/>', (ColLetter $cc), $r, $st)
                }
                if ($sp.C2 -gt $sp.C1) { $merges.Add("${c1L}${r}:${c2L}${r}") }
            }
            [void]$sb.Append('</row>')
        }

        Add-XmlKpiBand 4 @(
            @{ Col1 = 1;  Col2 = 2;  Label = '総テーブル数';           Display = ("{0:#,##0} テーブル" -f $AnalysisObj.TableCount); Warn = $false },
            @{ Col1 = 3;  Col2 = 4;  Label = '総レコード数';           Display = ("{0:#,##0} 件" -f $AnalysisObj.TotalRows); Warn = ($AnalysisObj.TotalRows -eq 0) },
            @{ Col1 = 5;  Col2 = 6;  Label = $SizeKpiLbl;              Display = ("{0:#,##0.00} MB" -f ($AnalysisObj.TotalFileBytes / 1MB)); Warn = $false },
            @{ Col1 = 7;  Col2 = 8;  Label = '推定InnoDB容量 (MB)';    Display = ("{0:#,##0.00} MB" -f ($AnalysisObj.TotalEstInnoDbBytes / 1MB)); Warn = $false },
            @{ Col1 = 9;  Col2 = 10; Label = '総カラム数 / PK候補';    Display = ("{0} 列 / PK {1} 列" -f $AnalysisObj.TotalColumns, $AnalysisObj.CandidatePkColumnCount); Warn = $false },
            @{ Col1 = 11; Col2 = 12; Label = '検出警告 (高/中/低)';    Display = ("高 {0} / 中 {1} / 低 {2} (計 {3})" -f $AnalysisObj.HighWarningCount, $AnalysisObj.MediumWarningCount, $AnalysisObj.LowWarningCount, $AnalysisObj.TotalWarningCount); Warn = ($AnalysisObj.TotalWarningCount -gt 0) }
        )
        [void]$sb.Append('<row r="6" ht="8" customHeight="1"/>')
        Add-XmlKpiBand 7 @(
            @{ Col1 = 1;  Col2 = 2;  Label = '全体セルNULL率';            Display = ("{0:0.0%}" -f $AnalysisObj.OverallAvgNullRate); Warn = ($AnalysisObj.OverallAvgNullRate -ge 0.30) },
            @{ Col1 = 3;  Col2 = 4;  Label = '平均行長 / 最大行長';       Display = ("{0:#,##0.0} B / {1:#,##0} B" -f $AnalysisObj.OverallAvgRowBytes, $AnalysisObj.OverallMaxRowBytes); Warn = $false },
            @{ Col1 = 5;  Col2 = 6;  Label = '偏り高リスク / 大容量列';   Display = ("偏り {0} 列 / 大容量 {1} 列" -f $AnalysisObj.HighSkewColumnCount, $AnalysisObj.LargeColumnCount); Warn = ($AnalysisObj.HighSkewColumnCount -gt 0 -or $AnalysisObj.LargeColumnCount -gt 0) },
            @{ Col1 = 7;  Col2 = 8;  Label = '検出リレーション / 孤立行'; Display = ("{0} 件 / 孤立 {1:#,##0} 行" -f $AnalysisObj.Relationships.Count, $AnalysisObj.TotalOrphanRows); Warn = ($AnalysisObj.TotalOrphanRows -gt 0) },
            @{ Col1 = 9;  Col2 = 10; Label = '空テーブル / 書式異常行';   Display = ("空 {0} 表 / 異常 {1:#,##0} 行" -f $AnalysisObj.EmptyTableCount, $AnalysisObj.RaggedRowTotal); Warn = ($AnalysisObj.EmptyTableCount -gt 0 -or $AnalysisObj.RaggedRowTotal -gt 0) },
            @{ Col1 = 11; Col2 = 12; Label = '推定ACU目安 (現状 / 10倍)'; Display = ("{0:0.0} ACU / {1:0.0} ACU" -f $AnalysisObj.EstMinAcuCurrent, $AnalysisObj.EstMinAcu10x); Warn = $false }
        )
        [void]$sb.Append('<row r="9" ht="10" customHeight="1"/>')

        $r = 10
        Add-XmlSecHdr $r '1. テーブル別レコード数・容量構成比 TOP15 (性能試験データ規模)'; $r++
        Add-XmlTblHdr $r @('順位','テーブル名','推定役割','レコード数','件数構成比','列数',$SizeColHdr,'容量構成比','推定InnoDB(KB)','平均NULL率','警告数','健全性・注意点'); $r++
        $topList = @($AnalysisObj.Tables | Select-Object -First 15)
        $OutTopStartRow.Value = $r
        for ($i = 0; $i -lt $topList.Count; $i++) {
            $t = $topList[$i]
            Add-XmlDataRow $r @(($i + 1), $t.TableName, $t.InferredRole, [double]$t.RowCount, [double]$t.RowShare, [double]$t.ColumnCount, [Math]::Round($t.FileSizeBytes / 1024.0, 2), [double]$t.SizeShare, [Math]::Round($t.EstInnoDbBytes / 1024.0, 1), [double]$t.AvgNullRate, [double]$t.WarningCount, $(if ($t.WarningSummary) { $t.WarningSummary } else { '正常' })) (($i % 2) -eq 1) @{1=6;4=6;5=10;6=6;7=8;8=10;9=12;10=10;11=6} @(1,3,6,11) @(12) @(12) @(11)
            $r++
        }
        $OutTopEndRow.Value = $r - 1
        $totRowLabel = if ($AnalysisObj.TableCount -gt 15) { "全$($AnalysisObj.TableCount)テーブル合計" } else { "全体合計 ($($AnalysisObj.TableCount)テーブル)" }
        Add-XmlTotalRow $r @('計', $totRowLabel, '-', [double]$AnalysisObj.TotalRows, 1.0, [double]$AnalysisObj.TotalColumns, [Math]::Round($AnalysisObj.TotalFileBytes / 1024.0, 2), 1.0, [Math]::Round($AnalysisObj.TotalEstInnoDbBytes / 1024.0, 1), [double]$AnalysisObj.OverallAvgNullRate, [double]$AnalysisObj.TotalWarningCount, ("高 {0} / 中 {1} / 低 {2}" -f $AnalysisObj.HighWarningCount, $AnalysisObj.MediumWarningCount, $AnalysisObj.LowWarningCount)) @{4=30;5=32;6=30;7=31;8=32;9=33;10=32;11=30} @(1,3,6,11) @(11)
        $r += 2

        Add-XmlSecHdr $r '2. 推定役割別テーブル構成サマリー (データ分類・規模内訳)'; $r++
        Add-XmlTblHdr $r @('No','推定役割','テーブル数','テーブル比','合計レコード数','件数構成比','合計列数',$SizeColHdr,'容量構成比','推定InnoDB(KB)','警告数','該当テーブル例 (上位)'); $r++
        for ($i = 0; $i -lt $RoleSummaryItems.Count; $i++) {
            $rg = $RoleSummaryItems[$i]
            $tblSh = if ($AnalysisObj.TableCount -gt 0) { [double]$rg.TableCount / $AnalysisObj.TableCount } else { 0.0 }
            $rowSh = if ($AnalysisObj.TotalRows -gt 0) { [double]$rg.TotalRows / $AnalysisObj.TotalRows } else { 0.0 }
            $szSh  = if ($AnalysisObj.TotalFileBytes -gt 0) { [double]$rg.TotalFileBytes / $AnalysisObj.TotalFileBytes } else { 0.0 }
            $exTxt = ($rg.Examples -join ', ') + $(if ($rg.TableCount -gt $rg.Examples.Count) { " 他$($rg.TableCount - $rg.Examples.Count)" } else { '' })
            Add-XmlDataRow $r @(($i + 1), $rg.Role, [double]$rg.TableCount, $tblSh, [double]$rg.TotalRows, $rowSh, [double]$rg.TotalCols, [Math]::Round($rg.TotalFileBytes / 1024.0, 2), $szSh, [Math]::Round($rg.TotalInnoBytes / 1024.0, 1), [double]$rg.WarningCount, $exTxt) (($i % 2) -eq 1) @{1=6;3=6;4=10;5=6;6=10;7=6;8=8;9=10;10=12;11=6} @(1,3,7,11) @(12) @(2) @(11)
            $r++
        }
        Add-XmlTotalRow $r @('計', '全役割合計', [double]$AnalysisObj.TableCount, 1.0, [double]$AnalysisObj.TotalRows, 1.0, [double]$AnalysisObj.TotalColumns, [Math]::Round($AnalysisObj.TotalFileBytes / 1024.0, 2), 1.0, [Math]::Round($AnalysisObj.TotalEstInnoDbBytes / 1024.0, 1), [double]$AnalysisObj.TotalWarningCount, "空テーブル $($AnalysisObj.EmptyTableCount) 件") @{3=30;4=32;5=30;6=32;7=30;8=31;9=32;10=33;11=30} @(1,3,7,11) @(11)
        $r += 2

        Add-XmlSecHdr $r '3. データ型別カラム構成 ＆ 品質サマリー (全カラム内訳)'; $r++
        Add-XmlTblHdr $r @('No','推定データ型','カラム数','構成比','PK候補','区分列','平均NULL率','平均ユニーク率','平均長(B)','最大長(B)','偏り/全NULL','代表カラム例'); $r++
        $totCD = [Math]::Max(1, $AnalysisObj.TotalColumns)
        for ($i = 0; $i -lt $TypeSummaryItems.Count; $i++) {
            $tg = $TypeSummaryItems[$i]
            $avgN = if ($tg.ColCount -gt 0) { $tg.NullRateSum / $tg.ColCount } else { 0.0 }
            $avgC = if ($tg.ColCount -gt 0) { $tg.CardRateSum / $tg.ColCount } else { 0.0 }
            $avgB = if ($tg.NonEmptyCols -gt 0) { [Math]::Round($tg.AvgByteSum / $tg.NonEmptyCols, 1) } else { 0.0 }
            $rTxt = if ($tg.HighSkewCnt -gt 0 -or $tg.AllNullCnt -gt 0) { "偏り:$($tg.HighSkewCnt) / 全NULL:$($tg.AllNullCnt)" } else { '正常' }
            $exTxt = ($tg.Examples -join ', ') + $(if ($tg.ColCount -gt $tg.Examples.Count) { " 他$($tg.ColCount - $tg.Examples.Count)" } else { '' })
            Add-XmlDataRow $r @(($i + 1), $tg.DataType, [double]$tg.ColCount, [double]($tg.ColCount / $totCD), [double]$tg.PkCount, [double]$tg.CatCount, $avgN, $avgC, $avgB, [double]$tg.MaxByteLen, $rTxt, $exTxt) (($i % 2) -eq 1) @{1=6;3=6;4=10;5=6;6=6;7=10;8=10;9=12;10=6} @(1,3,5,6,11) @(11,12) @(11) @()
            $r++
        }
        $r++

        $topW = @($AnalysisObj.Warnings | Select-Object -First 20)
        $wSpans = @(@{C1=1;C2=1}, @{C1=2;C2=2}, @{C1=3;C2=3}, @{C1=4;C2=4}, @{C1=5;C2=7}, @{C1=8;C2=12})
        Add-XmlSecHdr $r ("4. 検出された主要警告・リスクサマリー (上位{0}件 / 全{1}件)" -f $topW.Count, $AnalysisObj.Warnings.Count); $r++
        Add-XmlMergedRow $r $wSpans @('重要度','カテゴリ','テーブル名','カラム名','検出値・指標','Aurora / 性能試験への影響 ＆ 推奨アクション') @(3,3,3,3,3,3) 26; $r++
        if ($topW.Count -eq 0) {
            Add-XmlMergedRow $r @(@{C1=1;C2=12}) @('警告・注意項目は検出されませんでした (良好)') @(23) 24; $r++
        }
        else {
            for ($i = 0; $i -lt $topW.Count; $i++) {
                $w = $topW[$i]; $z = (($i % 2) -eq 1)
                $actTxt = "【影響】$($w.PerfImpact)  【対策】$($w.Recommendation)"
                $isHM = ($w.Severity -eq '高' -or $w.Severity -eq '中')
                $s0 = if ($isHM) { 21 } else { if ($z) { 17 } else { 16 } }
                $s2 = if ($z) { 5 } else { 4 }
                $sW = if ($z) { 19 } else { 18 }
                $ht = Get-EstimatedRowHeight -Texts @($w.MetricValue, $actTxt) -ColWidths @(34.0, 76.0) -MinHeight 24.0 -MaxHeight 85.0
                Add-XmlMergedRow $r $wSpans @($w.Severity, $w.Category, $w.TableName, $w.ColumnName, $w.MetricValue, $actTxt) @($s0, $s0, $s2, $s2, $sW, $sW) $ht; $r++
            }
        }
        $r++

        Add-XmlSecHdr $r '5. テーブル間リレーション概要 ＆ JMeterパラメータ化適性ハイライト'; $r++
        $rSpans = @(@{C1=1;C2=1}, @{C1=2;C2=3}, @{C1=4;C2=4}, @{C1=5;C2=5}, @{C1=6;C2=7}, @{C1=8;C2=8}, @{C1=9;C2=10}, @{C1=11;C2=12})
        Add-XmlMergedRow $r $rSpans @('No','親テーブル.PK → 子テーブル.FK','参照一致率','孤立行数','平均 / P95 (子/親)','最大(子/親)','リスク評価','所見・スーパーユーザー情報') @(3,3,3,3,3,3,3,3) 25; $r++
        $topR = @($AnalysisObj.Relationships | Select-Object -First 10)
        if ($topR.Count -eq 0) {
            Add-XmlMergedRow $r @(@{C1=1;C2=12}) @('自動推定されたテーブル間FKリレーションはありません') @(23) 24; $r++
        }
        else {
            for ($i = 0; $i -lt $topR.Count; $i++) {
                $rl = $topR[$i]; $z = (($i % 2) -eq 1)
                $sT = if ($z) { 5 } else { 4 }; $sI = if ($z) { 7 } else { 6 }; $sP = if ($z) { 11 } else { 10 }; $sC = if ($z) { 17 } else { 16 }; $sW = if ($z) { 19 } else { 18 }
                $sOrp = if ($rl.OrphanRows -gt 0) { 22 } else { $sI }
                $sRsk = if (Test-WarningCellText -Text $rl.RiskLevel) { 21 } else { $sC }
                $sNt  = if (Test-WarningCellText -Text $rl.Note) { 20 } else { $sW }
                $ht = Get-EstimatedRowHeight -Texts @($rl.Note) -ColWidths @(40.0) -MinHeight 22.0 -MaxHeight 65.0
                Add-XmlMergedRow $r $rSpans @([double]($i + 1), "$($rl.ParentTable).$($rl.ParentColumn) → $($rl.ChildTable).$($rl.ChildColumn)", [double]$rl.MatchRate, [double]$rl.OrphanRows, ("平均 {0:0.00} / P95 {1:0.0}" -f $rl.FanOutAvg, $rl.FanOutP95), [double]$rl.FanOutMax, $rl.RiskLevel, $rl.Note) @($sI, $sT, $sP, $sOrp, $sC, $sI, $sRsk, $sNt) $ht; $r++
            }
        }
        $r++

        $jSpans = @(@{C1=1;C2=1}, @{C1=2;C2=2}, @{C1=3;C2=3}, @{C1=4;C2=4}, @{C1=5;C2=5}, @{C1=6;C2=7}, @{C1=8;C2=12})
        Add-XmlMergedRow $r $jSpans @('No','テーブル名','キーカラム','データ型','一意キー数','JMeter並列適性','推奨設定・活用指針') @(3,3,3,3,3,3,3) 25; $r++
        $topJ = @($AnalysisObj.JMeterCandidates | Select-Object -First 8)
        if ($topJ.Count -eq 0) {
            Add-XmlMergedRow $r @(@{C1=1;C2=12}) @('該当するJMeter一意キー候補 (10件以上) はありません') @(23) 24; $r++
        }
        else {
            for ($i = 0; $i -lt $topJ.Count; $i++) {
                $jc = $topJ[$i]; $z = (($i % 2) -eq 1)
                $sT = if ($z) { 5 } else { 4 }; $sI = if ($z) { 7 } else { 6 }; $sC = if ($z) { 17 } else { 16 }; $sW = if ($z) { 19 } else { 18 }
                $sSuit = if ($jc.Suitability -like '最適*') { if ($z) { 38 } else { 37 } } elseif (Test-WarningCellText -Text $jc.Suitability) { 21 } else { $sC }
                $ht = Get-EstimatedRowHeight -Texts @($jc.RecommendedUsage) -ColWidths @(76.0) -MinHeight 22.0 -MaxHeight 65.0
                Add-XmlMergedRow $r $jSpans @([double]($i + 1), $jc.TableName, $jc.KeyColumn, $jc.DataType, [double]$jc.UniqueCount, $jc.Suitability, $jc.RecommendedUsage) @($sI, $sT, $sT, $sC, $sI, $sSuit, $sW) $ht; $r++
            }
        }
        $r++

        Add-XmlSecHdr $r '6. Aurora Serverless v2 (MySQL 8.4.7) 性能試験向け総合所見'; $r++
        foreach ($ins in $AnalysisObj.AuroraInsights) {
            $ht = Get-EstimatedRowHeight -Texts @($ins) -ColWidths @(165.0) -MinHeight 26.0 -MaxHeight 85.0
            Add-XmlMergedRow $r @(@{C1=1;C2=12}) @($ins) @(36) $ht; $r++
        }
        if (-not $IsDetailed) {
            Add-XmlMergedRow $r @(@{C1=1;C2=12}) @('※ カテゴリ分布・数値統計ヒストグラム・時系列分析・リレーション詳細・性能試験ガイド全件シートを出力するには、ダイアログの「詳細分析」にチェックを入れて実行してください。') @(23) 22; $r++
        }

        $OutLastRow.Value = $r - 1
        [void]$sb.Append('</sheetData>')
        [void]$sb.AppendFormat('<mergeCells count="{0}">', $merges.Count)
        foreach ($mg in $merges) { [void]$sb.AppendFormat('<mergeCell ref="{0}"/>', $mg) }
        [void]$sb.Append('</mergeCells>')
        [void]$sb.Append('<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>')
        [void]$sb.Append('<pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/>')
        if ($topList.Count -gt 0) { [void]$sb.Append('<drawing r:id="rId1"/>') }
        [void]$sb.Append('</worksheet>')
        return $sb.ToString()
    }

    $OutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
    $stagingPath = New-AnalysisStagingPath $OutputPath
    $modeSheetText = if ($DetailedAnalysis) { '詳細分析8シート' } else { '標準3シート' }
    if ($null -ne $LogCallback) { & $LogCallback "Excelレポート生成中 (OpenXML / $modeSheetText)" 60 }

    $isDirectDb = [bool]$Analysis.IsDirectDb
    $dbLabelPart = if ([string]::IsNullOrWhiteSpace($DbLabel)) { '' } else { " | A5M2登録DB: $DbLabel" }
    $subInfo = if ($isDirectDb) {
        "分析日時: {0} | 接続先DB: {1}{2} | 取得方式: DB直接分析(CSV未出力) | 出力構成: {3} | Aurora Serverless v2 (MySQL 8.4.7)" -f $Analysis.AnalyzedAt.ToString('yyyy/MM/dd HH:mm:ss'), $Analysis.InputFolder, $dbLabelPart, $modeSheetText
    }
    else {
        "分析日時: {0} | 対象フォルダ: {1}{2} | 出力構成: {3} | Aurora Serverless v2 (MySQL 8.4.7)" -f $Analysis.AnalyzedAt.ToString('yyyy/MM/dd HH:mm:ss'), $Analysis.InputFolder, $dbLabelPart, $modeSheetText
    }
    $sizeKpiLabel = if ($isDirectDb) { 'CSV換算サイズ (MB)' } else { '総CSVサイズ (MB)' }
    $sizeColHeader = if ($isDirectDb) { 'CSV換算サイズ(KB)' } else { 'CSVサイズ(KB)' }
    $srcColHeader = if ($isDirectDb) { '取得元 (DB.テーブル)' } else { 'CSVファイル名' }
    $tblSheetTitle = if ($isDirectDb) { 'テーブル別詳細一覧 (DB直接分析・CSV未出力)' } else { 'テーブル別詳細一覧 (全CSVファイルプロファイル)' }

    $allCols = @($Analysis.Tables | ForEach-Object { $_.Columns })
    $roleSummary = @(Get-RoleSummaryItems -Analysis $Analysis)
    $typeSummary = @(Get-TypeSummaryItems -AllCols $allCols)

    $s2Rows = @()
    $rk = 0
    $minRowNonEmpty = [long]::MaxValue
    foreach ($t in $Analysis.Tables) {
        $rk++
        if ($t.RowCount -gt 0 -and $t.MinRowBytes -lt $minRowNonEmpty) { $minRowNonEmpty = $t.MinRowBytes }
        $s2Rows += , @(
            $rk, $t.TableName, $t.FileName, $t.EncodingName, $t.InferredRole,
            [double]$t.RowCount, [double]$t.RowShare, [double]$t.ColumnCount,
            $(if ($t.ColumnTypeBreakdown) { $t.ColumnTypeBreakdown } else { '-' }),
            [Math]::Round($t.FileSizeBytes / 1024.0, 2), [double]$t.SizeShare,
            [double]$t.MinRowBytes, [Math]::Round($t.AvgRowBytes, 1), [double]$t.MaxRowBytes,
            [Math]::Round($t.EstInnoDbBytes / 1024.0, 1), [double]$t.EstInnoDbPages, [Math]::Round($t.RowsPerInnoDbPage, 1),
            [double]$t.AvgNullRate, [double]$t.AvgCardinalityRate,
            $(if ($t.QualityRiskBreakdown) { $t.QualityRiskBreakdown } else { '問題なし' }),
            [double]$t.RaggedRowCount,
            $(if ($t.CandidatePks) { $t.CandidatePks } else { '-' }),
            $(if ($t.CategoricalColumns) { $t.CategoricalColumns } else { '-' }),
            $(if ($t.DateColumns) { $t.DateColumns } else { '-' }),
            $(if ($t.PrimaryDateRange) { $t.PrimaryDateRange } else { '-' }),
            $(if ($t.RelationSummary) { $t.RelationSummary } else { '-' }),
            $(if ($t.JMeterReadiness) { $t.JMeterReadiness } else { '対象外' }),
            [double]$t.WarningCount,
            $(if ($t.WarningSummary) { $t.WarningSummary } else { '正常' })
        )
    }
    $totMinRowB = if ($minRowNonEmpty -eq [long]::MaxValue) { 0L } else { $minRowNonEmpty }
    $totRowsPerPg = if ($Analysis.TotalEstInnoDbPages -gt 0) { [Math]::Round($Analysis.TotalRows / [double]$Analysis.TotalEstInnoDbPages, 1) } else { 0.0 }
    $avgTblCard = if ($Analysis.Tables.Count -gt 0) { ($Analysis.Tables | Measure-Object -Property AvgCardinalityRate -Average).Average } else { 0.0 }
    $s2Footer = @('合計', "全 $($Analysis.TableCount) テーブル集計", '-', '-', "空 $($Analysis.EmptyTableCount) 表", [double]$Analysis.TotalRows, 1.0, [double]$Analysis.TotalColumns, "PK候補 $($Analysis.CandidatePkColumnCount) 列", [Math]::Round($Analysis.TotalFileBytes / 1024.0, 2), 1.0, [double]$totMinRowB, [Math]::Round($Analysis.OverallAvgRowBytes, 1), [double]$Analysis.OverallMaxRowBytes, [Math]::Round($Analysis.TotalEstInnoDbBytes / 1024.0, 1), [double]$Analysis.TotalEstInnoDbPages, $totRowsPerPg, [double]$Analysis.OverallAvgNullRate, [double]$avgTblCard, "全NULL:$($Analysis.AllNullColumnCount) / 偏り:$($Analysis.HighSkewColumnCount) / 大容量:$($Analysis.LargeColumnCount)", [double]$Analysis.RaggedRowTotal, '-', '-', '-', '-', "関連 $($Analysis.Relationships.Count) 件 / 孤立 $($Analysis.TotalOrphanRows)", "候補 $($Analysis.JMeterCandidates.Count) 列", [double]$Analysis.TotalWarningCount, "高 $($Analysis.HighWarningCount) / 中 $($Analysis.MediumWarningCount) / 低 $($Analysis.LowWarningCount)")

    $s3Rows = @()
    foreach ($c in $allCols) {
        $topVal = if ($c.Top1Value.Length -gt 35) { $c.Top1Value.Substring(0, 32) + '...' } else { $c.Top1Value }
        $charLenTxt = if ($c.NonNullCount -gt 0) { "{0} / {1:0.0} / {2}" -f $c.MinCharLen, $c.AvgCharLen, $c.MaxCharLen } else { '-' }
        $s3Rows += , @(
            $c.TableName, $c.ColumnName, [double]($c.ColumnIndex + 1), $c.InferredType,
            $(if ($c.KeyRole) { $c.KeyRole } else { '一般' }),
            [double]$c.TotalRows, [double]$c.NonNullCount, [double]$c.NullCount, [double]$c.NullRate,
            [double]$c.DistinctCount, [double]$c.CardinalityRate, $c.SelectivityRating,
            [double]$c.MinByteLen, [Math]::Round($c.AvgByteLen, 1), [double]$c.P95ByteLen, [double]$c.MaxByteLen, $charLenTxt,
            $topVal, [double]$c.Top1Count, [double]$c.Top1Share, [double]$c.Top3Share, [double]$c.Top5Share,
            [Math]::Round($c.GiniCoefficient, 3), [Math]::Round($c.NormalizedEntropy, 3), $c.SkewRisk,
            $(if ($c.MinValueText) { $c.MinValueText } else { '-' }),
            $(if ($c.MaxValueText) { $c.MaxValueText } else { '-' }),
            $(if ($c.TopValuesSummary) { $c.TopValuesSummary } else { '-' }),
            $(if ($c.DetailStatsSummary) { $c.DetailStatsSummary } else { '-' }),
            $(if ($c.RelationSummary) { $c.RelationSummary } else { '-' }),
            $(if ($c.AlertNote) { $c.AlertNote } else { '正常' })
        )
    }

    $topStartRow1 = 12
    $topEndRow1 = 12
    $summaryLastRow1 = 20
    $s1Xml = Build-SummarySheetXml `
        -Title 'Aurora Serverless v2 (MySQL 8.4.7) 性能試験用データ分析レポート — サマリー' `
        -Subtitle $subInfo -AnalysisObj $Analysis -RoleSummaryItems $roleSummary -TypeSummaryItems $typeSummary `
        -SizeKpiLbl $sizeKpiLabel -SizeColHdr $sizeColHeader -IsDetailed ([bool]$DetailedAnalysis) `
        -OutTopStartRow ([ref]$topStartRow1) -OutTopEndRow ([ref]$topEndRow1) -OutLastRow ([ref]$summaryLastRow1)

    $s2Headers = @('No','テーブル名',$srcColHeader,'文字コード','推定役割','レコード数','件数構成比','列数','データ型構成',$sizeColHeader,'容量構成比','最小行長(B)','平均行長(B)','最大行長(B)','推定InnoDB(KB)','推定InnoDBページ数','推定行/ページ','平均NULL率','平均ユニーク率','品質リスク列内訳','列数不一致行','PK/UK候補','カテゴリ区分カラム','日時カラム','日時データ期間','リレーション概要','JMeter並列適性','警告数','警告・注意点')
    $s2Widths  = @(5,22,24,15,16,12,10,7,24,15,10,11,11,11,14,13,12,10,11,22,11,18,22,20,24,28,22,8,32)
    $s2Xml = Build-SheetXml -Title $tblSheetTitle -Subtitle $subInfo -Headers $s2Headers -Rows $s2Rows -ColWidths $s2Widths `
        -NumColStyles @{1=6;6=6;7=10;8=6;10=8;11=10;12=6;13=12;14=6;15=12;16=6;17=12;18=10;19=10;21=6;28=6} `
        -CenterCols @(1,4,5,8,25,27) -WrapCols @(9,20,22,23,24,26,29) -WarnCols @(20,27,29) -WarnPositiveIntCols @(21,28) -NavyOptimalCols @(27) `
        -SummaryFooterRow $s2Footer -SummaryFooterStyles @{6=30;7=32;8=30;10=31;11=32;12=30;13=33;14=30;15=33;16=30;17=33;18=32;19=32;21=30;28=30} -SummaryFooterWarnCols @(21,28)

    $s3Headers = @('テーブル名','カラム名','列順','推定データ型','キー・役割','総行数','有効件数','NULL件数','NULL率','ユニーク数','ユニーク率','INDEX選択性評価','最小長(B)','平均長(B)','P95長(B)','最大長(B)','文字長(最小/平均/最大)','最頻値(Top1)','Top1件数','Top1占有率','Top3占有率','Top5占有率','ジニ係数','正規化エントロピー','偏りリスク','最小値 / 最古日時','最大値 / 最新日時','上位値分布 (Top5)','型別詳細統計 (数値・日時・文字特性)','リレーション推定','警告・注意点')
    $s3Widths  = @(22,20,6,16,14,10,10,10,9,11,10,16,9,10,10,10,17,22,10,11,11,11,9,11,20,20,20,36,44,30,30)
    $s3Xml = Build-SheetXml -Title '全カラム詳細プロファイル (選択性・NULL率・偏り・バイト長・型別統計・上位値分布・関連)' -Subtitle $subInfo -Headers $s3Headers -Rows $s3Rows -ColWidths $s3Widths `
        -NumColStyles @{3=6;6=6;7=6;8=6;9=10;10=6;11=10;13=6;14=12;15=6;16=6;19=6;20=10;21=10;22=10;23=14;24=14} `
        -CenterCols @(3,4,5,12,17,25,26,27) -WrapCols @(28,29,30,31) -WarnCols @(25,31)

    $sheetNames = New-Object System.Collections.Generic.List[string]
    $sheetXmls  = New-Object System.Collections.Generic.List[string]
    $sheetNames.Add('01_サマリー');     $sheetXmls.Add($s1Xml)
    $sheetNames.Add('02_テーブル一覧'); $sheetXmls.Add($s2Xml)
    $sheetNames.Add('03_カラム分析');   $sheetXmls.Add($s3Xml)

    if ($DetailedAnalysis) {
        $s4Rows = @()
        foreach ($cc in ($allCols | Where-Object { $_.IsCategorical } | Sort-Object NonNullCount -Descending)) {
            $r = 0
            foreach ($tv in $cc.TopValues) {
                $r++
                $weightHint = if ($tv.Share -ge 0.50) { '主要パス (高頻度)' } elseif ($tv.Share -ge 0.15) { '標準パス' } elseif ($tv.Share -ge 0.03) { '副次パス' } else { '低頻度・エッジケース' }
                $s4Rows += , @($cc.TableName, $cc.ColumnName, [double]$cc.DistinctCount, [double]$r, $tv.Value, [double]$tv.Count, [double]$tv.Share, [double]$tv.CumulativeShare, $weightHint)
            }
        }

        $s5Rows = @()
        foreach ($nc in ($allCols | Where-Object { $_.IsNumeric -and $_.NonNullCount -gt 0 } | Sort-Object NonNullCount -Descending)) {
            $s5Rows += , @($nc.TableName, $nc.ColumnName, [double]$nc.NonNullCount, [Math]::Round($nc.NumMin, 2), [Math]::Round($nc.NumP25, 2), [Math]::Round($nc.NumMedian, 2), [Math]::Round($nc.NumMean, 2), [Math]::Round($nc.NumP75, 2), [Math]::Round($nc.NumP90, 2), [Math]::Round($nc.NumP95, 2), [Math]::Round($nc.NumP99, 2), [Math]::Round($nc.NumMax, 2), [Math]::Round($nc.NumStdDev, 2), [double]$nc.ZeroRate, [double]$nc.NegativeRate, [double]$nc.OutlierRate)
        }

        $s6Rows = @()
        foreach ($dc in ($allCols | Where-Object { $_.IsDateTime -and $_.NonNullCount -gt 0 } | Sort-Object NonNullCount -Descending)) {
            $s6Rows += , @($dc.TableName, $dc.ColumnName, [double]$dc.NonNullCount, $(if ($dc.MinDate) { $dc.MinDate.ToString('yyyy/MM/dd HH:mm:ss') } else { '-' }), $(if ($dc.MaxDate) { $dc.MaxDate.ToString('yyyy/MM/dd HH:mm:ss') } else { '-' }), [Math]::Round($dc.DateSpanDays, 1), [Math]::Round($dc.AvgPerDay, 1), $dc.PeakDay, [double]$dc.PeakDayCount, [Math]::Round($dc.PeakBurstRatio, 2))
        }

        $s7Rows = @()
        foreach ($rl in $Analysis.Relationships) {
            $s7Rows += , @("$($rl.ParentTable).$($rl.ParentColumn)", "$($rl.ChildTable).$($rl.ChildColumn)", [double]$rl.ParentDistinctCount, [double]$rl.ChildNonNullRows, [double]$rl.MatchRate, [double]$rl.OrphanRows, [double]$rl.FanOutMin, [Math]::Round($rl.FanOutAvg, 2), [Math]::Round($rl.FanOutP50, 1), [Math]::Round($rl.FanOutP95, 1), [Math]::Round($rl.FanOutP99, 1), [double]$rl.FanOutMax, $rl.RiskLevel, $rl.Note)
        }

        $s8Rows = @()
        foreach ($w in $Analysis.Warnings) {
            $s8Rows += , @($w.Severity, $w.Category, $w.TableName, $w.ColumnName, $w.MetricValue, $w.PerfImpact, $w.Recommendation)
        }

        $sheetNames.Add('04_分布・割合分析')
        $sheetXmls.Add((Build-SheetXml -Title 'カテゴリ・区分値 分布と割合' -Subtitle $subInfo -Headers @('テーブル名','カラム名','区分数','順位','カテゴリ値','件数','割合(%)','累積割合(%)','性能試験シナリオ区分') -Rows $s4Rows -ColWidths @(20,20,8,6,26,12,11,11,20) -NumColStyles @{3=6;4=6;6=6;7=10;8=10} -CenterCols @(9) -WrapCols @(5)))
        $sheetNames.Add('05_数値統計・ヒストグラム')
        $sheetXmls.Add((Build-SheetXml -Title '数値カラム統計量' -Subtitle $subInfo -Headers @('テーブル名','カラム名','有効件数','最小値','P25(Q1)','中央値(P50)','平均値','P75(Q3)','P90','P95','P99','最大値','標準偏差','ゼロ率','負値率','外れ値率') -Rows $s5Rows -ColWidths @(20,22,11,12,12,12,12,12,12,12,12,12,12,9,9,9) -NumColStyles @{3=6;4=8;5=8;6=8;7=8;8=8;9=8;10=8;11=8;12=8;13=8;14=10;15=10;16=10}))
        $sheetNames.Add('06_時系列分析')
        $sheetXmls.Add((Build-SheetXml -Title '時系列・日時分布分析' -Subtitle $subInfo -Headers @('テーブル名','日時カラム','有効件数','最古日時','最新日時','期間(日)','日平均件数','ピーク日','ピーク日件数','ピーク倍率') -Rows $s6Rows -ColWidths @(20,20,12,22,22,11,12,14,13,13) -NumColStyles @{3=6;6=12;7=12;9=6;10=8} -CenterCols @(4,5,8)))
        $sheetNames.Add('07_リレーション・ファンアウト')
        $sheetXmls.Add((Build-SheetXml -Title 'テーブル間リレーション・1:Nファンアウト' -Subtitle $subInfo -Headers @('親テーブル.PK','子テーブル.FK','親キー数','子有効行数','参照一致率','孤立行数','最小(子/親)','平均(子/親)','P50(子/親)','P95(子/親)','P99(子/親)','最大(子/親)','リスク評価','所見・スーパーユーザー情報') -Rows $s7Rows -ColWidths @(24,24,11,11,11,10,10,11,10,10,10,11,24,38) -NumColStyles @{3=6;4=6;5=10;6=6;7=6;8=8;9=12;10=12;11=12;12=6} -CenterCols @(13) -WrapCols @(14) -WarnCols @(13,14) -WarnPositiveIntCols @(6)))
        $sheetNames.Add('08_性能試験ガイド・警告')
        $sheetXmls.Add((Build-SheetXml -Title '性能試験設計ガイド・検出警告一覧' -Subtitle $subInfo -Headers @('重要度','カテゴリ','テーブル名','カラム名','検出値・指標','Aurora / 性能試験への影響','推奨アクション') -Rows $s8Rows -ColWidths @(10,18,22,18,28,42,46) -CenterCols @(1,2) -WrapCols @(5,6,7) -WarnCols @(1) -HighlightCatOnWarnCol1 $true))
    }

    $hasChart1 = ($Analysis.Tables.Count -gt 0)

    # styles.xml: strictly Meiryo UI + #333333 + #1A2B4C + #800020 on pure neutral monotone backgrounds
    $stylesXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="5">
    <numFmt numFmtId="164" formatCode="#,##0"/>
    <numFmt numFmtId="165" formatCode="#,##0.00"/>
    <numFmt numFmtId="166" formatCode="0.0%"/>
    <numFmt numFmtId="167" formatCode="#,##0.0"/>
    <numFmt numFmtId="168" formatCode="0.000"/>
  </numFmts>
  <fonts count="12">
    <font><sz val="9.5"/><color rgb="FF333333"/><name val="Meiryo UI"/></font>
    <font><b/><sz val="13.5"/><color rgb="FFFFFFFF"/><name val="Meiryo UI"/></font>
    <font><sz val="9"/><color rgb="FF333333"/><name val="Meiryo UI"/></font>
    <font><b/><sz val="9.5"/><color rgb="FFFFFFFF"/><name val="Meiryo UI"/></font>
    <font><b/><sz val="9.5"/><color rgb="FF800020"/><name val="Meiryo UI"/></font>
    <font><sz val="9.5"/><color rgb="FF595959"/><name val="Meiryo UI"/></font>
    <font><b/><sz val="10.5"/><color rgb="FF1A2B4C"/><name val="Meiryo UI"/></font>
    <font><b/><sz val="8.5"/><color rgb="FF595959"/><name val="Meiryo UI"/></font>
    <font><b/><sz val="12.5"/><color rgb="FF1A2B4C"/><name val="Meiryo UI"/></font>
    <font><b/><sz val="8.5"/><color rgb="FF800020"/><name val="Meiryo UI"/></font>
    <font><b/><sz val="12.5"/><color rgb="FF800020"/><name val="Meiryo UI"/></font>
    <font><b/><sz val="9.5"/><color rgb="FF1A2B4C"/><name val="Meiryo UI"/></font>
  </fonts>
  <fills count="8">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1A2B4C"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF7F7F7"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF5F5F5"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFEAEAEA"/></patternFill></fill>
  </fills>
  <borders count="10">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFD9D9D9"/></left>
      <right style="thin"><color rgb="FFD9D9D9"/></right>
      <top style="thin"><color rgb="FFD9D9D9"/></top>
      <bottom style="thin"><color rgb="FFD9D9D9"/></bottom>
      <diagonal/>
    </border>
    <border>
      <left/><right/><top/>
      <bottom style="thin"><color rgb="FF1A2B4C"/></bottom>
      <diagonal/>
    </border>
    <border>
      <left style="thin"><color rgb="FF4D5B75"/></left>
      <right style="thin"><color rgb="FF4D5B75"/></right>
      <top style="thin"><color rgb="FF1A2B4C"/></top>
      <bottom style="thin"><color rgb="FF1A2B4C"/></bottom>
      <diagonal/>
    </border>
    <border>
      <left style="medium"><color rgb="FF1A2B4C"/></left>
      <right/><top/>
      <bottom style="thin"><color rgb="FF1A2B4C"/></bottom>
      <diagonal/>
    </border>
    <border>
      <left style="thin"><color rgb="FFD9D9D9"/></left>
      <right style="thin"><color rgb="FFD9D9D9"/></right>
      <top style="medium"><color rgb="FF1A2B4C"/></top>
      <bottom/><diagonal/>
    </border>
    <border>
      <left style="thin"><color rgb="FFD9D9D9"/></left>
      <right style="thin"><color rgb="FFD9D9D9"/></right>
      <top style="medium"><color rgb="FF800020"/></top>
      <bottom/><diagonal/>
    </border>
    <border>
      <left style="thin"><color rgb="FFD9D9D9"/></left>
      <right style="thin"><color rgb="FFD9D9D9"/></right>
      <top/>
      <bottom style="thin"><color rgb="FFD9D9D9"/></bottom>
      <diagonal/>
    </border>
    <border>
      <left style="thin"><color rgb="FFD9D9D9"/></left>
      <right style="thin"><color rgb="FFD9D9D9"/></right>
      <top style="thin"><color rgb="FF1A2B4C"/></top>
      <bottom style="thin"><color rgb="FF1A2B4C"/></bottom>
      <diagonal/>
    </border>
    <border>
      <left style="medium"><color rgb="FF1A2B4C"/></left>
      <right style="thin"><color rgb="FFD9D9D9"/></right>
      <top style="thin"><color rgb="FFD9D9D9"/></top>
      <bottom style="thin"><color rgb="FFD9D9D9"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="39">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center" indent="1"/></xf>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" indent="1"/></xf>
    <xf numFmtId="0" fontId="3" fillId="2" borderId="3" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="164" fontId="0" fillId="6" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="164" fontId="0" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="165" fontId="0" fillId="6" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="165" fontId="0" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="166" fontId="0" fillId="6" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="166" fontId="0" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="167" fontId="0" fillId="6" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="167" fontId="0" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="168" fontId="0" fillId="6" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="168" fontId="0" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="4" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="4" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="164" fontId="4" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="0" fontId="5" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="6" fillId="3" borderId="4" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" indent="1"/></xf>
    <xf numFmtId="0" fontId="7" fillId="6" borderId="5" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="8" fillId="6" borderId="7" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="9" fillId="6" borderId="6" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="10" fillId="6" borderId="7" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="11" fillId="7" borderId="8" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="164" fontId="11" fillId="7" borderId="8" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="165" fontId="11" fillId="7" borderId="8" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="166" fontId="11" fillId="7" borderId="8" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="167" fontId="11" fillId="7" borderId="8" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="0" fontId="11" fillId="7" borderId="8" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="164" fontId="4" fillId="7" borderId="8" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="4" borderId="9" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1" indent="1"/></xf>
    <xf numFmtId="0" fontId="11" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="11" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
'@

    if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }
    $zip = [System.IO.Compression.ZipFile]::Open($stagingPath, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        function Add-ZipText([string]$entryName, [string]$text) {
            if ($null -ne $PumpCallback) { & $PumpCallback }
            if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }
            $e = $zip.CreateEntry($entryName, [System.IO.Compression.CompressionLevel]::Optimal)
            $sw = New-Object System.IO.StreamWriter($e.Open(), (New-Object System.Text.UTF8Encoding($false)))
            try { $sw.Write($text) } finally { $sw.Dispose() }
        }

        $snapshotXml = [AuroraAnalyzer.AnalysisSnapshotCodec]::SerializeToXml($Analysis, $DbLabel)
        $hasSnapshotXml = -not [string]::IsNullOrEmpty($snapshotXml)
        $ctCustomXml = if ($hasSnapshotXml) { '<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>' } else { '' }
        $ctExtra = if ($hasChart1) { '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>' } else { '' }
        $sheetCt = ($sheetNames | ForEach-Object -Begin { $i = 0 } -Process { $i++; "<Override PartName=`"/xl/worksheets/sheet$i.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>" }) -join ''
        Add-ZipText '[Content_Types].xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' + $sheetCt + $ctExtra + $ctCustomXml + '</Types>')
        Add-ZipText '_rels/.rels' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'

        $wbSheets = ''
        $wbRels = '<Relationship Id="rIdStyle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        if ($hasSnapshotXml) {
            $wbRels += '<Relationship Id="rIdCustomXml1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>'
        }
        for ($i = 0; $i -lt $sheetNames.Count; $i++) {
            $idx = $i + 1
            $wbSheets += "<sheet name=`"$(XmlEsc $sheetNames[$i])`" sheetId=`"$idx`" r:id=`"rId$idx`"/>"
            $wbRels += "<Relationship Id=`"rId$idx`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$idx.xml`"/>"
            Add-ZipText "xl/worksheets/sheet$idx.xml" $sheetXmls[$i]
        }

        Add-ZipText 'xl/workbook.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + $wbSheets + '</sheets></workbook>')
        Add-ZipText 'xl/_rels/workbook.xml.rels' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + $wbRels + '</Relationships>')
        Add-ZipText 'xl/styles.xml' $stylesXml
        if ($hasSnapshotXml) {
            $cxmlGuid = ('{' + [Guid]::NewGuid().ToString().ToUpperInvariant() + '}')
            Add-ZipText 'customXml/item1.xml' $snapshotXml
            Add-ZipText 'customXml/itemProps1.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="no"?><ds:datastoreItem ds:itemID="' + $cxmlGuid + '" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs><ds:schemaRef ds:uri="' + [AuroraAnalyzer.AnalysisSnapshotCodec]::SnapshotNamespace + '"/></ds:schemaRefs></ds:datastoreItem>')
            Add-ZipText 'customXml/_rels/item1.xml.rels' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/></Relationships>'
        }

        if ($hasChart1) {
            $lastRow1 = $summaryLastRow1
            Add-ZipText 'xl/worksheets/_rels/sheet1.xml.rels' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>'
            Add-ZipText 'xl/drawings/_rels/drawing1.xml.rels' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>'
            Add-ZipText 'xl/drawings/drawing1.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + ($lastRow1 + 2) + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + ($lastRow1 + 17) + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>')
            Add-ZipText 'xl/charts/chart1.xml' ("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><c:chartSpace xmlns:c=`"http://schemas.openxmlformats.org/drawingml/2006/chart`" xmlns:a=`"http://schemas.openxmlformats.org/drawingml/2006/main`"><c:spPr><a:solidFill><a:srgbClr val=`"FFFFFF`"/></a:solidFill><a:ln w=`"9525`"><a:solidFill><a:srgbClr val=`"D9D9D9`"/></a:solidFill></a:ln></c:spPr><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr b=`"1`" sz=`"1050`"><a:solidFill><a:srgbClr val=`"1A2B4C`"/></a:solidFill><a:latin typeface=`"Meiryo UI`"/><a:ea typeface=`"Meiryo UI`"/></a:defRPr></a:pPr><a:r><a:rPr lang=`"ja-JP`" b=`"1`" sz=`"1050`"><a:solidFill><a:srgbClr val=`"1A2B4C`"/></a:solidFill><a:latin typeface=`"Meiryo UI`"/><a:ea typeface=`"Meiryo UI`"/></a:rPr><a:t>テーブル別レコード数 TOP15</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=`"0`"/></c:title><c:plotArea><c:spPr><a:solidFill><a:srgbClr val=`"FFFFFF`"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr><c:barChart><c:barDir val=`"bar`"/><c:grouping val=`"clustered`"/><c:gapWidth val=`"70`"/><c:ser><c:idx val=`"0`"/><c:order val=`"0`"/><c:spPr><a:solidFill><a:srgbClr val=`"1A2B4C`"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr><c:dLbls><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz=`"800`"><a:solidFill><a:srgbClr val=`"333333`"/></a:solidFill><a:latin typeface=`"Meiryo UI`"/><a:ea typeface=`"Meiryo UI`"/></a:defRPr></a:pPr><a:endParaRPr lang=`"ja-JP`"/></a:p></c:txPr><c:dLblPos val=`"outEnd`"/><c:showLegendKey val=`"0`"/><c:showVal val=`"1`"/><c:showCatName val=`"0`"/><c:showSerName val=`"0`"/><c:showPercent val=`"0`"/><c:showBubbleSize val=`"0`"/></c:dLbls><c:cat><c:strRef><c:f>'01_サマリー'!`$B`$${topStartRow1}:`$B`$${topEndRow1}</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>'01_サマリー'!`$D`$${topStartRow1}:`$D`$${topEndRow1}</c:f></c:numRef></c:val></c:ser><c:axId val=`"1`"/><c:axId val=`"2`"/></c:barChart><c:catAx><c:axId val=`"1`"/><c:scaling><c:orientation val=`"maxMin`"/></c:scaling><c:delete val=`"0`"/><c:axPos val=`"l`"/><c:tickLblPos val=`"low`"/><c:spPr><a:ln w=`"9525`"><a:solidFill><a:srgbClr val=`"D9D9D9`"/></a:solidFill></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz=`"850`"><a:solidFill><a:srgbClr val=`"333333`"/></a:solidFill><a:latin typeface=`"Meiryo UI`"/><a:ea typeface=`"Meiryo UI`"/></a:defRPr></a:pPr><a:endParaRPr lang=`"ja-JP`"/></a:p></c:txPr><c:crossAx val=`"2`"/></c:catAx><c:valAx><c:axId val=`"2`"/><c:scaling><c:orientation val=`"minMax`"/></c:scaling><c:delete val=`"0`"/><c:axPos val=`"b`"/><c:majorGridlines><c:spPr><a:ln w=`"9525`"><a:solidFill><a:srgbClr val=`"E6E6E6`"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:spPr><a:ln><a:noFill/></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz=`"850`"><a:solidFill><a:srgbClr val=`"333333`"/></a:solidFill><a:latin typeface=`"Meiryo UI`"/><a:ea typeface=`"Meiryo UI`"/></a:defRPr></a:pPr><a:endParaRPr lang=`"ja-JP`"/></a:p></c:txPr><c:crossAx val=`"1`"/></c:valAx></c:plotArea></c:chart></c:chartSpace>")
        }
        $zip.Dispose()
        $zip = $null
        if ($null -ne $LogCallback) { & $LogCallback "Excelファイル保存中: $OutputPath" 98 }
        if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }
        [AuroraAnalyzer.AtomicFile]::Publish($stagingPath, $OutputPath)
    }
    finally {
        try { if ($null -ne $zip) { $zip.Dispose() } }
        finally { try { if ([IO.File]::Exists($stagingPath)) { [IO.File]::Delete($stagingPath) } } catch { } }
    }
}

function New-AnalysisStagingPath {
    param([string]$OutputPath)
    $dir = [IO.Path]::GetDirectoryName($OutputPath)
    [void][IO.Directory]::CreateDirectory($dir)
    return [IO.Path]::Combine($dir, ('.aurora-' + [guid]::NewGuid().ToString('N') + '.xlsx'))
}

function Export-AnalysisReport {
    param(
        [Parameter(Mandatory = $true)]$Analysis,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [string]$DbLabel = '',
        [string]$Engine = 'Auto',
        [switch]$DetailedAnalysis,
        [scriptblock]$LogCallback = $null,
        [AuroraAnalyzer.OperationCancelState]$CancelState = $null,
        [scriptblock]$PumpCallback = $null
    )

    if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }

    if ($Engine -eq 'OpenXml') {
        Export-AnalysisToOpenXml -Analysis $Analysis -OutputPath $OutputPath -DbLabel $DbLabel -DetailedAnalysis:$DetailedAnalysis -LogCallback $LogCallback -CancelState $CancelState -PumpCallback $PumpCallback
        return 'OpenXml'
    }

    $comAvailable = $false
    if ($Engine -eq 'Com' -or $Engine -eq 'Auto') {
        try {
            $progId = [Type]::GetTypeFromProgID('Excel.Application')
            if ($null -ne $progId) { $comAvailable = $true }
        }
        catch { $comAvailable = $false }
    }

    if ($comAvailable) {
        try {
            Export-AnalysisToExcelCom -Analysis $Analysis -OutputPath $OutputPath -DbLabel $DbLabel -DetailedAnalysis:$DetailedAnalysis -LogCallback $LogCallback -CancelState $CancelState -PumpCallback $PumpCallback
            return 'Com'
        }
        catch [AuroraAnalyzer.ExcelActivationException] {
            if ($Engine -eq 'Com') { throw }
        }
    }
    elseif ($Engine -eq 'Com') {
        throw "Microsoft Excel COM (Excel.Application) が利用できません。"
    }
    if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }
    if ($null -ne $LogCallback) { & $LogCallback 'Excel COM を利用できないため OpenXML でレポートを生成します。' 60 }
    Export-AnalysisToOpenXml -Analysis $Analysis -OutputPath $OutputPath -DbLabel $DbLabel -DetailedAnalysis:$DetailedAnalysis -LogCallback $LogCallback -CancelState $CancelState -PumpCallback $PumpCallback
    return 'OpenXml'
}

# ==============================================================================
#  Module 4C: Comparison Excel COM Report Generator (Monotone + Navy + Burgundy)
# ==============================================================================
function Set-ComparisonRowStyles {
    param(
        $Sheet,
        [int]$StartRow,
        [object[]]$Items,
        [int]$StatusCol,
        [int]$AlertCol = 0,
        [int[]]$SignedDeltaCols = @(),
        [int[]]$WrapTextCols = @(),
        [object[,]]$DataMatrix = $null
    )

    $navyOle   = ConvertTo-ExcelOleColor $script:THEME.AccentHex
    $subBgOle  = ConvertTo-ExcelOleColor $script:THEME.AccentLightHex
    $textOle   = ConvertTo-ExcelOleColor $script:THEME.TextHex
    $mutedOle  = ConvertTo-ExcelOleColor $script:THEME.MutedHex
    $warnOle   = ConvertTo-ExcelOleColor $script:THEME.WarnHex
    $warnBgOle = ConvertTo-ExcelOleColor $script:THEME.WarnBgHex

    $wrapWidthMap = @{}
    foreach ($wc in $WrapTextCols) {
        try { $wrapWidthMap[$wc] = [double]$Sheet.Columns.Item($wc).ColumnWidth } catch { $wrapWidthMap[$wc] = 26.0 }
    }

    for ($i = 0; $i -lt $Items.Count; $i++) {
        $excelRow = $StartRow + $i
        $it = $Items[$i]

        # Dynamic row height based on visual width of wrapped text columns
        if ($null -ne $DataMatrix -and $WrapTextCols.Count -gt 0) {
            $wTexts = New-Object System.Collections.Generic.List[string]
            $wWidths = New-Object System.Collections.Generic.List[double]
            foreach ($wc in $WrapTextCols) {
                $colIdx = $wc - 1
                $wVal = [string]$DataMatrix[$i, $colIdx]
                $wTexts.Add($wVal)
                $wWidths.Add([double]$wrapWidthMap[$wc])
            }
            $Sheet.Rows.Item($excelRow).RowHeight = Get-EstimatedRowHeight -Texts $wTexts.ToArray() -ColWidths $wWidths.ToArray() -MinHeight 22.0 -MaxHeight 110.0
        }
        else {
            $Sheet.Rows.Item($excelRow).RowHeight = 22
        }

        # Status badge styling (strictly neutral monotone backgrounds; zero same-color-family tint)
        if ($StatusCol -gt 0) {
            $sc = $Sheet.Cells.Item($excelRow, $StatusCol)
            $st = [string]$it.PrimaryStatus
            if ($st -match '(削除|減少)') {
                $sc.Font.Color = $warnOle
                $sc.Font.Bold = $true
                $sc.Interior.Color = $warnBgOle
            }
            elseif ($st -match '^追加') {
                $sc.Font.Color = $navyOle
                $sc.Font.Bold = $true
                $sc.Interior.Color = $subBgOle
            }
            elseif ($st -match '^更新') {
                $sc.Font.Color = $textOle
                $sc.Font.Bold = $true
                $sc.Interior.Color = $subBgOle
            }
            else {
                $sc.Font.Color = $mutedOle
            }
        }

        # Alert column styling (Burgundy bold on neutral monotone background for actual alerts)
        if ($AlertCol -gt 0 -and $it.HasAlert) {
            $ac = $Sheet.Cells.Item($excelRow, $AlertCol)
            $ac.Font.Color = $warnOle
            $ac.Font.Bold = $true
            $ac.Interior.Color = $warnBgOle
        }

        # Emphasize signed delta cells (Navy bold for positive, Burgundy bold for negative)
        foreach ($dc in $SignedDeltaCols) {
            if ($null -eq $DataMatrix) { continue }
            $rawVal = $DataMatrix[$i, ($dc - 1)]
            if ($rawVal -is [double] -or $rawVal -is [int] -or $rawVal -is [long]) {
                $dv = [double]$rawVal
                $cell = $Sheet.Cells.Item($excelRow, $dc)
                if ($dv -lt -1e-9) {
                    $cell.Font.Color = $warnOle
                    $cell.Font.Bold = $true
                }
                elseif ($dv -gt 1e-9) {
                    $cell.Font.Color = $navyOle
                    $cell.Font.Bold = $true
                }
                else {
                    $cell.Font.Color = $mutedOle
                }
            }
        }
    }
}

function Write-ComparisonTotalRow {
    param(
        $Sheet,
        [int]$Row,
        [object[]]$Values,
        [hashtable]$ColumnFormats = @{},
        [int[]]$CenterCols = @(),
        [int[]]$RightCols = @()
    )

    $colCount = $Values.Count
    $navyOle   = ConvertTo-ExcelOleColor $script:THEME.AccentHex
    $totalBg   = ConvertTo-ExcelOleColor $script:THEME.TotalBgHex
    $darkBdr   = ConvertTo-ExcelOleColor $script:THEME.BorderDarkHex
    $bdrOle    = ConvertTo-ExcelOleColor $script:THEME.BorderHex
    $warnOle   = ConvertTo-ExcelOleColor $script:THEME.WarnHex

    for ($c = 1; $c -le $colCount; $c++) {
        if ($ColumnFormats.ContainsKey($c)) {
            $Sheet.Cells.Item($Row, $c).NumberFormat = $ColumnFormats[$c]
        }
    }

    $mat = New-Object 'object[,]' 1, $colCount
    for ($c = 0; $c -lt $colCount; $c++) { $mat[0, $c] = $Values[$c] }
    $rng = $Sheet.Range($Sheet.Cells.Item($Row, 1), $Sheet.Cells.Item($Row, $colCount))
    [AuroraAnalyzer.ComHelper]::SetValue2($rng, $mat)
    $rng.Font.Name = $script:THEME.FontName
    $rng.Font.Size = 9.5
    $rng.Font.Bold = $true
    $rng.Font.Color = $navyOle
    $rng.Interior.Color = $totalBg
    $rng.VerticalAlignment = -4108
    $rng.HorizontalAlignment = -4131
    $Sheet.Rows.Item($Row).RowHeight = 24

    foreach ($e in @(7, 10, 11)) {
        try { $rng.Borders.Item($e).LineStyle = 1; $rng.Borders.Item($e).Weight = 2; $rng.Borders.Item($e).Color = $bdrOle } catch { }
    }
    try { $rng.Borders.Item(8).LineStyle = 1; $rng.Borders.Item(8).Weight = 2; $rng.Borders.Item(8).Color = $darkBdr } catch { }
    try { $rng.Borders.Item(9).LineStyle = 1; $rng.Borders.Item(9).Weight = 3; $rng.Borders.Item(9).Color = $navyOle } catch { }

    for ($c = 1; $c -le $colCount; $c++) {
        $cell = $Sheet.Cells.Item($Row, $c)
        if ($CenterCols -contains $c) { $cell.HorizontalAlignment = -4108 }
        elseif ($RightCols -contains $c) { $cell.HorizontalAlignment = -4152 }
        if (($Values[$c - 1] -is [double] -or $Values[$c - 1] -is [long]) -and [double]$Values[$c - 1] -lt -1e-9) {
            $cell.Font.Color = $warnOle
        }
    }
}

function Export-ComparisonToExcelCom {
    param(
        [Parameter(Mandatory = $true)]$Comparison,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [scriptblock]$LogCallback = $null,
        [AuroraAnalyzer.OperationCancelState]$CancelState = $null,
        [scriptblock]$PumpCallback = $null
    )

    function Write-CmpLog {
        param([string]$Msg, [int]$Pct = -1)
        if ($null -ne $LogCallback) { & $LogCallback $Msg $Pct }
        if ($Pct -lt 100 -and $null -ne $CancelState) { $CancelState.ThrowIfCancelled() }
    }

    if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }

    $OutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
    $stagingPath = New-AnalysisStagingPath $OutputPath
    $checkReport = {
        if ($null -ne $PumpCallback) { & $PumpCallback }
        if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }
    }

    $excel = $null
    $excelPid = 0
    $wb = $null
    $sheetsToRelease = New-Object System.Collections.Generic.List[object]

    try {
        Write-CmpLog "テスト前後比較Excelレポート生成開始 (Excel COM / Meiryo UI モノトーン仕様)..." 45
        try { $excel = New-Object -ComObject Excel.Application }
        catch [System.Runtime.InteropServices.COMException] {
            throw [AuroraAnalyzer.ExcelActivationException]::new($_.Exception)
        }
        try { $excelPid = [AuroraAnalyzer.ComHelper]::GetWindowProcessId([int]$excel.Hwnd) } catch { $excelPid = 0 }
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false
        try { $excel.EnableEvents = $false } catch { }
        $targetCmpSheets = 8
        $excel.SheetsInNewWorkbook = $targetCmpSheets
        $wb = $excel.Workbooks.Add()
        while ($wb.Worksheets.Count -lt $targetCmpSheets) {
            [void]$wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
        }
        while ($wb.Worksheets.Count -gt $targetCmpSheets) {
            $wb.Worksheets.Item($wb.Worksheets.Count).Delete()
        }
        $excel.Calculation = -4135 # xlCalculationManual

        $wsSummary   = $wb.Worksheets.Item(1); $wsSummary.Name   = '01_比較サマリー'
        $wsAll       = $wb.Worksheets.Item(2); $wsAll.Name       = '02_全テーブル比較一覧'
        $wsUpdated   = $wb.Worksheets.Item(3); $wsUpdated.Name   = '03_更新のあったテーブル'
        $wsAdded     = $wb.Worksheets.Item(4); $wsAdded.Name     = '04_追加のあったテーブル'
        $wsUnchanged = $wb.Worksheets.Item(5); $wsUnchanged.Name = '05_変更のなかったテーブル'
        $wsDiffs     = $wb.Worksheets.Item(6); $wsDiffs.Name     = '06_カラム・統計差分詳細'
        $wsAllCols   = $wb.Worksheets.Item(7); $wsAllCols.Name   = '07_全カラム比較一覧'
        $wsPerfDb    = $wb.Worksheets.Item(8); $wsPerfDb.Name    = '08_性能・DB・時系列・品質変化'

        foreach ($s in @($wsSummary, $wsAll, $wsUpdated, $wsAdded, $wsUnchanged, $wsDiffs, $wsAllCols, $wsPerfDb)) {
            $sheetsToRelease.Add($s)
            Initialize-ExcelSheetStyle -Sheet $s
        }

        $subInfo = "比較日時: {0}  |  開始前: {1}{2}  |  開始後: {3}{4}  |  照合方式: {5}" -f `
            $Comparison.ComparedAt.ToString('yyyy/MM/dd HH:mm:ss'), `
            $Comparison.BeforeFileName, `
            $(if ($Comparison.BeforeAnalyzedAt) { " ($($Comparison.BeforeAnalyzedAt))" } else { '' }), `
            $Comparison.AfterFileName, `
            $(if ($Comparison.AfterAnalyzedAt) { " ($($Comparison.AfterAnalyzedAt))" } else { '' }), `
            $Comparison.PrecisionLabel

        $navyOle   = ConvertTo-ExcelOleColor $script:THEME.AccentHex
        $textOle   = ConvertTo-ExcelOleColor $script:THEME.TextHex
        $mutedOle  = ConvertTo-ExcelOleColor $script:THEME.MutedHex
        $warnOle   = ConvertTo-ExcelOleColor $script:THEME.WarnHex
        $warnBgOle = ConvertTo-ExcelOleColor $script:THEME.WarnBgHex
        $cardBgOle = ConvertTo-ExcelOleColor $script:THEME.CardBgHex
        $borderOle = ConvertTo-ExcelOleColor $script:THEME.BorderHex
        $zebraOle  = ConvertTo-ExcelOleColor $script:THEME.ZebraHex
        $subBgOle  = ConvertTo-ExcelOleColor $script:THEME.AccentLightHex

        function Write-CmpKpiBand {
            param($Sheet, [int]$LabelRow, [object[]]$Kpis)
            $Sheet.Rows.Item($LabelRow).RowHeight = 20
            $Sheet.Rows.Item($LabelRow + 1).RowHeight = 30
            foreach ($kpi in $Kpis) {
                $lblRng = $Sheet.Range($Sheet.Cells.Item($LabelRow, $kpi.Col1), $Sheet.Cells.Item($LabelRow, $kpi.Col2))
                $lblRng.Merge()
                [AuroraAnalyzer.ComHelper]::SetValue2($Sheet.Cells.Item($LabelRow, $kpi.Col1), $kpi.Label)
                $lblRng.Font.Name = $script:THEME.FontName; $lblRng.Font.Size = 8.5; $lblRng.Font.Bold = $true
                $lblRng.Font.Color = if ($kpi.Warn) { $warnOle } else { $mutedOle }
                $lblRng.Interior.Color = $cardBgOle
                $lblRng.HorizontalAlignment = -4108; $lblRng.VerticalAlignment = -4108

                $valRng = $Sheet.Range($Sheet.Cells.Item($LabelRow + 1, $kpi.Col1), $Sheet.Cells.Item($LabelRow + 1, $kpi.Col2))
                $valRng.Merge()
                $valRng.NumberFormat = $kpi.Fmt
                Set-ExcelCellValue -Cell $Sheet.Cells.Item($LabelRow + 1, $kpi.Col1) -Value $kpi.Value
                $valRng.Font.Name = $script:THEME.FontName
                $valRng.Font.Size = if ($kpi.ContainsKey('FontSize')) { $kpi.FontSize } else { 13.0 }
                $valRng.Font.Bold = $true
                $valRng.Font.Color = if ($kpi.Warn) { $warnOle } else { $navyOle }
                $valRng.Interior.Color = $cardBgOle
                $valRng.HorizontalAlignment = -4108; $valRng.VerticalAlignment = -4108

                $box = $Sheet.Range($Sheet.Cells.Item($LabelRow, $kpi.Col1), $Sheet.Cells.Item($LabelRow + 1, $kpi.Col2))
                $box.Borders.Item(8).LineStyle = 1; $box.Borders.Item(8).Color = if ($kpi.Warn) { $warnOle } else { $navyOle }; $box.Borders.Item(8).Weight = 3
                foreach ($e in @(7, 9, 10)) { $box.Borders.Item($e).LineStyle = 1; $box.Borders.Item($e).Color = $borderOle; $box.Borders.Item($e).Weight = 2 }
            }
        }

        # ======================================================================
        #  SHEET 1: 01_比較サマリー
        # ======================================================================
        Write-CmpLog "シート作成中: 01_比較サマリー..." 48
        Set-ColumnWidths -Sheet $wsSummary -Widths @(21, 22, 18, 13, 13, 13, 12, 14, 15, 24, 28, 32)
        Set-SheetTitleBanner -Sheet $wsSummary -Title '性能試験 テスト開始前・開始後 分析ファイル比較レポート — 比較サマリー' -Subtitle $subInfo -EndCol 12

        $kpiRow1 = @(
            @{ Col1 = 1;  Col2 = 2;  Label = '比較対象テーブル (変更/不変)'; Value = ("全 {0} (変更 {1} / 不変 {2})" -f $Comparison.TotalUniqueTableCount, $Comparison.ChangedTableCount, $Comparison.UnchangedTableCount); Fmt = '@'; Warn = $false; FontSize = 11.5 },
            @{ Col1 = 3;  Col2 = 4;  Label = '更新あり / 追加ありテーブル';   Value = ("更新 {0} / 追加 {1}" -f $Comparison.UpdatedSheetCount, $Comparison.AddedSheetCount); Fmt = '@'; Warn = ($Comparison.RowDecreasedCount -gt 0 -or $Comparison.DeletedTableCount -gt 0); FontSize = 12.0 },
            @{ Col1 = 5;  Col2 = 6;  Label = '総レコード数増減';             Value = [double]$Comparison.TotalRowDelta; Fmt = '+#,##0" 件";-#,##0" 件";0" 件"'; Warn = ($Comparison.TotalRowDelta -lt 0); FontSize = 13.5 },
            @{ Col1 = 7;  Col2 = 8;  Label = '全カラム照合 (変更/不変)';     Value = ("変更 {0} / 不変 {1} (計{2})" -f $Comparison.ChangedColumnTotal, $Comparison.UnchangedColumnTotal, $Comparison.TotalUniqueColumnCount); Fmt = '@'; Warn = ($Comparison.RemovedColumnTotal -gt 0); FontSize = 11.0 },
            @{ Col1 = 9;  Col2 = 10; Label = '行レベル変動 (挿入/更新/削除)'; Value = $(if ($Comparison.HasAnyRowChurn) { ("+{0:#,##0} / 更新{1:#,##0} / -{2:#,##0}" -f $Comparison.TotalInsertedRows, $Comparison.TotalUpdatedRows, $Comparison.TotalDeletedRows) } else { '統計差分照合' }); Fmt = '@'; Warn = ($Comparison.TotalDeletedRows -gt 0); FontSize = 11.0 },
            @{ Col1 = 11; Col2 = 12; Label = '要注意・警告テーブル';         Value = [double]$Comparison.AlertTableCount; Fmt = '#,##0" テーブル"'; Warn = ($Comparison.AlertTableCount -gt 0); FontSize = 13.5 }
        )
        Write-CmpKpiBand -Sheet $wsSummary -LabelRow 4 -Kpis $kpiRow1
        $wsSummary.Rows.Item(6).RowHeight = 8

        $elapsedVal = if ($Comparison.ClockReversed) { $Comparison.ElapsedDisplay } elseif ($Comparison.HasElapsedTime) { ("{0} ({1}{2:0.0}件/分)" -f $Comparison.ElapsedDisplay, $(if ($Comparison.TotalRowsPerMinute -gt 0) { '+' } else { '' }), $Comparison.TotalRowsPerMinute) } else { '時刻同一 / 未計測' }
        $innoKpiVal = if ($Comparison.HasDbPhysicalMeta) {
            ("{0}{1:#,##0.0}KB (実Data{2}{3:#,##0.0}/Idx{4}{5:#,##0.0})" -f $(if ($Comparison.TotalInnoDbDeltaKB -gt 0) { '+' } else { '' }), $Comparison.TotalInnoDbDeltaKB, $(if ($Comparison.TotalDataLengthDeltaKB -gt 0) { '+' } else { '' }), $Comparison.TotalDataLengthDeltaKB, $(if ($Comparison.TotalIndexLengthDeltaKB -gt 0) { '+' } else { '' }), $Comparison.TotalIndexLengthDeltaKB)
        } else {
            ("{0}{1:#,##0.0} KB" -f $(if ($Comparison.TotalInnoDbDeltaKB -gt 0) { '+' } else { '' }), $Comparison.TotalInnoDbDeltaKB)
        }
        $warnDeltaVal = "新規 {0} / 悪化 {1} / 解消 {2}" -f $Comparison.NewWarningCount, $Comparison.WorsenedWarningCount, $Comparison.ResolvedWarningCount
        $relDeltaVal = if ($Comparison.RelationshipComparisons.Count -gt 0) {
            "孤立 {0}{1:#,##0}行 (FK {2}件)" -f $(if ($Comparison.OrphanRowDelta -gt 0) { '+' } else { '' }), $Comparison.OrphanRowDelta, $Comparison.RelationshipComparisons.Count
        } else { 'FK検出なし / 旧形式' }
        $dbMetaVal = if ($Comparison.DbRestartDetected) { 'DB再起動/FO検出(要注意)' } elseif ($Comparison.HasDbServerStatus) { ("DB負荷 {0}指標 / DDL差分 {1}" -f $Comparison.ServerMetricComparisons.Count, $Comparison.DbSchemaDiffs.Count) } elseif ($Comparison.DbSchemaDiffs.Count -gt 0) { ("DDL/INDEX差分 {0}件" -f $Comparison.DbSchemaDiffs.Count) } else { ("{0}×{1}" -f $Comparison.BeforeSourceMode, $Comparison.AfterSourceMode) }

        $kpiRow2 = @(
            @{ Col1 = 1;  Col2 = 2;  Label = '計測間隔 / 平均増減速度';     Value = $elapsedVal; Fmt = '@'; Warn = $Comparison.ClockReversed; FontSize = 10.5 },
            @{ Col1 = 3;  Col2 = 4;  Label = '推定InnoDB / 実DB容量増減';   Value = $innoKpiVal; Fmt = '@'; Warn = $false; FontSize = 10.0 },
            @{ Col1 = 5;  Col2 = 6;  Label = '性能警告推移 (新規/悪化/解消)'; Value = $warnDeltaVal; Fmt = '@'; Warn = ($Comparison.NewWarningCount -gt 0 -or $Comparison.WorsenedWarningCount -gt 0); FontSize = 11.0 },
            @{ Col1 = 7;  Col2 = 8;  Label = 'FK孤立行増減 / リレーション'; Value = $relDeltaVal; Fmt = '@'; Warn = ($Comparison.OrphanRowDelta -gt 0); FontSize = 10.5 },
            @{ Col1 = 9;  Col2 = 10; Label = 'DBサーバー / DDL差分状態';    Value = $dbMetaVal; Fmt = '@'; Warn = ($Comparison.DbRestartDetected -or $Comparison.DbSchemaDiffs.Count -gt 0); FontSize = 10.5 },
            @{ Col1 = 11; Col2 = 12; Label = '比較照合精度';                Value = $Comparison.PrecisionLabel; Fmt = '@'; Warn = $false; FontSize = 9.5 }
        )
        Write-CmpKpiBand -Sheet $wsSummary -LabelRow 7 -Kpis $kpiRow2
        $wsSummary.Rows.Item(9).RowHeight = 10

        # Section 1: Category Breakdown Table (Mutually Exclusive -> Sums to 100%)
        Set-SectionHeader -Sheet $wsSummary -Row 10 -StartCol 1 -EndCol 12 -Text '1. 変更区分別サマリー (重複なし内訳 / テスト開始前 vs テスト開始後)'
        $bucketDefs = @(
            @{ Status = '追加（新規テーブル）';   Desc = 'テスト開始前になく、開始後に新規追加されたテーブル' },
            @{ Status = '追加（レコード増加）';   Desc = '既存テーブルでレコード数が増加 (INSERT等)' },
            @{ Status = '追加・更新';             Desc = 'レコード増加に加え、既存値の更新・削除・スキーマ変更も検出 (更新・追加両シートに掲載)' },
            @{ Status = '更新（データ更新）';     Desc = 'レコード数変動なしでカラム値ハッシュ・統計・分布・数値・日時・サイズに変化 (UPDATE/行入替等)' },
            @{ Status = '更新（レコード減少）';   Desc = '既存テーブルでレコード数が減少 (DELETE/パージ等・要注意)' },
            @{ Status = '削除（テーブル消失）';   Desc = 'テスト開始前に存在したが開始後に未検出 (テーブル削除/出力漏れ・要注意)' },
            @{ Status = '変更なし';               Desc = 'レコード数・サイズ・全カラム統計・分布・数値・日時統計・内容ハッシュすべて完全一致' }
        )

        $totalTbl = [Math]::Max(1, $Comparison.TotalUniqueTableCount)
        $mCat = New-Object 'object[,]' $bucketDefs.Count, 12
        for ($i = 0; $i -lt $bucketDefs.Count; $i++) {
            $stName = $bucketDefs[$i].Status
            $matched = @($Comparison.AllTables | Where-Object { $_.PrimaryStatus -eq $stName })
            $bRows = 0L; $aRows = 0L; $dRows = 0L; $dSize = 0.0; $dInno = 0.0; $diffCnt = 0
            foreach ($mt in $matched) {
                $bRows += $mt.BeforeRows; $aRows += $mt.AfterRows; $dRows += $mt.RowDelta
                $dSize += $mt.SizeDeltaKB; $dInno += $mt.InnoDbDeltaKB; $diffCnt += $mt.ColumnDiffs.Count
            }
            $mCat[$i, 0]  = $stName
            $mCat[$i, 1]  = [double]$matched.Count
            $mCat[$i, 2]  = [double]($matched.Count / $totalTbl)
            $mCat[$i, 3]  = [double]$bRows
            $mCat[$i, 4]  = [double]$aRows
            $mCat[$i, 5]  = [double]$dRows
            $mCat[$i, 6]  = [Math]::Round($dSize, 2)
            $mCat[$i, 7]  = [Math]::Round($dInno, 1)
            $mCat[$i, 8]  = [double]$diffCnt
            $mCat[$i, 9]  = if ($stName -match '(削除|減少)' -and $matched.Count -gt 0) { '要注意' } elseif ($stName -eq '変更なし') { '変化なし' } elseif ($matched.Count -gt 0) { '変更検出' } else { '-' }
            $mCat[$i, 10] = if ($stName -match '^追加') { '04_追加のあったテーブル' } elseif ($stName -match '^更新') { '03_更新のあったテーブル' } elseif ($stName -match '^削除') { '02_全テーブル / 03_更新' } else { '05_変更のなかったテーブル' }
            $mCat[$i, 11] = $bucketDefs[$i].Desc
        }
        $catDataStart = 12
        $catEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsSummary -StartRow 11 -StartCol 1 `
            -Headers @('判定区分', 'テーブル数', '構成比', '開始前レコード数', '開始後レコード数', 'レコード増減', 'CSV増減(KB)', 'InnoDB増減(KB)', '検出差分項目数', '状態', '参照シート', '判定基準・説明') `
            -DataMatrix $mCat -RowCount $bucketDefs.Count `
            -ColumnFormats @{ 1 = '@'; 2 = '#,##0'; 3 = '0.0%'; 4 = '#,##0'; 5 = '#,##0'; 6 = '+#,##0;-#,##0;0'; 7 = '+#,##0.00;-#,##0.00;0.00'; 8 = '+#,##0.0;-#,##0.0;0.0'; 9 = '#,##0'; 10 = '@'; 11 = '@'; 12 = '@' } `
            -CenterCols @(1, 10, 11) -RightCols @(2, 3, 4, 5, 6, 7, 8, 9) -WrapCols @(12)

        for ($i = 0; $i -lt $bucketDefs.Count; $i++) {
            $rIdx = $catDataStart + $i
            $stName = $bucketDefs[$i].Status
            $c1 = $wsSummary.Cells.Item($rIdx, 1)
            if ($stName -match '(削除|減少)') {
                if ([double]$mCat[$i, 1] -gt 0) {
                    $c1.Font.Color = $warnOle; $c1.Font.Bold = $true; $c1.Interior.Color = $warnBgOle
                    $c10 = $wsSummary.Cells.Item($rIdx, 10)
                    $c10.Font.Color = $warnOle; $c10.Font.Bold = $true; $c10.Interior.Color = $warnBgOle
                }
                else { $c1.Font.Color = $mutedOle }
            }
            elseif ($stName -match '^追加') { $c1.Font.Color = $navyOle; $c1.Font.Bold = $true; $c1.Interior.Color = $subBgOle }
            elseif ($stName -match '^更新') { $c1.Font.Color = $textOle; $c1.Font.Bold = $true; $c1.Interior.Color = $subBgOle }
            else { $c1.Font.Color = $mutedOle }

            foreach ($dc in @(6, 7, 8)) {
                $dv = [double]$mCat[$i, ($dc - 1)]
                $dcCell = $wsSummary.Cells.Item($rIdx, $dc)
                if ($dv -lt -1e-9) { $dcCell.Font.Color = $warnOle; $dcCell.Font.Bold = $true }
                elseif ($dv -gt 1e-9) { $dcCell.Font.Color = $navyOle; $dcCell.Font.Bold = $true }
                else { $dcCell.Font.Color = $mutedOle }
            }
        }

        $catTotalRow = $catEnd + 1
        Write-ComparisonTotalRow -Sheet $wsSummary -Row $catTotalRow `
            -Values @('合計', [double]$Comparison.TotalUniqueTableCount, 1.0, [double]$Comparison.BeforeTotalRows, [double]$Comparison.AfterTotalRows, [double]$Comparison.TotalRowDelta, [double]$Comparison.TotalSizeDeltaKB, [double]$Comparison.TotalInnoDbDeltaKB, [double]$Comparison.AllColumnDiffs.Count, '', '', '全比較対象テーブル合計') `
            -ColumnFormats @{ 2 = '#,##0'; 3 = '0.0%'; 4 = '#,##0'; 5 = '#,##0'; 6 = '+#,##0;-#,##0;0'; 7 = '+#,##0.00;-#,##0.00;0.00'; 8 = '+#,##0.0;-#,##0.0;0.0'; 9 = '#,##0' } `
            -CenterCols @(1) -RightCols @(2, 3, 4, 5, 6, 7, 8, 9)

        # Section 2: Top 15 Changed Tables by Absolute Row Delta / Changed Columns
        $topSecRow = $catTotalRow + 2
        Set-SectionHeader -Sheet $wsSummary -Row $topSecRow -StartCol 1 -EndCol 12 -Text '2. 件数推移・変更影響度 TOP15 テーブル (行変動内訳・変更/不変カラム数付き)'
        $topChanged = @($Comparison.AllTables | Where-Object { -not $_.IsUnchanged } | Sort-Object @{ Expression = { [Math]::Abs($_.RowDelta) }; Descending = $true }, @{ Expression = { $_.ChangedColumnCount }; Descending = $true } | Select-Object -First 15)
        $mTopC = New-Object 'object[,]' ([Math]::Max(1, $topChanged.Count)), 12
        for ($i = 0; $i -lt $topChanged.Count; $i++) {
            $t = $topChanged[$i]
            $mTopC[$i, 0]  = $i + 1
            $mTopC[$i, 1]  = $t.TableName
            $mTopC[$i, 2]  = $t.PrimaryStatus
            $mTopC[$i, 3]  = [double]$t.BeforeRows
            $mTopC[$i, 4]  = [double]$t.AfterRows
            $mTopC[$i, 5]  = [double]$t.RowDelta
            $mTopC[$i, 6]  = if ($t.HasRowGrowthRate) { [double]$t.RowGrowthRate } elseif ($t.AfterRows -gt 0 -and $t.BeforeRows -eq 0) { '新規' } else { '-' }
            $mTopC[$i, 7]  = [double]$t.InnoDbDeltaKB
            $mTopC[$i, 8]  = ("変更{0} / 不変{1}" -f $t.ChangedColumnCount, $t.UnchangedColumnCount)
            $mTopC[$i, 9]  = if ($t.HasRowChurn) { $t.ChurnSummary } else { '-' }
            $mTopC[$i, 10] = $t.ChangeDetail
            $mTopC[$i, 11] = $t.AlertNote
        }
        $topEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsSummary -StartRow ($topSecRow + 1) -StartCol 1 `
            -Headers @('順位', 'テーブル名', '判定区分', '開始前件数', '開始後件数', 'レコード増減', '増減率', 'InnoDB増減(KB)', '変更/不変列', '行変動内訳 (挿入/更新/削除)', '変更内訳', '警告・注意点') `
            -DataMatrix $mTopC -RowCount $topChanged.Count `
            -ColumnFormats @{ 1 = '#,##0'; 2 = '@'; 3 = '@'; 4 = '#,##0'; 5 = '#,##0'; 6 = '+#,##0;-#,##0;0'; 7 = '+0.0%;-0.0%;0.0%'; 8 = '+#,##0.0;-#,##0.0;0.0'; 9 = '@'; 10 = '@'; 11 = '@'; 12 = '@' } `
            -CenterCols @(1, 3, 9) -RightCols @(4, 5, 6, 7, 8) -WrapCols @(10, 11, 12)
        if ($topChanged.Count -gt 0) {
            Set-ComparisonRowStyles -Sheet $wsSummary -StartRow ($topSecRow + 2) -Items $topChanged -StatusCol 3 -AlertCol 12 -SignedDeltaCols @(6, 8) -WrapTextCols @(10, 11, 12) -DataMatrix $mTopC
        }

        # Charts on Comparison Summary Sheet (3 Charts: Category Doughnut, Row Delta Bar, Before vs After Transition Bar!)
        $chartTopRow = $topEnd + 2
        $chartTopPt = $wsSummary.Rows.Item($chartTopRow).Top
        $chartHeight = 250.0
        [void](Add-StyledExcelChart -Sheet $wsSummary -Left 6 -Top $chartTopPt -Width 440 -Height $chartHeight `
            -ChartType -4120 -Title '変更区分別テーブル数構成' `
            -CategoryAddress "A${catDataStart}:A${catEnd}" -ValueAddresses @("B${catDataStart}:B${catEnd}") `
            -SeriesNames @('テーブル数') -ShowLegend $true -ShowDataLabels $true)

        $chartsBottomPt = $chartTopPt + $chartHeight
        if ($topChanged.Count -ge 1) {
            $barStart = $topSecRow + 2
            $barEnd = $barStart + $topChanged.Count - 1
            [void](Add-StyledExcelChart -Sheet $wsSummary -Left 460 -Top $chartTopPt -Width 480 -Height $chartHeight `
                -ChartType 57 -Title 'テーブル別レコード増減数 TOP15' `
                -CategoryAddress "B${barStart}:B${barEnd}" -ValueAddresses @("F${barStart}:F${barEnd}") `
                -SeriesNames @('レコード増減') -SeriesColors @($script:THEME.AccentHex) `
                -ReverseCategories $true -ShowDataLabels $true -ColorNegativeBurgundy $true)

            $trEnd = $barStart + [Math]::Min(12, $topChanged.Count) - 1
            $chart2TopPt = $chartTopPt + $chartHeight + 14
            [void](Add-StyledExcelChart -Sheet $wsSummary -Left 6 -Top $chart2TopPt -Width 934 -Height 245 `
                -ChartType 57 -Title 'テスト前後 レコード数推移比較 TOP12 (開始前 vs 開始後)' `
                -CategoryAddress "B${barStart}:B${trEnd}" -ValueAddresses @("D${barStart}:D${trEnd}", "E${barStart}:E${trEnd}") `
                -SeriesNames @('開始前レコード数', '開始後レコード数') -SeriesColors @('#8C8C8C', $script:THEME.AccentHex) `
                -ShowLegend $true -ReverseCategories $true -ShowDataLabels $true)
            $chartsBottomPt = $chart2TopPt + 245
        }

        # Section 3: Comparison Insights (placed dynamically below charts)
        $insRow = Get-RowBelowPoints -Sheet $wsSummary -StartRow ($chartTopRow + 1) -TargetTopPt ($chartsBottomPt + 16)
        Set-SectionHeader -Sheet $wsSummary -Row $insRow -StartCol 1 -EndCol 12 -Text '3. テスト前後比較 総合所見・性能試験チェックポイント'
        $insRow++
        foreach ($ins in $Comparison.Insights) {
            $rIns = $wsSummary.Range($wsSummary.Cells.Item($insRow, 1), $wsSummary.Cells.Item($insRow, 12))
            $rIns.Merge()
            $rIns.NumberFormat = '@'
            [AuroraAnalyzer.ComHelper]::SetValue2($wsSummary.Cells.Item($insRow, 1), $ins)
            $rIns.Font.Name = $script:THEME.FontName; $rIns.Font.Size = 9.5
            $isWarnIns = ($ins -match '要注意')
            $rIns.Font.Color = if ($isWarnIns) { $warnOle } else { $textOle }
            if ($isWarnIns) { $rIns.Font.Bold = $true }
            $rIns.WrapText = $true; $rIns.HorizontalAlignment = -4131; $rIns.VerticalAlignment = -4108; $rIns.IndentLevel = 1
            $rIns.Interior.Color = if ($isWarnIns) { $warnBgOle } else { $zebraOle }
            $rIns.Borders.Item(7).LineStyle = 1; $rIns.Borders.Item(7).Color = if ($isWarnIns) { $warnOle } else { $navyOle }; $rIns.Borders.Item(7).Weight = 3
            $rIns.Borders.Item(9).LineStyle = 1; $rIns.Borders.Item(9).Color = $borderOle; $rIns.Borders.Item(9).Weight = 2
            $wsSummary.Rows.Item($insRow).RowHeight = Get-EstimatedRowHeight -Texts @($ins) -ColWidths @(210.0) -MinHeight 26.0 -MaxHeight 85.0
            $insRow++
        }

        Format-ExcelSheetCommon -Sheet $wsSummary -FreezeRow 0 -Landscape $true

        # ======================================================================
        #  SHEET 2: 02_全テーブル比較一覧
        # ======================================================================
        Write-CmpLog "シート作成中: 02_全テーブル比較一覧..." 56
        Set-ColumnWidths -Sheet $wsAll -Widths @(5, 22, 20, 15, 13, 13, 12, 10, 11, 26, 9, 9, 9, 9, 9, 12, 12, 12, 12, 13, 13, 13, 38, 28)
        Set-SheetTitleBanner -Sheet $wsAll -Title '全テーブル テスト開始前・開始後 比較一覧 (件数推移・行変動・変更/不変列数・ハッシュ照合)' -Subtitle $subInfo -EndCol 24

        $allItems = @($Comparison.AllTables)
        $mAll = New-Object 'object[,]' ([Math]::Max(1, $allItems.Count)), 24
        for ($i = 0; $i -lt $allItems.Count; $i++) {
            $t = $allItems[$i]
            $mAll[$i, 0]  = $i + 1
            $mAll[$i, 1]  = $t.TableName
            $mAll[$i, 2]  = $t.PrimaryStatus
            $mAll[$i, 3]  = $t.InferredRole
            $mAll[$i, 4]  = [double]$t.BeforeRows
            $mAll[$i, 5]  = [double]$t.AfterRows
            $mAll[$i, 6]  = [double]$t.RowDelta
            $mAll[$i, 7]  = if ($t.HasRowGrowthRate) { [double]$t.RowGrowthRate } elseif ($t.AfterRows -gt 0 -and $t.BeforeRows -eq 0) { '新規' } else { '-' }
            $mAll[$i, 8]  = if ($t.HasVelocity) { [Math]::Round($t.RowsPerMinute, 1) } else { '-' }
            $mAll[$i, 9]  = if ($t.HasRowChurn) { $t.ChurnSummary } else { '-' }
            $mAll[$i, 10] = [double]$t.BeforeCols
            $mAll[$i, 11] = [double]$t.AfterCols
            $mAll[$i, 12] = [double]$t.ColDelta
            $mAll[$i, 13] = [double]$t.ChangedColumnCount
            $mAll[$i, 14] = [double]$t.UnchangedColumnCount
            $mAll[$i, 15] = $t.ContentHashStatus
            $mAll[$i, 16] = [double]$t.BeforeSizeKB
            $mAll[$i, 17] = [double]$t.AfterSizeKB
            $mAll[$i, 18] = [double]$t.SizeDeltaKB
            $mAll[$i, 19] = [double]$t.BeforeInnoDbKB
            $mAll[$i, 20] = [double]$t.AfterInnoDbKB
            $mAll[$i, 21] = [double]$t.InnoDbDeltaKB
            $mAll[$i, 22] = $t.DiffSummary
            $mAll[$i, 23] = $t.AlertNote
        }
        $allEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsAll -StartRow 4 -StartCol 1 `
            -Headers @('No', 'テーブル名', '判定区分', '推定役割', '開始前レコード数', '開始後レコード数', 'レコード増減', '増減率', '増減速度(件/分)', '行変動内訳 (挿入/更新/削除)', '開始前列数', '開始後列数', '列数増減', '変更列数', '不変列数', '内容ハッシュ', '開始前CSV(KB)', '開始後CSV(KB)', 'CSV増減(KB)', '開始前InnoDB(KB)', '開始後InnoDB(KB)', 'InnoDB増減(KB)', '主な変更内容・所見', '警告・注意点') `
            -DataMatrix $mAll -RowCount $allItems.Count `
            -ColumnFormats @{ 1 = '#,##0'; 2 = '@'; 3 = '@'; 5 = '#,##0'; 6 = '#,##0'; 7 = '+#,##0;-#,##0;0'; 8 = '+0.0%;-0.0%;0.0%'; 9 = '+#,##0.0;-#,##0.0;0.0'; 10 = '@'; 11 = '#,##0'; 12 = '#,##0'; 13 = '+#,##0;-#,##0;0'; 14 = '#,##0'; 15 = '#,##0'; 16 = '@'; 17 = '#,##0.00'; 18 = '#,##0.00'; 19 = '+#,##0.00;-#,##0.00;0.00'; 20 = '#,##0.0'; 21 = '#,##0.0'; 22 = '+#,##0.0;-#,##0.0;0.0'; 23 = '@'; 24 = '@' } `
            -CenterCols @(1, 3, 4, 11, 12, 13, 14, 15, 16) -RightCols @(5, 6, 7, 8, 9, 17, 18, 19, 20, 21, 22) -WrapCols @(10, 23, 24)
        if ($allItems.Count -gt 0) {
            Set-ComparisonRowStyles -Sheet $wsAll -StartRow 5 -Items $allItems -StatusCol 3 -AlertCol 24 -SignedDeltaCols @(7, 9, 13, 19, 22) -WrapTextCols @(10, 23, 24) -DataMatrix $mAll
            $totGrowth = if ($Comparison.HasTotalRowGrowthRate) { [double]$Comparison.TotalRowGrowthRate } else { '-' }
            $totVel = if ($Comparison.HasElapsedTime) { [Math]::Round($Comparison.TotalRowsPerMinute, 1) } else { '-' }
            $totChurn = if ($Comparison.HasAnyRowChurn) { ("+{0:#,##0} / 更新{1:#,##0} / -{2:#,##0}" -f $Comparison.TotalInsertedRows, $Comparison.TotalUpdatedRows, $Comparison.TotalDeletedRows) } else { '-' }
            Write-ComparisonTotalRow -Sheet $wsAll -Row ($allEnd + 1) `
                -Values @('', '合計', ("計 {0} テーブル" -f $allItems.Count), '', [double]$Comparison.BeforeTotalRows, [double]$Comparison.AfterTotalRows, [double]$Comparison.TotalRowDelta, $totGrowth, $totVel, $totChurn, '', '', '', [double]$Comparison.ChangedColumnTotal, [double]$Comparison.UnchangedColumnTotal, '', [Math]::Round($Comparison.BeforeTotalSizeKB, 2), [Math]::Round($Comparison.AfterTotalSizeKB, 2), [double]$Comparison.TotalSizeDeltaKB, [Math]::Round($Comparison.BeforeTotalInnoDbKB, 1), [Math]::Round($Comparison.AfterTotalInnoDbKB, 1), [double]$Comparison.TotalInnoDbDeltaKB, '', '') `
                -ColumnFormats @{ 5 = '#,##0'; 6 = '#,##0'; 7 = '+#,##0;-#,##0;0'; 8 = '+0.0%;-0.0%;0.0%'; 9 = '+#,##0.0;-#,##0.0;0.0'; 14 = '#,##0'; 15 = '#,##0'; 17 = '#,##0.00'; 18 = '#,##0.00'; 19 = '+#,##0.00;-#,##0.00;0.00'; 20 = '#,##0.0'; 21 = '#,##0.0'; 22 = '+#,##0.0;-#,##0.0;0.0' } `
                -CenterCols @(3, 14, 15) -RightCols @(5, 6, 7, 8, 9, 17, 18, 19, 20, 21, 22)
        }
        Format-ExcelSheetCommon -Sheet $wsAll -FreezeRow 4 -AutoFilterRange "A4:X$allEnd" -Landscape $true

        # ======================================================================
        #  SHEET 3: 03_更新のあったテーブル
        # ======================================================================
        Write-CmpLog "シート作成中: 03_更新のあったテーブル..." 64
        Set-ColumnWidths -Sheet $wsUpdated -Widths @(5, 22, 20, 15, 13, 13, 12, 24, 11, 13, 13, 12, 12, 9, 24, 24, 38, 28)
        Set-SheetTitleBanner -Sheet $wsUpdated -Title '更新のあったテーブル一覧 (既存データ更新・レコード減少・追加＋更新併用 / 変更・不変カラム明示)' -Subtitle $subInfo -EndCol 18

        $updItems = @($Comparison.UpdatedTables)
        $mUpd = New-Object 'object[,]' ([Math]::Max(1, $updItems.Count)), 18
        for ($i = 0; $i -lt $updItems.Count; $i++) {
            $t = $updItems[$i]
            $mUpd[$i, 0]  = $i + 1
            $mUpd[$i, 1]  = $t.TableName
            $mUpd[$i, 2]  = $t.PrimaryStatus
            $mUpd[$i, 3]  = $t.InferredRole
            $mUpd[$i, 4]  = [double]$t.BeforeRows
            $mUpd[$i, 5]  = [double]$t.AfterRows
            $mUpd[$i, 6]  = [double]$t.RowDelta
            $mUpd[$i, 7]  = if ($t.HasRowChurn) { $t.ChurnSummary } else { '-' }
            $mUpd[$i, 8]  = $t.ContentHashStatus
            $mUpd[$i, 9]  = [double]$t.BeforeAvgRowBytes
            $mUpd[$i, 10] = [double]$t.AfterAvgRowBytes
            $mUpd[$i, 11] = [double]$t.AvgRowBytesDelta
            $mUpd[$i, 12] = [double]$t.SizeDeltaKB
            $mUpd[$i, 13] = [double]$t.ChangedColumnCount
            $mUpd[$i, 14] = $t.ChangedColumnsSummary
            $mUpd[$i, 15] = $t.UnchangedColumnsSummary
            $mUpd[$i, 16] = $t.DiffSummary
            $mUpd[$i, 17] = $t.AlertNote
        }
        $updEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsUpdated -StartRow 4 -StartCol 1 `
            -Headers @('No', 'テーブル名', '判定区分', '推定役割', '開始前レコード数', '開始後レコード数', 'レコード増減', '行変動内訳 (挿入/更新/削除)', '内容ハッシュ', '開始前平均行長(B)', '開始後平均行長(B)', '平均行長増減(B)', 'CSV増減(KB)', '変化列数', '変更検出カラム', '変更なしカラム', '具体的な変更内容', '警告・注意点') `
            -DataMatrix $mUpd -RowCount $updItems.Count `
            -ColumnFormats @{ 1 = '#,##0'; 2 = '@'; 3 = '@'; 5 = '#,##0'; 6 = '#,##0'; 7 = '+#,##0;-#,##0;0'; 8 = '@'; 9 = '@'; 10 = '#,##0.0'; 11 = '#,##0.0'; 12 = '+#,##0.0;-#,##0.0;0.0'; 13 = '+#,##0.00;-#,##0.00;0.00'; 14 = '#,##0'; 15 = '@'; 16 = '@'; 17 = '@'; 18 = '@' } `
            -CenterCols @(1, 3, 4, 9, 14) -RightCols @(5, 6, 7, 10, 11, 12, 13) -WrapCols @(8, 15, 16, 17, 18)
        if ($updItems.Count -gt 0) {
            Set-ComparisonRowStyles -Sheet $wsUpdated -StartRow 5 -Items $updItems -StatusCol 3 -AlertCol 18 -SignedDeltaCols @(7, 12, 13) -WrapTextCols @(8, 15, 16, 17, 18) -DataMatrix $mUpd
        }

        $updNextRow = $updEnd + 2
        if ($Comparison.DeletedTables.Count -gt 0) {
            Set-SectionHeader -Sheet $wsUpdated -Row $updNextRow -StartCol 1 -EndCol 18 -Text '【要注意】削除・消失テーブル一覧 (テスト開始前に存在し、テスト開始後の分析ファイルに未検出)'
            $delItems = @($Comparison.DeletedTables)
            $mDel = New-Object 'object[,]' $delItems.Count, 8
            for ($i = 0; $i -lt $delItems.Count; $i++) {
                $dt = $delItems[$i]
                $mDel[$i, 0] = $i + 1
                $mDel[$i, 1] = $dt.TableName
                $mDel[$i, 2] = $dt.PrimaryStatus
                $mDel[$i, 3] = $dt.InferredRole
                $mDel[$i, 4] = [double]$dt.BeforeRows
                $mDel[$i, 5] = [double]$dt.BeforeCols
                $mDel[$i, 6] = [double]$dt.BeforeSizeKB
                $mDel[$i, 7] = $dt.AlertNote
            }
            $delHdrRow = $updNextRow + 1
            $delEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsUpdated -StartRow $delHdrRow -StartCol 1 `
                -Headers @('No', 'テーブル名', '判定区分', '推定役割', '開始前レコード数', '開始前列数', '開始前CSV(KB)', '警告・注意点') `
                -DataMatrix $mDel -RowCount $delItems.Count `
                -ColumnFormats @{ 1 = '#,##0'; 2 = '@'; 3 = '@'; 5 = '#,##0'; 6 = '#,##0'; 7 = '#,##0.00'; 8 = '@' } `
                -CenterCols @(1, 3, 4, 6) -RightCols @(5, 7) -WrapCols @(8)
            Set-ComparisonRowStyles -Sheet $wsUpdated -StartRow ($delHdrRow + 1) -Items $delItems -StatusCol 3 -AlertCol 8 -WrapTextCols @(8) -DataMatrix $mDel
            $delHdrMerge = $wsUpdated.Range($wsUpdated.Cells.Item($delHdrRow, 8), $wsUpdated.Cells.Item($delHdrRow, 11))
            $delHdrMerge.Merge()
            $delHdrMerge.Interior.Color = $navyOle
            $delHdrMerge.Font.Color = ConvertTo-ExcelOleColor $script:THEME.WhiteHex
            $delHdrMerge.Font.Bold = $true
            $delHdrMerge.HorizontalAlignment = -4108
            for ($i = 0; $i -lt $delItems.Count; $i++) {
                $drIdx = $delHdrRow + 1 + $i
                $dm = $wsUpdated.Range($wsUpdated.Cells.Item($drIdx, 8), $wsUpdated.Cells.Item($drIdx, 11))
                $dm.Merge()
                $dm.Font.Color = $warnOle; $dm.Font.Bold = $true; $dm.Interior.Color = $warnBgOle
                foreach ($edge in @(7, 8, 9, 10)) { try { $dm.Borders.Item($edge).LineStyle = 1; $dm.Borders.Item($edge).Weight = 2; $dm.Borders.Item($edge).Color = $borderOle } catch { } }
                $wsUpdated.Rows.Item($drIdx).RowHeight = Get-EstimatedRowHeight -Texts @([string]$delItems[$i].AlertNote) -ColWidths @(46.0) -MinHeight 24.0
            }
            $updNextRow = $delEnd + 2
        }

        if ($updItems.Count -ge 1) {
            $uChartCnt = [Math]::Min(15, $updItems.Count)
            $uChartEnd = 4 + $uChartCnt
            $uChartTop = $wsUpdated.Rows.Item($updNextRow).Top
            [void](Add-StyledExcelChart -Sheet $wsUpdated -Left 6 -Top $uChartTop -Width 540 -Height 245 `
                -ChartType 57 -Title '更新テーブル別 変更検出カラム数 TOP15' `
                -CategoryAddress "B5:B${uChartEnd}" -ValueAddresses @("N5:N${uChartEnd}") `
                -SeriesNames @('変更検出カラム数') -SeriesColors @('#4D4D4D') -ReverseCategories $true -ShowDataLabels $true)
        }
        Format-ExcelSheetCommon -Sheet $wsUpdated -FreezeRow 4 -AutoFilterRange "A4:R$updEnd" -Landscape $true

        # ======================================================================
        #  SHEET 4: 04_追加のあったテーブル
        # ======================================================================
        Write-CmpLog "シート作成中: 04_追加のあったテーブル..." 72
        Set-ColumnWidths -Sheet $wsAdded -Widths @(5, 22, 18, 15, 13, 13, 13, 10, 11, 24, 7, 12, 12, 12, 13, 18, 36, 26)
        Set-SheetTitleBanner -Sheet $wsAdded -Title '追加のあったテーブル一覧 (新規テーブル追加 ＆ レコード増加テーブル)' -Subtitle $subInfo -EndCol 18

        $addItems = @($Comparison.AddedTables)
        $mAdd = New-Object 'object[,]' ([Math]::Max(1, $addItems.Count)), 18
        $sumAddBefore = 0L; $sumAddAfter = 0L; $sumAddDelta = 0L; $sumAddSizeD = 0.0; $sumAddInnoD = 0.0
        for ($i = 0; $i -lt $addItems.Count; $i++) {
            $t = $addItems[$i]
            $sumAddBefore += $t.BeforeRows; $sumAddAfter += $t.AfterRows; $sumAddDelta += $t.RowDelta
            $sumAddSizeD += $t.SizeDeltaKB; $sumAddInnoD += $t.InnoDbDeltaKB
            $mAdd[$i, 0]  = $i + 1
            $mAdd[$i, 1]  = $t.TableName
            $mAdd[$i, 2]  = $t.AddedType
            $mAdd[$i, 3]  = $t.InferredRole
            $mAdd[$i, 4]  = [double]$t.BeforeRows
            $mAdd[$i, 5]  = [double]$t.AfterRows
            $mAdd[$i, 6]  = [double]$t.RowDelta
            $mAdd[$i, 7]  = if ($t.HasRowGrowthRate) { [double]$t.RowGrowthRate } elseif ($t.AfterRows -gt 0 -and $t.BeforeRows -eq 0) { '新規' } else { '-' }
            $mAdd[$i, 8]  = if ($t.HasVelocity) { [Math]::Round($t.RowsPerMinute, 1) } else { '-' }
            $mAdd[$i, 9]  = if ($t.HasRowChurn) { $t.ChurnSummary } else { '-' }
            $mAdd[$i, 10] = [double]$t.AfterCols
            $mAdd[$i, 11] = [double]$t.BeforeSizeKB
            $mAdd[$i, 12] = [double]$t.AfterSizeKB
            $mAdd[$i, 13] = [double]$t.SizeDeltaKB
            $mAdd[$i, 14] = [double]$t.InnoDbDeltaKB
            $mAdd[$i, 15] = if ([string]::IsNullOrEmpty($t.CandidatePks)) { '-' } else { $t.CandidatePks }
            $mAdd[$i, 16] = $t.DiffSummary
            $mAdd[$i, 17] = $t.AlertNote
        }
        $addEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsAdded -StartRow 4 -StartCol 1 `
            -Headers @('No', 'テーブル名', '追加種別', '推定役割', '開始前レコード数', '開始後レコード数', '追加レコード数', '増加率', '増加速度(件/分)', '行変動内訳', '列数', '開始前CSV(KB)', '開始後CSV(KB)', 'CSV増加(KB)', 'InnoDB増加(KB)', 'PK/UK候補', '変更サマリー', '警告・注意点') `
            -DataMatrix $mAdd -RowCount $addItems.Count `
            -ColumnFormats @{ 1 = '#,##0'; 2 = '@'; 3 = '@'; 5 = '#,##0'; 6 = '#,##0'; 7 = '+#,##0;-#,##0;0'; 8 = '+0.0%;-0.0%;0.0%'; 9 = '+#,##0.0;-#,##0.0;0.0'; 10 = '@'; 11 = '#,##0'; 12 = '#,##0.00'; 13 = '#,##0.00'; 14 = '+#,##0.00;-#,##0.00;0.00'; 15 = '+#,##0.0;-#,##0.0;0.0'; 16 = '@'; 17 = '@'; 18 = '@' } `
            -CenterCols @(1, 3, 4, 11) -RightCols @(5, 6, 7, 8, 9, 12, 13, 14, 15) -WrapCols @(10, 17, 18)
        $addChartAnchorRow = $addEnd + 2
        if ($addItems.Count -gt 0) {
            Set-ComparisonRowStyles -Sheet $wsAdded -StartRow 5 -Items $addItems -StatusCol 3 -AlertCol 18 -SignedDeltaCols @(7, 9, 14, 15) -WrapTextCols @(10, 17, 18) -DataMatrix $mAdd
            $addTotGrowth = if ($sumAddBefore -gt 0) { [double]($sumAddDelta / $sumAddBefore) } else { '-' }
            Write-ComparisonTotalRow -Sheet $wsAdded -Row ($addEnd + 1) `
                -Values @('', '合計', ("計 {0} テーブル" -f $addItems.Count), '', [double]$sumAddBefore, [double]$sumAddAfter, [double]$sumAddDelta, $addTotGrowth, '', '', '', '', '', [Math]::Round($sumAddSizeD, 2), [Math]::Round($sumAddInnoD, 1), '', '', '') `
                -ColumnFormats @{ 5 = '#,##0'; 6 = '#,##0'; 7 = '+#,##0;-#,##0;0'; 8 = '+0.0%;-0.0%;0.0%'; 14 = '+#,##0.00;-#,##0.00;0.00'; 15 = '+#,##0.0;-#,##0.0;0.0' } `
                -CenterCols @(3) -RightCols @(5, 6, 7, 8, 14, 15)
            $addChartAnchorRow = $addEnd + 3

            $aChartCnt = [Math]::Min(15, $addItems.Count)
            $aChartEnd = 4 + $aChartCnt
            $aChartTop = $wsAdded.Rows.Item($addChartAnchorRow).Top
            [void](Add-StyledExcelChart -Sheet $wsAdded -Left 6 -Top $aChartTop -Width 540 -Height 250 `
                -ChartType 57 -Title 'テーブル別 追加レコード数 TOP15' `
                -CategoryAddress "B5:B${aChartEnd}" -ValueAddresses @("G5:G${aChartEnd}") `
                -SeriesNames @('追加レコード数') -SeriesColors @($script:THEME.AccentHex) -ReverseCategories $true -ShowDataLabels $true)
        }
        Format-ExcelSheetCommon -Sheet $wsAdded -FreezeRow 4 -AutoFilterRange "A4:R$addEnd" -Landscape $true

        # ======================================================================
        #  SHEET 5: 05_変更のなかったテーブル
        # ======================================================================
        Write-CmpLog "シート作成中: 05_変更のなかったテーブル..." 78
        Set-ColumnWidths -Sheet $wsUnchanged -Widths @(5, 24, 14, 15, 16, 13, 8, 14, 15, 12, 11, 20, 20, 30)
        Set-SheetTitleBanner -Sheet $wsUnchanged -Title '変更のなかったテーブル一覧 (開始前後で件数・サイズ・統計・値ハッシュ完全一致)' -Subtitle $subInfo -EndCol 14

        $uncItems = @($Comparison.UnchangedTables)
        $mUnc = New-Object 'object[,]' ([Math]::Max(1, $uncItems.Count)), 14
        $sumUncRows = 0L; $sumUncSize = 0.0; $sumUncInno = 0.0
        for ($i = 0; $i -lt $uncItems.Count; $i++) {
            $t = $uncItems[$i]
            $sumUncRows += $t.AfterRows; $sumUncSize += $t.AfterSizeKB; $sumUncInno += $t.AfterInnoDbKB
            $mUnc[$i, 0]  = $i + 1
            $mUnc[$i, 1]  = $t.TableName
            $mUnc[$i, 2]  = $t.PrimaryStatus
            $mUnc[$i, 3]  = $t.InferredRole
            $mUnc[$i, 4]  = if ($t.ContentHashStatus -eq '一致') { 'ハッシュ完全一致' } else { '全統計一致' }
            $mUnc[$i, 5]  = [double]$t.AfterRows
            $mUnc[$i, 6]  = [double]$t.AfterCols
            $mUnc[$i, 7]  = [double]$t.AfterSizeKB
            $mUnc[$i, 8]  = [double]$t.AfterInnoDbKB
            $mUnc[$i, 9]  = [double]$t.AfterAvgRowBytes
            $mUnc[$i, 10] = [double]$t.AfterAvgNullRate
            $mUnc[$i, 11] = if ([string]::IsNullOrEmpty($t.CandidatePks)) { '-' } else { $t.CandidatePks }
            $mUnc[$i, 12] = if ([string]::IsNullOrEmpty($t.DateColumns)) { '-' } else { $t.DateColumns }
            $mUnc[$i, 13] = $t.AfterWarning
        }
        $uncEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsUnchanged -StartRow 4 -StartCol 1 `
            -Headers @('No', 'テーブル名', '判定区分', '推定役割', '照合根拠', 'レコード数', '列数', 'CSVサイズ(KB)', '推定InnoDB(KB)', '平均行長(B)', '平均NULL率', 'PK/UK候補', '日時カラム', '既存警告・注意点') `
            -DataMatrix $mUnc -RowCount $uncItems.Count `
            -ColumnFormats @{ 1 = '#,##0'; 2 = '@'; 3 = '@'; 5 = '@'; 6 = '#,##0'; 7 = '#,##0'; 8 = '#,##0.00'; 9 = '#,##0.0'; 10 = '#,##0.0'; 11 = '0.0%'; 12 = '@'; 13 = '@'; 14 = '@' } `
            -CenterCols @(1, 3, 4, 5, 7) -RightCols @(6, 8, 9, 10, 11) -WrapCols @(12, 13, 14) -WarnCheckCols @(14)
        if ($uncItems.Count -gt 0) {
            Set-ComparisonRowStyles -Sheet $wsUnchanged -StartRow 5 -Items $uncItems -StatusCol 3 -WrapTextCols @(12, 13, 14) -DataMatrix $mUnc
            Write-ComparisonTotalRow -Sheet $wsUnchanged -Row ($uncEnd + 1) `
                -Values @('', '合計', ("計 {0} テーブル" -f $uncItems.Count), '', '', [double]$sumUncRows, '', [Math]::Round($sumUncSize, 2), [Math]::Round($sumUncInno, 1), '', '', '', '', '') `
                -ColumnFormats @{ 6 = '#,##0'; 8 = '#,##0.00'; 9 = '#,##0.0' } `
                -CenterCols @(3) -RightCols @(6, 8, 9)
        }
        Format-ExcelSheetCommon -Sheet $wsUnchanged -FreezeRow 4 -AutoFilterRange "A4:N$uncEnd" -Landscape $true

        # ======================================================================
        #  SHEET 6: 06_カラム・統計差分詳細
        # ======================================================================
        Write-CmpLog "シート作成中: 06_カラム・統計差分詳細..." 83
        Set-ColumnWidths -Sheet $wsDiffs -Widths @(22, 20, 16, 24, 24, 24, 28, 10)
        Set-SheetTitleBanner -Sheet $wsDiffs -Title 'カラム・分布・数値・日時・DDL・ハッシュ 統計差分詳細一覧 (変更検出の根拠データ)' -Subtitle $subInfo -EndCol 8

        $diffs = @($Comparison.AllColumnDiffs)
        $mDif = New-Object 'object[,]' ([Math]::Max(1, $diffs.Count)), 8
        for ($i = 0; $i -lt $diffs.Count; $i++) {
            $d = $diffs[$i]
            $mDif[$i, 0] = $d.TableName
            $mDif[$i, 1] = $d.ColumnName
            $mDif[$i, 2] = $d.Category
            $mDif[$i, 3] = $d.MetricName
            $mDif[$i, 4] = $d.BeforeValue
            $mDif[$i, 5] = $d.AfterValue
            $mDif[$i, 6] = $d.DeltaSummary
            $mDif[$i, 7] = $d.Severity
        }
        $difEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsDiffs -StartRow 4 -StartCol 1 `
            -Headers @('テーブル名', 'カラム名', '差分カテゴリ', '指標名', 'テスト開始前', 'テスト開始後', '差分・変化内容', '重要度') `
            -DataMatrix $mDif -RowCount $diffs.Count `
            -ColumnFormats @{ 1 = '@'; 2 = '@'; 3 = '@'; 4 = '@'; 5 = '@'; 6 = '@'; 7 = '@'; 8 = '@' } `
            -CenterCols @(3, 8) -WrapCols @(5, 6, 7)

        for ($i = 0; $i -lt $diffs.Count; $i++) {
            $rIdx = 5 + $i
            $sev = $diffs[$i].Severity
            if ($sev -eq '警告' -or $sev -eq '注意') {
                $cs = $wsDiffs.Cells.Item($rIdx, 8)
                $cs.Font.Color = $warnOle; $cs.Font.Bold = $true; $cs.Interior.Color = $warnBgOle
                if ($sev -eq '警告') {
                    $cc = $wsDiffs.Cells.Item($rIdx, 3)
                    $cc.Font.Color = $warnOle; $cc.Font.Bold = $true
                }
            }
            else {
                $wsDiffs.Cells.Item($rIdx, 8).Font.Color = $mutedOle
            }
        }
        Format-ExcelSheetCommon -Sheet $wsDiffs -FreezeRow 4 -AutoFilterRange "A4:H$difEnd" -Landscape $true

        # ======================================================================
        #  SHEET 7: 07_全カラム比較一覧 (Changed & Unchanged Column Matrix!)
        # ======================================================================
        Write-CmpLog "シート作成中: 07_全カラム比較一覧..." 88
        Set-ColumnWidths -Sheet $wsAllCols -Widths @(5, 20, 20, 18, 22, 12, 20, 22, 12, 18, 18, 22, 20, 20, 22, 14, 9, 36, 10)
        Set-SheetTitleBanner -Sheet $wsAllCols -Title '全カラム 変更あり・変更なし 比較マトリクス (オートフィルタで変更あり/変更なしを即座に抽出可能)' -Subtitle $subInfo -EndCol 19

        $colCmps = @($Comparison.AllColumnComparisons)
        $mColC = New-Object 'object[,]' ([Math]::Max(1, $colCmps.Count)), 19
        for ($i = 0; $i -lt $colCmps.Count; $i++) {
            $c = $colCmps[$i]
            $mColC[$i, 0]  = $i + 1
            $mColC[$i, 1]  = $c.TableName
            $mColC[$i, 2]  = $c.ColumnName
            $mColC[$i, 3]  = $c.TableStatus
            $mColC[$i, 4]  = $c.ColumnStatus
            $mColC[$i, 5]  = $c.ContentHashStatus
            $mColC[$i, 6]  = $c.TypeDisplay
            $mColC[$i, 7]  = $c.DdlKeyDisplay
            $mColC[$i, 8]  = [double]$c.NonNullDelta
            $mColC[$i, 9]  = $c.NullChangeDisplay
            $mColC[$i, 10] = $c.DistinctChangeDisplay
            $mColC[$i, 11] = $c.ByteLenDisplay
            $mColC[$i, 12] = $c.MinValDisplay
            $mColC[$i, 13] = $c.MaxValDisplay
            $mColC[$i, 14] = $c.Top1Display
            $mColC[$i, 15] = $c.NumericSumDeltaDisplay
            $mColC[$i, 16] = [double]$c.DiffCount
            $mColC[$i, 17] = $c.ChangeSummary
            $mColC[$i, 18] = $c.Severity
        }
        $colCmpEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsAllCols -StartRow 4 -StartCol 1 `
            -Headers @('No', 'テーブル名', 'カラム名', 'テーブル判定', 'カラム判定', '内容ハッシュ', '推定データ型 (前→後)', 'DDL型 / キー・役割', '有効件数増減', 'NULL件数 (前→後)', 'ユニーク数 (前→後)', 'バイト長 (平均/最大)', '最小値/最古 (前→後)', '最大値/最新 (前→後)', '最頻値Top1 (前→後)', '数値合計増減', '差分件数', '変更内容サマリー / 無変更根拠', '重要度') `
            -DataMatrix $mColC -RowCount $colCmps.Count `
            -ColumnFormats @{ 1 = '#,##0'; 2 = '@'; 3 = '@'; 4 = '@'; 5 = '@'; 6 = '@'; 7 = '@'; 8 = '@'; 9 = '+#,##0;-#,##0;0'; 10 = '@'; 11 = '@'; 12 = '@'; 13 = '@'; 14 = '@'; 15 = '@'; 16 = '@'; 17 = '#,##0'; 18 = '@'; 19 = '@' } `
            -CenterCols @(1, 4, 5, 6, 17, 19) -RightCols @(9, 16) -WrapCols @(7, 8, 12, 13, 14, 15, 18)

        for ($i = 0; $i -lt $colCmps.Count; $i++) {
            $rIdx = 5 + $i
            $c = $colCmps[$i]
            $cSt = $wsAllCols.Cells.Item($rIdx, 5)
            if ($c.ColumnStatus -match '削除') {
                $cSt.Font.Color = $warnOle; $cSt.Font.Bold = $true; $cSt.Interior.Color = $warnBgOle
            }
            elseif ($c.ColumnStatus -match '^追加') {
                $cSt.Font.Color = $navyOle; $cSt.Font.Bold = $true; $cSt.Interior.Color = $subBgOle
            }
            elseif ($c.IsChanged) {
                $cSt.Font.Color = $textOle; $cSt.Font.Bold = $true; $cSt.Interior.Color = $subBgOle
            }
            else {
                $cSt.Font.Color = $mutedOle
            }
            if ($c.ContentHashStatus -eq '変化') {
                $cH = $wsAllCols.Cells.Item($rIdx, 6); $cH.Font.Color = $navyOle; $cH.Font.Bold = $true
            }
            elseif ($c.ContentHashStatus -eq '一致') {
                $wsAllCols.Cells.Item($rIdx, 6).Font.Color = $mutedOle
            }
            $cNn = $wsAllCols.Cells.Item($rIdx, 9)
            if ($c.NonNullDelta -lt 0) { $cNn.Font.Color = $warnOle; $cNn.Font.Bold = $true }
            elseif ($c.NonNullDelta -gt 0) { $cNn.Font.Color = $navyOle; $cNn.Font.Bold = $true }
            else { $cNn.Font.Color = $mutedOle }

            $cSev = $wsAllCols.Cells.Item($rIdx, 19)
            if ($c.Severity -eq '警告' -or $c.Severity -eq '注意') {
                $cSev.Font.Color = $warnOle; $cSev.Font.Bold = $true; $cSev.Interior.Color = $warnBgOle
            }
            else {
                $cSev.Font.Color = $mutedOle
            }
        }
        if ($colCmps.Count -gt 0) {
            Write-ComparisonTotalRow -Sheet $wsAllCols -Row ($colCmpEnd + 1) `
                -Values @('', '合計', ("計 {0} カラム" -f $Comparison.TotalUniqueColumnCount), '', ("変更 {0} 列 / 不変 {1} 列" -f $Comparison.ChangedColumnTotal, $Comparison.UnchangedColumnTotal), '', '', ("追加 {0} / 削除 {1} / DDL変更 {2}" -f $Comparison.AddedColumnTotal, $Comparison.RemovedColumnTotal, $Comparison.SchemaChangedColumnTotal), '', '', '', '', '', '', '', '', [double]$Comparison.AllColumnDiffs.Count, '全テーブルの全カラム照合集計', '') `
                -ColumnFormats @{ 17 = '#,##0' } -CenterCols @(5, 17) -RightCols @()
        }
        Format-ExcelSheetCommon -Sheet $wsAllCols -FreezeRow 4 -AutoFilterRange "A4:S$colCmpEnd" -Landscape $true

        # ======================================================================
        #  SHEET 8: 08_性能・DB・時系列・品質変化 (Test-Significant Deep Diagnostics)
        # ======================================================================
        Write-CmpLog "シート作成中: 08_性能・DB・時系列・品質変化..." 93
        Set-ColumnWidths -Sheet $wsPerfDb -Widths @(16, 24, 16, 16, 15, 15, 16, 18, 22, 34, 24)
        Set-SheetTitleBanner -Sheet $wsPerfDb -Title 'テスト有意情報サマリー (性能警告推移・FK整合性・時系列増分・Aurora/MySQL負荷＆物理容量・JMeter適性)' -Subtitle $subInfo -EndCol 11

        $pRow = 4
        # Section 1: Warnings lifecycle
        Set-SectionHeader -Sheet $wsPerfDb -Row $pRow -StartCol 1 -EndCol 11 -Text ("1. 性能リスク・警告の推移 (新規発生 {0} / 悪化 {1} / 解消 {2} / 継続 {3})" -f $Comparison.NewWarningCount, $Comparison.WorsenedWarningCount, $Comparison.ResolvedWarningCount, $Comparison.OngoingWarningCount)
        $pRow++
        $wCmps = @($Comparison.WarningComparisons)
        $mW = New-Object 'object[,]' ([Math]::Max(1, $wCmps.Count)), 8
        for ($i = 0; $i -lt $wCmps.Count; $i++) {
            $w = $wCmps[$i]
            $mW[$i, 0] = $w.Status
            $mW[$i, 1] = $w.Severity
            $mW[$i, 2] = $w.Category
            $mW[$i, 3] = $w.TableName
            $mW[$i, 4] = $w.ColumnName
            $mW[$i, 5] = $w.BeforeMetric
            $mW[$i, 6] = $w.AfterMetric
            $mW[$i, 7] = if ([string]::IsNullOrEmpty($w.PerfImpact)) { $w.Recommendation } else { ("【影響】{0} / 【対策】{1}" -f $w.PerfImpact, $w.Recommendation) }
        }
        $wEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsPerfDb -StartRow $pRow -StartCol 1 `
            -Headers @('変化状態', '重要度', 'カテゴリ', 'テーブル名', 'カラム名', '開始前指標', '開始後指標', '影響・推奨アクション') `
            -DataMatrix $mW -RowCount $wCmps.Count `
            -ColumnFormats @{ 1 = '@'; 2 = '@'; 3 = '@'; 4 = '@'; 5 = '@'; 6 = '@'; 7 = '@'; 8 = '@' } `
            -CenterCols @(1, 2, 3) -WrapCols @(6, 7, 8)
        for ($i = 0; $i -lt $wCmps.Count; $i++) {
            $rr = $pRow + 1 + $i
            $wc = $wCmps[$i]
            $c1 = $wsPerfDb.Cells.Item($rr, 1)
            if ($wc.IsAlert) { $c1.Font.Color = $warnOle; $c1.Font.Bold = $true; $c1.Interior.Color = $warnBgOle }
            elseif ($wc.Status -eq '解消' -or $wc.Status -eq '改善') { $c1.Font.Color = $navyOle; $c1.Font.Bold = $true; $c1.Interior.Color = $subBgOle }
            else { $c1.Font.Color = $mutedOle }
            if ($wc.Severity -eq '高' -and $wc.Status -ne '解消') {
                $c2 = $wsPerfDb.Cells.Item($rr, 2); $c2.Font.Color = $warnOle; $c2.Font.Bold = $true; $c2.Interior.Color = $warnBgOle
            }
        }
        $pRow = $wEnd + 2

        # Section 2: Relationship & FK Integrity
        Set-SectionHeader -Sheet $wsPerfDb -Row $pRow -StartCol 1 -EndCol 11 -Text ("2. テーブル間リレーション・FK参照整合性・1:Nファンアウト変化 (孤立行増減: {0}{1:#,##0}行)" -f $(if ($Comparison.OrphanRowDelta -gt 0) { '+' } else { '' }), $Comparison.OrphanRowDelta)
        $pRow++
        $rCmps = @($Comparison.RelationshipComparisons)
        $mR = New-Object 'object[,]' ([Math]::Max(1, $rCmps.Count)), 10
        for ($i = 0; $i -lt $rCmps.Count; $i++) {
            $r = $rCmps[$i]
            $mR[$i, 0] = $r.Status
            $mR[$i, 1] = $r.RelationKey
            $mR[$i, 2] = [double]$r.BeforeMatchRate
            $mR[$i, 3] = [double]$r.AfterMatchRate
            $mR[$i, 4] = [double]$r.BeforeOrphanRows
            $mR[$i, 5] = [double]$r.AfterOrphanRows
            $mR[$i, 6] = [double]$r.OrphanDelta
            $mR[$i, 7] = ("{0:#,##0} → {1:#,##0}" -f $r.BeforeFanOutMax, $r.AfterFanOutMax)
            $mR[$i, 8] = ("{0} → {1}" -f $r.BeforeRisk, $r.AfterRisk)
            $mR[$i, 9] = $r.Note
        }
        $rEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsPerfDb -StartRow $pRow -StartCol 1 `
            -Headers @('変化状態', 'リレーション (親.PK → 子.FK)', '開始前一致率', '開始後一致率', '開始前孤立行', '開始後孤立行', '孤立行増減', '最大ファンアウト(前→後)', 'リスク変化', '所見') `
            -DataMatrix $mR -RowCount $rCmps.Count `
            -ColumnFormats @{ 1 = '@'; 2 = '@'; 3 = '0.0%'; 4 = '0.0%'; 5 = '#,##0'; 6 = '#,##0'; 7 = '+#,##0;-#,##0;0'; 8 = '@'; 9 = '@'; 10 = '@' } `
            -CenterCols @(1, 8, 9) -RightCols @(3, 4, 5, 6, 7) -WrapCols @(2, 9, 10)
        for ($i = 0; $i -lt $rCmps.Count; $i++) {
            $rr = $pRow + 1 + $i
            $rc = $rCmps[$i]
            if ($rc.HasAlert) {
                $c1 = $wsPerfDb.Cells.Item($rr, 1); $c1.Font.Color = $warnOle; $c1.Font.Bold = $true; $c1.Interior.Color = $warnBgOle
            }
            if ($rc.OrphanDelta -gt 0) {
                $c7 = $wsPerfDb.Cells.Item($rr, 7); $c7.Font.Color = $warnOle; $c7.Font.Bold = $true; $c7.Interior.Color = $warnBgOle
            }
            elseif ($rc.OrphanDelta -lt 0) {
                $c7 = $wsPerfDb.Cells.Item($rr, 7); $c7.Font.Color = $navyOle; $c7.Font.Bold = $true
            }
        }
        $pRow = $rEnd + 2

        # Section 3: Time-Series Activity Deltas
        Set-SectionHeader -Sheet $wsPerfDb -Row $pRow -StartCol 1 -EndCol 11 -Text '3. テスト期間中の時系列アクティビティ増分 (月別・日別・時間帯別の件数推移)'
        $pRow++
        $tsCmps = @($Comparison.TimeSeriesDeltas | Select-Object -First 80)
        $mTs = New-Object 'object[,]' ([Math]::Max(1, $tsCmps.Count)), 7
        for ($i = 0; $i -lt $tsCmps.Count; $i++) {
            $ts = $tsCmps[$i]
            $mTs[$i, 0] = $ts.Granularity
            $mTs[$i, 1] = $ts.TableColumn
            $mTs[$i, 2] = $ts.PeriodLabel
            $mTs[$i, 3] = [double]$ts.BeforeCount
            $mTs[$i, 4] = [double]$ts.AfterCount
            $mTs[$i, 5] = [double]$ts.DeltaCount
            $mTs[$i, 6] = if ($ts.HasGrowthRate) { [double]$ts.GrowthRate } else { '新規' }
        }
        $tsEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsPerfDb -StartRow $pRow -StartCol 1 `
            -Headers @('粒度', '対象テーブル.日時カラム', '期間・時間帯', '開始前件数', '開始後件数', '増減件数', '増減率') `
            -DataMatrix $mTs -RowCount $tsCmps.Count `
            -ColumnFormats @{ 1 = '@'; 2 = '@'; 3 = '@'; 4 = '#,##0'; 5 = '#,##0'; 6 = '+#,##0;-#,##0;0'; 7 = '+0.0%;-0.0%;0.0%' } `
            -CenterCols @(1, 3) -RightCols @(4, 5, 6, 7)
        for ($i = 0; $i -lt $tsCmps.Count; $i++) {
            $rr = $pRow + 1 + $i
            $c6 = $wsPerfDb.Cells.Item($rr, 6)
            if ($tsCmps[$i].DeltaCount -gt 0) { $c6.Font.Color = $navyOle; $c6.Font.Bold = $true }
            elseif ($tsCmps[$i].DeltaCount -lt 0) { $c6.Font.Color = $warnOle; $c6.Font.Bold = $true }
        }
        $pRow = $tsEnd + 2

        # Section 4A: DirectDB Physical InnoDB Sizes & AUTO_INCREMENT
        Set-SectionHeader -Sheet $wsPerfDb -Row $pRow -StartCol 1 -EndCol 11 -Text '4A. Aurora / MySQL 実物理容量 (Data/Index/断片化) ＆ AUTO_INCREMENT 推移 (DB直接分析時)'
        $pRow++
        $physTbls = @($Comparison.AllTables | Where-Object { $_.HasDbPhysicalMeta })
        $mPh = New-Object 'object[,]' ([Math]::Max(1, $physTbls.Count)), 11
        for ($i = 0; $i -lt $physTbls.Count; $i++) {
            $t = $physTbls[$i]
            $aiDeltaTxt = if (-not $t.HasAutoIncrement) { '-' } elseif ($t.AutoIncrementGap -gt 0) { ("+{0:#,##0} (ギャップ {1:#,##0}件)" -f $t.AutoIncrementDelta, $t.AutoIncrementGap) } else { ("{0}{1:#,##0}" -f $(if ($t.AutoIncrementDelta -gt 0) { '+' } else { '' }), $t.AutoIncrementDelta) }
            $mPh[$i, 0]  = $t.TableName
            $mPh[$i, 1]  = [double]$t.BeforeDataLengthKB
            $mPh[$i, 2]  = [double]$t.AfterDataLengthKB
            $mPh[$i, 3]  = [double]$t.DataLengthDeltaKB
            $mPh[$i, 4]  = [double]$t.BeforeIndexLengthKB
            $mPh[$i, 5]  = [double]$t.AfterIndexLengthKB
            $mPh[$i, 6]  = [double]$t.IndexLengthDeltaKB
            $mPh[$i, 7]  = [double]$t.DataFreeDeltaKB
            $mPh[$i, 8]  = if ($t.HasAutoIncrement) { [double]$t.BeforeAutoIncrement } else { '-' }
            $mPh[$i, 9]  = if ($t.HasAutoIncrement) { [double]$t.AfterAutoIncrement } else { '-' }
            $mPh[$i, 10] = $aiDeltaTxt
        }
        $phEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsPerfDb -StartRow $pRow -StartCol 1 `
            -Headers @('テーブル名', '実Data前(KB)', '実Data後(KB)', 'Data増減(KB)', '実Index前(KB)', '実Index後(KB)', 'Index増減(KB)', '断片化空き増減(KB)', 'AI開始前', 'AI開始後', 'AUTO_INCREMENT増減/ギャップ') `
            -DataMatrix $mPh -RowCount $physTbls.Count `
            -ColumnFormats @{ 1 = '@'; 2 = '#,##0.0'; 3 = '#,##0.0'; 4 = '+#,##0.0;-#,##0.0;0.0'; 5 = '#,##0.0'; 6 = '#,##0.0'; 7 = '+#,##0.0;-#,##0.0;0.0'; 8 = '+#,##0.0;-#,##0.0;0.0'; 9 = '#,##0'; 10 = '#,##0'; 11 = '@' } `
            -CenterCols @(11) -RightCols @(2, 3, 4, 5, 6, 7, 8, 9, 10) -WrapCols @(11)
        for ($i = 0; $i -lt $physTbls.Count; $i++) {
            $rr = $pRow + 1 + $i
            if ($physTbls[$i].AutoIncrementGap -gt 0) {
                $c11 = $wsPerfDb.Cells.Item($rr, 11); $c11.Font.Color = $warnOle; $c11.Font.Bold = $true; $c11.Interior.Color = $warnBgOle
            }
        }
        $pRow = $phEnd + 2

        # Section 4B: MySQL Server Workload (SHOW GLOBAL STATUS Before vs After)
        Set-SectionHeader -Sheet $wsPerfDb -Row $pRow -StartCol 1 -EndCol 11 -Text '4B. Aurora / MySQL サーバー負荷統計差分 (SHOW GLOBAL STATUS テスト前後比較)'
        $pRow++
        $srvItems = @($Comparison.ServerMetricComparisons)
        $mSrv = New-Object 'object[,]' ([Math]::Max(1, $srvItems.Count)), 8
        for ($i = 0; $i -lt $srvItems.Count; $i++) {
            $sm = $srvItems[$i]
            $mSrv[$i, 0] = $sm.Category
            $mSrv[$i, 1] = $sm.MetricName
            $mSrv[$i, 2] = $sm.DisplayBefore
            $mSrv[$i, 3] = $sm.DisplayAfter
            $mSrv[$i, 4] = $sm.DisplayDelta
            $mSrv[$i, 5] = $sm.RatePerSecDisplay
            $mSrv[$i, 6] = $sm.Severity
            $mSrv[$i, 7] = $sm.Note
        }
        $srvEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsPerfDb -StartRow $pRow -StartCol 1 `
            -Headers @('カテゴリ', 'サーバー指標', '開始前', '開始後', 'テスト期間増分', '平均レート(/秒)', '判定', '性能試験上の意味・所見') `
            -DataMatrix $mSrv -RowCount $srvItems.Count `
            -ColumnFormats @{ 1 = '@'; 2 = '@'; 3 = '@'; 4 = '@'; 5 = '@'; 6 = '@'; 7 = '@'; 8 = '@' } `
            -CenterCols @(1, 7) -RightCols @(3, 4, 5, 6) -WrapCols @(8)
        for ($i = 0; $i -lt $srvItems.Count; $i++) {
            $rr = $pRow + 1 + $i
            if ($srvItems[$i].Severity -eq '警告' -or $srvItems[$i].Severity -eq '注意') {
                $c7 = $wsPerfDb.Cells.Item($rr, 7); $c7.Font.Color = $warnOle; $c7.Font.Bold = $true; $c7.Interior.Color = $warnBgOle
            }
        }
        $pRow = $srvEnd + 2

        # Section 4C: DDL / Index / Declared FK Schema Diffs
        Set-SectionHeader -Sheet $wsPerfDb -Row $pRow -StartCol 1 -EndCol 11 -Text '4C. DDLスキーマ・インデックス・外部キー制約 変更一覧 (DB直接分析時)'
        $pRow++
        $schItems = @($Comparison.DbSchemaDiffs)
        $mSch = New-Object 'object[,]' ([Math]::Max(1, $schItems.Count)), 8
        for ($i = 0; $i -lt $schItems.Count; $i++) {
            $sd = $schItems[$i]
            $mSch[$i, 0] = $sd.TableName
            $mSch[$i, 1] = $sd.ObjectType
            $mSch[$i, 2] = $sd.ObjectName
            $mSch[$i, 3] = $sd.ChangeType
            $mSch[$i, 4] = $sd.BeforeDef
            $mSch[$i, 5] = $sd.AfterDef
            $mSch[$i, 6] = $sd.Severity
            $mSch[$i, 7] = $sd.Note
        }
        $schEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsPerfDb -StartRow $pRow -StartCol 1 `
            -Headers @('テーブル名', '種別', '対象名', '変更区分', '開始前定義', '開始後定義', '重要度', '備考') `
            -DataMatrix $mSch -RowCount $schItems.Count `
            -ColumnFormats @{ 1 = '@'; 2 = '@'; 3 = '@'; 4 = '@'; 5 = '@'; 6 = '@'; 7 = '@'; 8 = '@' } `
            -CenterCols @(2, 4, 7) -WrapCols @(5, 6, 8)
        for ($i = 0; $i -lt $schItems.Count; $i++) {
            $rr = $pRow + 1 + $i
            if ($schItems[$i].Severity -eq '警告' -or $schItems[$i].Severity -eq '注意') {
                $c7 = $wsPerfDb.Cells.Item($rr, 7); $c7.Font.Color = $warnOle; $c7.Font.Bold = $true; $c7.Interior.Color = $warnBgOle
            }
        }
        $pRow = $schEnd + 2

        # Section 5: JMeter Readiness Changes
        Set-SectionHeader -Sheet $wsPerfDb -Row $pRow -StartCol 1 -EndCol 11 -Text '5. JMeter 並列適性・一意キー候補の変化'
        $pRow++
        $jmItems = @($Comparison.JMeterComparisons)
        $mJmC = New-Object 'object[,]' ([Math]::Max(1, $jmItems.Count)), 8
        for ($i = 0; $i -lt $jmItems.Count; $i++) {
            $jm = $jmItems[$i]
            $mJmC[$i, 0] = $jm.Status
            $mJmC[$i, 1] = $jm.TableName
            $mJmC[$i, 2] = $jm.KeyColumn
            $mJmC[$i, 3] = $jm.DataType
            $mJmC[$i, 4] = [double]$jm.BeforeUnique
            $mJmC[$i, 5] = [double]$jm.AfterUnique
            $mJmC[$i, 6] = [double]$jm.UniqueDelta
            $mJmC[$i, 7] = ("{0} → {1}" -f $jm.BeforeSuitability, $jm.AfterSuitability)
        }
        $jmEnd = Write-StyledTable -CheckCancellation $checkReport -Sheet $wsPerfDb -StartRow $pRow -StartCol 1 `
            -Headers @('変化状態', 'テーブル名', 'キーカラム', 'データ型', '開始前キー数', '開始後キー数', 'キー数増減', 'JMeter並列適性 (前→後)') `
            -DataMatrix $mJmC -RowCount $jmItems.Count `
            -ColumnFormats @{ 1 = '@'; 2 = '@'; 3 = '@'; 4 = '@'; 5 = '#,##0'; 6 = '#,##0'; 7 = '+#,##0;-#,##0;0'; 8 = '@' } `
            -CenterCols @(1, 4) -RightCols @(5, 6, 7) -WrapCols @(8)
        for ($i = 0; $i -lt $jmItems.Count; $i++) {
            $rr = $pRow + 1 + $i
            if ($jmItems[$i].HasAlert) {
                $c1 = $wsPerfDb.Cells.Item($rr, 1); $c1.Font.Color = $warnOle; $c1.Font.Bold = $true; $c1.Interior.Color = $warnBgOle
            }
            if ($jmItems[$i].UniqueDelta -lt 0) {
                $c7 = $wsPerfDb.Cells.Item($rr, 7); $c7.Font.Color = $warnOle; $c7.Font.Bold = $true
            }
            elseif ($jmItems[$i].UniqueDelta -gt 0) {
                $c7 = $wsPerfDb.Cells.Item($rr, 7); $c7.Font.Color = $navyOle; $c7.Font.Bold = $true
            }
        }
        Format-ExcelSheetCommon -Sheet $wsPerfDb -FreezeRow 0 -Landscape $true

        $wsSummary.Activate()
        $excel.Calculation = -4105 # xlCalculationAutomatic
        $excel.ScreenUpdating = $true

        Write-CmpLog "比較Excelファイル保存中: $OutputPath" 98
        $wb.SaveAs($stagingPath, 51)
        $wb.Close($false)
        $excel.Quit()
        if ($null -ne $PumpCallback) { & $PumpCallback }
        if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }
        [AuroraAnalyzer.AtomicFile]::Publish($stagingPath, $OutputPath)
        Write-CmpLog "テスト前後比較Excelレポート出力完了: $OutputPath" 100
    }
    finally {
        foreach ($s in $sheetsToRelease) {
            if ($null -ne $s) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($s) } catch { } }
        }
        if ($null -ne $wb) {
            try { $wb.Close($false) } catch { }
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } catch { }
        }
        if ($null -ne $excel) {
            try { $excel.Quit() } catch { }
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch { }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()

        if ($excelPid -gt 0) {
            try {
                $ep = [System.Diagnostics.Process]::GetProcessById($excelPid)
                if ($null -ne $ep -and -not $ep.HasExited) {
                    if (-not $ep.WaitForExit(1500)) {
                        $ep.Kill()
                        [void]$ep.WaitForExit(2000)
                    }
                }
            }
            catch { }
        }
        try { if ([IO.File]::Exists($stagingPath)) { [IO.File]::Delete($stagingPath) } } catch { }
    }
}

# ==============================================================================
#  Module 4D: Pure .NET OpenXML Comparison Report Fallback Generator
# ==============================================================================
function Export-ComparisonToOpenXml {
    param(
        [Parameter(Mandatory = $true)]$Comparison,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [scriptblock]$LogCallback = $null,
        [AuroraAnalyzer.OperationCancelState]$CancelState = $null,
        [scriptblock]$PumpCallback = $null
    )

    if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    function XmlEscCmp([string]$s) {
        return [AuroraAnalyzer.XmlText]::Escape($s)
    }

    function ColLetterCmp([int]$colIndex1Based) {
        $dividend = $colIndex1Based
        $colName = ''
        while ($dividend -gt 0) {
            $modulo = ($dividend - 1) % 26
            $colName = [char](65 + $modulo) + $colName
            $dividend = [int][Math]::Floor(($dividend - $modulo) / 26)
        }
        return $colName
    }

    function Build-CmpSheetXml {
        param(
            [string]$Title,
            [string]$Subtitle,
            [string[]]$Headers,
            [object[]]$Rows,
            [double[]]$ColWidths,
            [hashtable]$NumColStyles = @{},
            [int]$StatusCol = 0,
            [int]$AlertCol = 0,
            [int[]]$SignedDeltaCols = @(),
            [int[]]$CenterCols = @(),
            [int[]]$WrapCols = @(),
            [object[]]$TotalRowValues = $null,
            [hashtable]$TotalColStyles = @{},
            [bool]$IncludeChartDrawing = $false
        )

        # Style ID mapping (strictly neutral monotone backgrounds; paired White / #F7F7F7 for clean zebra rows):
        # 0: Default | 1: Banner Title (#FFF on #1A2B4C) | 2: Subtitle (#333 on #F2F2F2 + Navy bottom) | 3: Table Header (#FFF bold on #1A2B4C)
        # 4/5: Text Left (White / #F7F7F7)
        # 6/7: Int #,##0 (White / #F7F7F7)
        # 8/9: Dec2 #,##0.00 (White / #F7F7F7)
        # 10/11: Pct 0.0% (White / #F7F7F7)
        # 12/13: Dec1 #,##0.0 (White / #F7F7F7)
        # 14/15: Signed Int Navy bold (White / #F7F7F7)
        # 16/17: Signed Dec2 Navy bold (White / #F7F7F7)
        # 18/19: Signed Dec1 Navy bold (White / #F7F7F7)
        # 20/21: Signed Pct DarkGray (White / #F7F7F7)
        # 22/23: Text Wrap Left (White / #F7F7F7)
        # 24/25: Text Center (White / #F7F7F7)
        # 26: Alert / Warning text center wrap (#800020 bold on #F5F5F5)
        # 27: Status Added (#1A2B4C bold on #F2F2F2)
        # 28: Status Updated (#333333 bold on #F2F2F2)
        # 29/30: Muted Center (#595959 on White / #F7F7F7)
        # 31: Burgundy Signed Int (#800020 bold on #F5F5F5)
        # 32: Burgundy Signed Dec2 (#800020 bold on #F5F5F5)
        # 33: Burgundy Signed Dec1 (#800020 bold on #F5F5F5)
        # 34: Total Text (#1A2B4C bold on #EAEAEA)
        # 35: Total Int | 36: Total Signed Int | 37: Total Dec2 | 38: Total Signed Dec2 | 39: Total Dec1 | 40: Total Signed Dec1 | 41: Total Pct
        # 42: Total Burgundy Signed Int | 43: Total Burgundy Signed Dec2 | 44: Total Burgundy Signed Dec1 | 45: Total Burgundy Pct
        # 46: Alert text left wrap (#800020 bold on #F5F5F5)

        $sb = New-Object System.Text.StringBuilder 32768
        [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
        [void]$sb.Append('<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>')
        [void]$sb.Append('<sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="4" topLeftCell="A5" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
        [void]$sb.Append('<sheetFormatPr defaultRowHeight="22"/>')
        [void]$sb.Append('<cols>')
        for ($i = 0; $i -lt $ColWidths.Count; $i++) {
            [void]$sb.AppendFormat([System.Globalization.CultureInfo]::InvariantCulture, '<col min="{0}" max="{0}" width="{1}" customWidth="1"/>', ($i + 1), $ColWidths[$i])
        }
        [void]$sb.Append('</cols><sheetData>')

        $endL = ColLetterCmp $Headers.Count
        [void]$sb.Append('<row r="1" ht="32" customHeight="1">')
        [void]$sb.AppendFormat('<c r="A1" t="inlineStr" s="1"><is><t>{0}</t></is></c>', (XmlEscCmp $Title))
        for ($c = 1; $c -lt $Headers.Count; $c++) { [void]$sb.AppendFormat('<c r="{0}1" s="1"/>', (ColLetterCmp ($c + 1))) }
        [void]$sb.Append('</row>')

        [void]$sb.Append('<row r="2" ht="22" customHeight="1">')
        [void]$sb.AppendFormat('<c r="A2" t="inlineStr" s="2"><is><t>{0}</t></is></c>', (XmlEscCmp $Subtitle))
        for ($c = 1; $c -lt $Headers.Count; $c++) { [void]$sb.AppendFormat('<c r="{0}2" s="2"/>', (ColLetterCmp ($c + 1))) }
        [void]$sb.Append('</row>')
        [void]$sb.Append('<row r="3" ht="10" customHeight="1"/>')

        [void]$sb.Append('<row r="4" ht="28" customHeight="1">')
        for ($c = 0; $c -lt $Headers.Count; $c++) {
            $ref = (ColLetterCmp ($c + 1)) + '4'
            [void]$sb.AppendFormat('<c r="{0}" t="inlineStr" s="3"><is><t>{1}</t></is></c>', $ref, (XmlEscCmp $Headers[$c]))
        }
        [void]$sb.Append('</row>')

        $rNum = 5
        $merges = New-Object System.Collections.Generic.List[string]
        $merges.Add("A1:${endL}1")
        $merges.Add("A2:${endL}2")

        if ($Rows.Count -eq 0) {
            [void]$sb.Append('<row r="5" ht="24" customHeight="1">')
            [void]$sb.Append('<c r="A5" t="inlineStr" s="29"><is><t>該当データなし</t></is></c>')
            for ($c = 1; $c -lt $Headers.Count; $c++) { [void]$sb.AppendFormat('<c r="{0}5" s="29"/>', (ColLetterCmp ($c + 1))) }
            [void]$sb.Append('</row>')
            $merges.Add("A5:${endL}5")
            $rNum = 6
        }
        else {
            for ($r = 0; $r -lt $Rows.Count; $r++) {
                if (($r % 32) -eq 0) {
                    if ($null -ne $PumpCallback) { & $PumpCallback }
                    if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }
                }
                $rowVals = $Rows[$r]
                $isZebra = (($r % 2) -eq 1)
                $ht = 22.0
                if ($WrapCols.Count -gt 0) {
                    $wTexts = New-Object System.Collections.Generic.List[string]
                    $wWidths = New-Object System.Collections.Generic.List[double]
                    foreach ($wc in $WrapCols) {
                        $wTexts.Add([string]$rowVals[$wc - 1])
                        $wWidths.Add($(if (($wc - 1) -lt $ColWidths.Count) { [double]$ColWidths[$wc - 1] } else { 26.0 }))
                    }
                    $ht = Get-EstimatedRowHeight -Texts $wTexts.ToArray() -ColWidths $wWidths.ToArray() -MinHeight 22.0 -MaxHeight 110.0
                }
                [void]$sb.AppendFormat([System.Globalization.CultureInfo]::InvariantCulture, '<row r="{0}" ht="{1:0.#}" customHeight="1">', $rNum, $ht)

                for ($c = 0; $c -lt $Headers.Count; $c++) {
                    $ref = (ColLetterCmp ($c + 1)) + $rNum
                    $v = $rowVals[$c]
                    $col1 = $c + 1

                    if ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal]) {
                        $dv = [double]$v
                        $baseSt = if ($NumColStyles.ContainsKey($col1)) { [int]$NumColStyles[$col1] } else { 4 }
                        $st = if ($isZebra) { $baseSt + 1 } else { $baseSt }
                        if ($SignedDeltaCols -contains $col1) {
                            if ($dv -lt -1e-9) {
                                if ($baseSt -eq 14) { $st = 31 }
                                elseif ($baseSt -eq 16) { $st = 32 }
                                elseif ($baseSt -eq 18) { $st = 33 }
                            }
                            elseif ([Math]::Abs($dv) -le 1e-9) {
                                if ($baseSt -eq 14) { $st = if ($isZebra) { 7 } else { 6 } }
                                elseif ($baseSt -eq 16) { $st = if ($isZebra) { 9 } else { 8 } }
                                elseif ($baseSt -eq 18) { $st = if ($isZebra) { 13 } else { 12 } }
                            }
                        }
                        [void]$sb.AppendFormat([System.Globalization.CultureInfo]::InvariantCulture, '<c r="{0}" s="{1}"><v>{2}</v></c>', $ref, $st, $v)
                    }
                    else {
                        $sv = [string]$v
                        $st = if ($WrapCols -contains $col1) { if ($isZebra) { 23 } else { 22 } } elseif ($CenterCols -contains $col1) { if ($isZebra) { 25 } else { 24 } } else { if ($isZebra) { 5 } else { 4 } }
                        if ($col1 -eq $StatusCol) {
                            $isZeroSummaryBucket = ($StatusCol -eq 1 -and $rowVals.Count -gt 1 -and ($rowVals[1] -is [double]) -and [double]$rowVals[1] -eq 0.0)
                            if ($isZeroSummaryBucket) { $st = if ($isZebra) { 30 } else { 29 } }
                            elseif ($sv -match '(削除|減少)') { $st = 26 }
                            elseif ($sv -match '(追加|新規)') { $st = 27 }
                            elseif ($sv -match '更新|変更あり') { $st = 28 }
                            elseif ($sv -match '変更なし') { $st = if ($isZebra) { 30 } else { 29 } }
                        }
                        elseif ($col1 -eq $AlertCol -and (Test-WarningCellText -Text $sv)) {
                            $st = if ($CenterCols -contains $col1) { 26 } else { 46 }
                        }
                        elseif ($col1 -eq $AlertCol -and $sv -eq '通常') {
                            $st = if ($isZebra) { 30 } else { 29 }
                        }
                        [void]$sb.AppendFormat('<c r="{0}" t="inlineStr" s="{1}"><is><t>{2}</t></is></c>', $ref, $st, (XmlEscCmp $sv))
                    }
                }
                [void]$sb.Append('</row>')
                $rNum++
            }
        }

        $filterLast = [Math]::Max(4, $rNum - 1)
        if ($null -ne $TotalRowValues -and $Rows.Count -gt 0) {
            [void]$sb.AppendFormat('<row r="{0}" ht="24" customHeight="1">', $rNum)
            for ($c = 0; $c -lt $Headers.Count; $c++) {
                $ref = (ColLetterCmp ($c + 1)) + $rNum
                $v = $TotalRowValues[$c]
                $col1 = $c + 1
                if ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal]) {
                    $st = if ($TotalColStyles.ContainsKey($col1)) { [int]$TotalColStyles[$col1] } else { 35 }
                    if ([double]$v -lt -1e-9) {
                        if ($st -eq 36) { $st = 42 }
                        elseif ($st -eq 38) { $st = 43 }
                        elseif ($st -eq 40) { $st = 44 }
                        elseif ($st -eq 41) { $st = 45 }
                        elseif ($st -eq 47) { $st = 48 }
                    }
                    [void]$sb.AppendFormat([System.Globalization.CultureInfo]::InvariantCulture, '<c r="{0}" s="{1}"><v>{2}</v></c>', $ref, $st, $v)
                }
                else {
                    [void]$sb.AppendFormat('<c r="{0}" t="inlineStr" s="34"><is><t>{1}</t></is></c>', $ref, (XmlEscCmp ([string]$v)))
                }
            }
            [void]$sb.Append('</row>')
            $rNum++
        }

        [void]$sb.Append('</sheetData>')
        [void]$sb.AppendFormat('<autoFilter ref="A4:{0}{1}"/>', $endL, $filterLast)
        [void]$sb.AppendFormat('<mergeCells count="{0}">', $merges.Count)
        foreach ($mg in $merges) { [void]$sb.AppendFormat('<mergeCell ref="{0}"/>', $mg) }
        [void]$sb.Append('</mergeCells>')
        [void]$sb.Append('<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>')
        [void]$sb.Append('<pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/>')
        if ($IncludeChartDrawing) { [void]$sb.Append('<drawing r:id="rId1"/>') }
        [void]$sb.Append('</worksheet>')
        return $sb.ToString()
    }

    function Build-CmpMultiSectionSheetXml {
        param(
            [string]$Title,
            [string]$Subtitle,
            [double[]]$ColWidths,
            [object[]]$Sections
        )
        $endL = ColLetterCmp $ColWidths.Count
        $sb = New-Object System.Text.StringBuilder 32768
        $merges = New-Object System.Collections.Generic.List[string]
        [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
        [void]$sb.Append('<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>')
        [void]$sb.Append('<sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>')
        [void]$sb.Append('<sheetFormatPr defaultRowHeight="22"/>')
        [void]$sb.Append('<cols>')
        for ($i = 0; $i -lt $ColWidths.Count; $i++) {
            [void]$sb.AppendFormat([System.Globalization.CultureInfo]::InvariantCulture, '<col min="{0}" max="{0}" width="{1}" customWidth="1"/>', ($i + 1), $ColWidths[$i])
        }
        [void]$sb.Append('</cols><sheetData>')

        [void]$sb.Append('<row r="1" ht="32" customHeight="1">')
        [void]$sb.AppendFormat('<c r="A1" t="inlineStr" s="1"><is><t>{0}</t></is></c>', (XmlEscCmp $Title))
        for ($c = 1; $c -lt $ColWidths.Count; $c++) { [void]$sb.AppendFormat('<c r="{0}1" s="1"/>', (ColLetterCmp ($c + 1))) }
        [void]$sb.Append('</row>')
        $merges.Add("A1:${endL}1")

        [void]$sb.Append('<row r="2" ht="22" customHeight="1">')
        [void]$sb.AppendFormat('<c r="A2" t="inlineStr" s="2"><is><t>{0}</t></is></c>', (XmlEscCmp $Subtitle))
        for ($c = 1; $c -lt $ColWidths.Count; $c++) { [void]$sb.AppendFormat('<c r="{0}2" s="2"/>', (ColLetterCmp ($c + 1))) }
        [void]$sb.Append('</row>')
        $merges.Add("A2:${endL}2")
        [void]$sb.Append('<row r="3" ht="10" customHeight="1"/>')

        $rNum = 4
        foreach ($sec in $Sections) {
            # Section banner row
            [void]$sb.AppendFormat('<row r="{0}" ht="25" customHeight="1">', $rNum)
            [void]$sb.AppendFormat('<c r="A{0}" t="inlineStr" s="49"><is><t>{1}</t></is></c>', $rNum, (XmlEscCmp $sec.Title))
            for ($c = 1; $c -lt $ColWidths.Count; $c++) { [void]$sb.AppendFormat('<c r="{0}{1}" s="49"/>', (ColLetterCmp ($c + 1)), $rNum) }
            [void]$sb.Append('</row>')
            $merges.Add("A${rNum}:${endL}${rNum}")
            $rNum++

            $hdrs = @($sec.Headers)
            $rows = @($sec.Rows)
            $numSt = if ($sec.ContainsKey('NumStyles')) { $sec.NumStyles } else { @{} }
            $cCols = @(if ($sec.ContainsKey('CenterCols')) { $sec.CenterCols })
            $wCols = @(if ($sec.ContainsKey('WrapCols')) { $sec.WrapCols })
            $sdCols = @(if ($sec.ContainsKey('SignedCols')) { $sec.SignedCols })
            $invCols = @(if ($sec.ContainsKey('InvertSignedCols')) { $sec.InvertSignedCols })
            $stCol = if ($sec.ContainsKey('StatusCol')) { [int]$sec.StatusCol } else { 0 }
            $alCol = if ($sec.ContainsKey('AlertCol')) { [int]$sec.AlertCol } else { 0 }

            [void]$sb.AppendFormat('<row r="{0}" ht="26" customHeight="1">', $rNum)
            for ($c = 0; $c -lt $hdrs.Count; $c++) {
                [void]$sb.AppendFormat('<c r="{0}{1}" t="inlineStr" s="3"><is><t>{2}</t></is></c>', (ColLetterCmp ($c + 1)), $rNum, (XmlEscCmp $hdrs[$c]))
            }
            [void]$sb.Append('</row>')
            $rNum++

            if ($rows.Count -eq 0) {
                $sEndL = ColLetterCmp $hdrs.Count
                [void]$sb.AppendFormat('<row r="{0}" ht="22" customHeight="1">', $rNum)
                [void]$sb.AppendFormat('<c r="A{0}" t="inlineStr" s="29"><is><t>該当データなし</t></is></c>', $rNum)
                for ($c = 1; $c -lt $hdrs.Count; $c++) { [void]$sb.AppendFormat('<c r="{0}{1}" s="29"/>', (ColLetterCmp ($c + 1)), $rNum) }
                [void]$sb.Append('</row>')
                $merges.Add("A${rNum}:${sEndL}${rNum}")
                $rNum += 2
                continue
            }

            for ($r = 0; $r -lt $rows.Count; $r++) {
                $rowVals = $rows[$r]
                $isZebra = (($r % 2) -eq 1)
                $ht = 22.0
                if ($wCols.Count -gt 0) {
                    $wt = @(); $ww = @()
                    foreach ($wc in $wCols) {
                        $wt += [string]$rowVals[$wc - 1]
                        $ww += if (($wc - 1) -lt $ColWidths.Count) { [double]$ColWidths[$wc - 1] } else { 24.0 }
                    }
                    $ht = Get-EstimatedRowHeight -Texts $wt -ColWidths $ww -MinHeight 22.0 -MaxHeight 95.0
                }
                [void]$sb.AppendFormat([System.Globalization.CultureInfo]::InvariantCulture, '<row r="{0}" ht="{1:0.#}" customHeight="1">', $rNum, $ht)
                for ($c = 0; $c -lt $hdrs.Count; $c++) {
                    $col1 = $c + 1; $ref = (ColLetterCmp $col1) + $rNum; $v = $rowVals[$c]
                    if ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal]) {
                        $dv = [double]$v
                        $baseSt = if ($numSt.ContainsKey($col1)) { [int]$numSt[$col1] } else { 6 }
                        $st = if ($isZebra) { $baseSt + 1 } else { $baseSt }
                        if ($sdCols -contains $col1) {
                            $isNegWarn = if ($invCols -contains $col1) { ($dv -gt 1e-9) } else { ($dv -lt -1e-9) }
                            if ($isNegWarn) {
                                if ($baseSt -eq 14) { $st = 31 } elseif ($baseSt -eq 16) { $st = 32 } elseif ($baseSt -eq 18) { $st = 33 }
                            } elseif ([Math]::Abs($dv) -le 1e-9) {
                                if ($baseSt -eq 14) { $st = if ($isZebra) { 7 } else { 6 } } elseif ($baseSt -eq 16) { $st = if ($isZebra) { 9 } else { 8 } } elseif ($baseSt -eq 18) { $st = if ($isZebra) { 13 } else { 12 } }
                            }
                        }
                        [void]$sb.AppendFormat([System.Globalization.CultureInfo]::InvariantCulture, '<c r="{0}" s="{1}"><v>{2}</v></c>', $ref, $st, $v)
                    }
                    else {
                        $sv = [string]$v
                        $st = if ($wCols -contains $col1) { if ($isZebra) { 23 } else { 22 } } elseif ($cCols -contains $col1) { if ($isZebra) { 25 } else { 24 } } else { if ($isZebra) { 5 } else { 4 } }
                        if ($col1 -eq $stCol) {
                            if ($sv -match '(新規発生|悪化|消失|削除|減少|要注意)') { $st = 26 }
                            elseif ($sv -match '(解消|改善|新規|増加)') { $st = 27 }
                            elseif ($sv -match '変化なし|継続|所見') { $st = if ($isZebra) { 30 } else { 29 } }
                        }
                        elseif ($col1 -eq $alCol -and ($sv -eq '警告' -or $sv -eq '注意' -or $sv -eq '高' -or $sv -match 'ギャップ')) {
                            if ($sv -ne '高' -or $stCol -le 0 -or ([string]$rowVals[$stCol - 1]) -ne '解消') {
                                $st = if ($cCols -contains $col1) { 26 } else { 46 }
                            }
                        }
                        [void]$sb.AppendFormat('<c r="{0}" t="inlineStr" s="{1}"><is><t>{2}</t></is></c>', $ref, $st, (XmlEscCmp $sv))
                    }
                }
                [void]$sb.Append('</row>')
                $rNum++
            }
            $rNum++
        }

        [void]$sb.Append('</sheetData>')
        [void]$sb.AppendFormat('<mergeCells count="{0}">', $merges.Count)
        foreach ($mg in $merges) { [void]$sb.AppendFormat('<mergeCell ref="{0}"/>', $mg) }
        [void]$sb.Append('</mergeCells>')
        [void]$sb.Append('<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>')
        [void]$sb.Append('<pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/>')
        [void]$sb.Append('</worksheet>')
        return $sb.ToString()
    }

    $OutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
    $stagingPath = New-AnalysisStagingPath $OutputPath
    if ($null -ne $LogCallback) { & $LogCallback '比較Excelレポート生成中 (OpenXML / 8シート)' 45 }

    $subInfo = "比較日時: {0} | 開始前: {1} | 開始後: {2} | 照合: {3}" -f $Comparison.ComparedAt.ToString('yyyy/MM/dd HH:mm:ss'), $Comparison.BeforeFileName, $Comparison.AfterFileName, $Comparison.PrecisionLabel

    # Sheet 1 rows (Summary breakdown)
    $bucketDefs = @(
        @{ Status = '追加（新規テーブル）'; Desc = 'テスト開始前になく、開始後に新規追加されたテーブル' },
        @{ Status = '追加（レコード増加）'; Desc = '既存テーブルでレコード数が増加 (INSERT等)' },
        @{ Status = '追加・更新';           Desc = 'レコード増加＋既存値/削除/スキーマ更新を検出 (更新・追加両シートに掲載)' },
        @{ Status = '更新（データ更新）';   Desc = 'レコード数変動なしで値ハッシュ・統計・分布・数値・日時・サイズに変化 (UPDATE等)' },
        @{ Status = '更新（レコード減少）'; Desc = '既存テーブルでレコード数が減少 (DELETE/パージ等・要注意)' },
        @{ Status = '削除（テーブル消失）'; Desc = 'テスト開始前に存在したが開始後に未検出 (要注意)' },
        @{ Status = '変更なし';             Desc = 'レコード数・サイズ・全カラム統計・分布・数値・日時統計・内容ハッシュすべて完全一致' }
    )
    $totTbl = [Math]::Max(1, $Comparison.TotalUniqueTableCount)
    $s1Rows = @()
    foreach ($bd in $bucketDefs) {
        $stName = $bd.Status
        $matched = @($Comparison.AllTables | Where-Object { $_.PrimaryStatus -eq $stName })
        $bR = 0L; $aR = 0L; $dR = 0L; $dS = 0.0; $dI = 0.0; $dc = 0
        foreach ($mt in $matched) { $bR += $mt.BeforeRows; $aR += $mt.AfterRows; $dR += $mt.RowDelta; $dS += $mt.SizeDeltaKB; $dI += $mt.InnoDbDeltaKB; $dc += $mt.ColumnDiffs.Count }
        $flag = if ($stName -match '(削除|減少)' -and $matched.Count -gt 0) { '要注意' } elseif ($stName -eq '変更なし') { '変化なし' } elseif ($matched.Count -gt 0) { '変更検出' } else { '-' }
        $s1Rows += , @($stName, [double]$matched.Count, [double]($matched.Count / $totTbl), [double]$bR, [double]$aR, [double]$dR, [Math]::Round($dS, 2), [Math]::Round($dI, 1), [double]$dc, $flag, $bd.Desc)
    }
    $s1Total = @('合計', [double]$Comparison.TotalUniqueTableCount, 1.0, [double]$Comparison.BeforeTotalRows, [double]$Comparison.AfterTotalRows, [double]$Comparison.TotalRowDelta, [double]$Comparison.TotalSizeDeltaKB, [double]$Comparison.TotalInnoDbDeltaKB, [double]$Comparison.AllColumnDiffs.Count, '', ("全{0}列 (変更{1}/不変{2}) | {3}" -f $Comparison.TotalUniqueColumnCount, $Comparison.ChangedColumnTotal, $Comparison.UnchangedColumnTotal, $Comparison.PrecisionLabel))

    # Sheet 2 rows (All tables)
    $s2Rows = @()
    $rk = 0
    foreach ($t in $Comparison.AllTables) {
        $rk++
        $gr = if ($t.HasRowGrowthRate) { [double]$t.RowGrowthRate } elseif ($t.AfterRows -gt 0 -and $t.BeforeRows -eq 0) { '新規' } else { '-' }
        $vel = if ($t.HasVelocity) { [Math]::Round($t.RowsPerMinute, 1) } else { '-' }
        $chTxt = if ($t.HasRowChurn) { $t.ChurnSummary } else { '-' }
        $s2Rows += , @($rk, $t.TableName, $t.PrimaryStatus, $t.InferredRole, [double]$t.BeforeRows, [double]$t.AfterRows, [double]$t.RowDelta, $gr, $vel, $chTxt, [double]$t.BeforeCols, [double]$t.AfterCols, [double]$t.ColDelta, [double]$t.ChangedColumnCount, [double]$t.UnchangedColumnCount, $t.ContentHashStatus, [double]$t.BeforeSizeKB, [double]$t.AfterSizeKB, [double]$t.SizeDeltaKB, [double]$t.BeforeInnoDbKB, [double]$t.AfterInnoDbKB, [double]$t.InnoDbDeltaKB, $t.DiffSummary, $t.AlertNote)
    }
    $totGr = if ($Comparison.HasTotalRowGrowthRate) { [double]$Comparison.TotalRowGrowthRate } else { '-' }
    $totVel = if ($Comparison.HasElapsedTime) { [Math]::Round($Comparison.TotalRowsPerMinute, 1) } else { '-' }
    $totChurn = if ($Comparison.HasAnyRowChurn) { ("+{0:#,##0} / 更新{1:#,##0} / -{2:#,##0}" -f $Comparison.TotalInsertedRows, $Comparison.TotalUpdatedRows, $Comparison.TotalDeletedRows) } else { '-' }
    $s2Total = @('', '合計', ("計 {0} テーブル" -f $Comparison.AllTables.Count), '', [double]$Comparison.BeforeTotalRows, [double]$Comparison.AfterTotalRows, [double]$Comparison.TotalRowDelta, $totGr, $totVel, $totChurn, '', '', '', [double]$Comparison.ChangedColumnTotal, [double]$Comparison.UnchangedColumnTotal, '', [Math]::Round($Comparison.BeforeTotalSizeKB, 2), [Math]::Round($Comparison.AfterTotalSizeKB, 2), [double]$Comparison.TotalSizeDeltaKB, [Math]::Round($Comparison.BeforeTotalInnoDbKB, 1), [Math]::Round($Comparison.AfterTotalInnoDbKB, 1), [double]$Comparison.TotalInnoDbDeltaKB, '', '')

    # Sheet 3 rows (Updated tables + deleted warning entries if any)
    $s3Rows = @()
    $rk = 0
    foreach ($t in @($Comparison.UpdatedTables) + @($Comparison.DeletedTables)) {
        $rk++
        $chTxt = if ($t.HasRowChurn) { $t.ChurnSummary } else { '-' }
        $s3Rows += , @($rk, $t.TableName, $t.PrimaryStatus, $t.InferredRole, [double]$t.BeforeRows, [double]$t.AfterRows, [double]$t.RowDelta, $chTxt, $t.ContentHashStatus, [double]$t.BeforeAvgRowBytes, [double]$t.AfterAvgRowBytes, [double]$t.AvgRowBytesDelta, [double]$t.SizeDeltaKB, [double]$t.ChangedColumnCount, $t.ChangedColumnsSummary, $t.UnchangedColumnsSummary, $t.DiffSummary, $t.AlertNote)
    }

    # Sheet 4 rows (Added tables)
    $s4Rows = @()
    $rk = 0; $addB = 0L; $addA = 0L; $addD = 0L; $addSd = 0.0; $addId = 0.0
    foreach ($t in $Comparison.AddedTables) {
        $rk++; $addB += $t.BeforeRows; $addA += $t.AfterRows; $addD += $t.RowDelta; $addSd += $t.SizeDeltaKB; $addId += $t.InnoDbDeltaKB
        $gr = if ($t.HasRowGrowthRate) { [double]$t.RowGrowthRate } elseif ($t.AfterRows -gt 0 -and $t.BeforeRows -eq 0) { '新規' } else { '-' }
        $vel = if ($t.HasVelocity) { [Math]::Round($t.RowsPerMinute, 1) } else { '-' }
        $chTxt = if ($t.HasRowChurn) { $t.ChurnSummary } else { '-' }
        $s4Rows += , @($rk, $t.TableName, $t.AddedType, $t.InferredRole, [double]$t.BeforeRows, [double]$t.AfterRows, [double]$t.RowDelta, $gr, $vel, $chTxt, [double]$t.AfterCols, [double]$t.BeforeSizeKB, [double]$t.AfterSizeKB, [double]$t.SizeDeltaKB, [double]$t.InnoDbDeltaKB, $(if ($t.CandidatePks) { $t.CandidatePks } else { '-' }), $t.DiffSummary, $t.AlertNote)
    }
    $addTotGr = if ($addB -gt 0) { [double]($addD / $addB) } else { '-' }
    $s4Total = @('', '合計', ("計 {0} テーブル" -f $Comparison.AddedTables.Count), '', [double]$addB, [double]$addA, [double]$addD, $addTotGr, '', '', '', '', '', [Math]::Round($addSd, 2), [Math]::Round($addId, 1), '', '', '')

    # Sheet 5 rows (Unchanged tables)
    $s5Rows = @()
    $rk = 0; $uncR = 0L; $uncS = 0.0; $uncI = 0.0
    foreach ($t in $Comparison.UnchangedTables) {
        $rk++; $uncR += $t.AfterRows; $uncS += $t.AfterSizeKB; $uncI += $t.AfterInnoDbKB
        $basis = if ($t.ContentHashStatus -eq '一致') { 'ハッシュ完全一致' } else { '全統計一致' }
        $s5Rows += , @($rk, $t.TableName, $t.PrimaryStatus, $t.InferredRole, $basis, [double]$t.AfterRows, [double]$t.AfterCols, [double]$t.AfterSizeKB, [double]$t.AfterInnoDbKB, [double]$t.AfterAvgRowBytes, [double]$t.AfterAvgNullRate, $(if ($t.CandidatePks) { $t.CandidatePks } else { '-' }), $(if ($t.DateColumns) { $t.DateColumns } else { '-' }), $t.AfterWarning)
    }
    $s5Total = @('', '合計', ("計 {0} テーブル" -f $Comparison.UnchangedTables.Count), '', '', [double]$uncR, '', [Math]::Round($uncS, 2), [Math]::Round($uncI, 1), '', '', '', '', '')

    # Sheet 6 rows (Column/Stat diff details)
    $s6Rows = @()
    foreach ($d in $Comparison.AllColumnDiffs) {
        $s6Rows += , @($d.TableName, $d.ColumnName, $d.Category, $d.MetricName, $d.BeforeValue, $d.AfterValue, $d.DeltaSummary, $d.Severity)
    }

    # Sheet 7 rows (Complete Changed & Unchanged Column Comparison Matrix)
    $s7Rows = @()
    $rk = 0
    foreach ($c in $Comparison.AllColumnComparisons) {
        $rk++
        $s7Rows += , @($rk, $c.TableName, $c.ColumnName, $c.TableStatus, $c.ColumnStatus, $c.ContentHashStatus, $c.TypeDisplay, $c.DdlKeyDisplay, [double]$c.NonNullDelta, $c.NullChangeDisplay, $c.DistinctChangeDisplay, $c.ByteLenDisplay, $c.MinValDisplay, $c.MaxValDisplay, $c.Top1Display, $c.NumericSumDeltaDisplay, [double]$c.DiffCount, $c.ChangeSummary, $c.Severity)
    }
    $s7Total = @('', '合計', ("計 {0} カラム" -f $Comparison.TotalUniqueColumnCount), '', ("変更 {0} 列 / 不変 {1} 列" -f $Comparison.ChangedColumnTotal, $Comparison.UnchangedColumnTotal), '', '', ("追加 {0} / 削除 {1} / DDL変更 {2}" -f $Comparison.AddedColumnTotal, $Comparison.RemovedColumnTotal, $Comparison.SchemaChangedColumnTotal), '', '', '', '', '', '', '', '', [double]$Comparison.AllColumnDiffs.Count, '全テーブルの全カラム照合集計', '')

    # Sheet 8 sections (Warnings, Relationships, Time-Series, DB Physical/Server/DDL, JMeter, Insights)
    $s8WarnRows = @()
    foreach ($w in $Comparison.WarningComparisons) {
        $act = if ([string]::IsNullOrEmpty($w.PerfImpact)) { $w.Recommendation } else { ("【影響】{0} / 【対策】{1}" -f $w.PerfImpact, $w.Recommendation) }
        $s8WarnRows += , @($w.Status, $w.Severity, $w.Category, $w.TableName, $w.ColumnName, $w.BeforeMetric, $w.AfterMetric, $act)
    }
    $s8RelRows = @()
    foreach ($r in $Comparison.RelationshipComparisons) {
        $s8RelRows += , @($r.Status, $r.RelationKey, [double]$r.BeforeMatchRate, [double]$r.AfterMatchRate, [double]$r.BeforeOrphanRows, [double]$r.AfterOrphanRows, [double]$r.OrphanDelta, ("{0:#,##0} → {1:#,##0}" -f $r.BeforeFanOutMax, $r.AfterFanOutMax), ("{0} → {1}" -f $r.BeforeRisk, $r.AfterRisk), $r.Note)
    }
    $s8TsRows = @()
    foreach ($ts in ($Comparison.TimeSeriesDeltas | Select-Object -First 80)) {
        $gr = if ($ts.HasGrowthRate) { [double]$ts.GrowthRate } else { '新規' }
        $s8TsRows += , @($ts.Granularity, $ts.TableColumn, $ts.PeriodLabel, [double]$ts.BeforeCount, [double]$ts.AfterCount, [double]$ts.DeltaCount, $gr)
    }
    $s8PhysRows = @()
    foreach ($t in ($Comparison.AllTables | Where-Object { $_.HasDbPhysicalMeta })) {
        $aiD = if (-not $t.HasAutoIncrement) { '-' } elseif ($t.AutoIncrementGap -gt 0) { ("+{0:#,##0} (ギャップ {1:#,##0}件)" -f $t.AutoIncrementDelta, $t.AutoIncrementGap) } else { ("{0}{1:#,##0}" -f $(if ($t.AutoIncrementDelta -gt 0) { '+' } else { '' }), $t.AutoIncrementDelta) }
        $s8PhysRows += , @($t.TableName, [double]$t.BeforeDataLengthKB, [double]$t.AfterDataLengthKB, [double]$t.DataLengthDeltaKB, [double]$t.BeforeIndexLengthKB, [double]$t.AfterIndexLengthKB, [double]$t.IndexLengthDeltaKB, [double]$t.DataFreeDeltaKB, $(if ($t.HasAutoIncrement) { [double]$t.BeforeAutoIncrement } else { '-' }), $(if ($t.HasAutoIncrement) { [double]$t.AfterAutoIncrement } else { '-' }), $aiD)
    }
    $s8SrvRows = @()
    foreach ($sm in $Comparison.ServerMetricComparisons) {
        $s8SrvRows += , @($sm.Category, $sm.MetricName, $sm.DisplayBefore, $sm.DisplayAfter, $sm.DisplayDelta, $sm.RatePerSecDisplay, $sm.Severity, $sm.Note)
    }
    $s8SchRows = @()
    foreach ($sd in $Comparison.DbSchemaDiffs) {
        $s8SchRows += , @($sd.TableName, $sd.ObjectType, $sd.ObjectName, $sd.ChangeType, $sd.BeforeDef, $sd.AfterDef, $sd.Severity, $sd.Note)
    }
    $s8JmRows = @()
    foreach ($jm in $Comparison.JMeterComparisons) {
        $s8JmRows += , @($jm.Status, $jm.TableName, $jm.KeyColumn, $jm.DataType, [double]$jm.BeforeUnique, [double]$jm.AfterUnique, [double]$jm.UniqueDelta, ("{0} → {1}" -f $jm.BeforeSuitability, $jm.AfterSuitability))
    }
    $s8InsRows = @()
    $insIdx = 0
    foreach ($ins in $Comparison.Insights) {
        $insIdx++
        $s8InsRows += , @($insIdx, $(if ($ins -match '要注意') { '要注意' } else { '所見' }), $ins)
    }

    $s8Sections = @(
        @{ Title = ("1. 性能リスク・警告の推移 (新規 {0} / 悪化 {1} / 解消 {2} / 継続 {3})" -f $Comparison.NewWarningCount, $Comparison.WorsenedWarningCount, $Comparison.ResolvedWarningCount, $Comparison.OngoingWarningCount); Headers = @('変化状態','重要度','カテゴリ','テーブル名','カラム名','開始前指標','開始後指標','影響・推奨アクション'); Rows = $s8WarnRows; CenterCols = @(1,2,3); WrapCols = @(6,7,8); StatusCol = 1; AlertCol = 2 },
        @{ Title = ("2. テーブル間リレーション・FK整合性・ファンアウト変化 (孤立行増減: {0}{1:#,##0}行)" -f $(if ($Comparison.OrphanRowDelta -gt 0) { '+' } else { '' }), $Comparison.OrphanRowDelta); Headers = @('変化状態','リレーション (親.PK → 子.FK)','開始前一致率','開始後一致率','開始前孤立行','開始後孤立行','孤立行増減','最大ファンアウト(前→後)','リスク変化','所見'); Rows = $s8RelRows; NumStyles = @{3=10;4=10;5=6;6=6;7=14}; SignedCols = @(7); InvertSignedCols = @(7); CenterCols = @(1,8,9); WrapCols = @(2,9,10); StatusCol = 1 },
        @{ Title = '3. テスト期間中の時系列アクティビティ増分 (月別・日別・時間帯別)'; Headers = @('粒度','対象テーブル.日時カラム','期間・時間帯','開始前件数','開始後件数','増減件数','増減率'); Rows = $s8TsRows; NumStyles = @{4=6;5=6;6=14;7=20}; SignedCols = @(6); CenterCols = @(1,3) },
        @{ Title = '4A. Aurora / MySQL 実物理容量・断片化・AUTO_INCREMENT 推移 (DB直接分析時)'; Headers = @('テーブル名','実Data前(KB)','実Data後(KB)','Data増減(KB)','実Index前(KB)','実Index後(KB)','Index増減(KB)','断片化空き増減(KB)','AI開始前','AI開始後','AUTO_INCREMENT増減/ギャップ'); Rows = $s8PhysRows; NumStyles = @{2=12;3=12;4=18;5=12;6=12;7=18;8=18;9=6;10=6}; SignedCols = @(4,7,8); CenterCols = @(11); WrapCols = @(11); AlertCol = 11 },
        @{ Title = '4B. Aurora / MySQL サーバー負荷統計差分 (SHOW GLOBAL STATUS)'; Headers = @('カテゴリ','サーバー指標','開始前','開始後','テスト期間増分','平均レート(/秒)','判定','性能試験上の意味・所見'); Rows = $s8SrvRows; CenterCols = @(1,7); WrapCols = @(8); AlertCol = 7 },
        @{ Title = '4C. DDLスキーマ・インデックス・外部キー制約 変更一覧 (DB直接分析時)'; Headers = @('テーブル名','種別','対象名','変更区分','開始前定義','開始後定義','重要度','備考'); Rows = $s8SchRows; CenterCols = @(2,4,7); WrapCols = @(5,6,8); AlertCol = 7 },
        @{ Title = '5. JMeter 並列適性・一意キー候補の変化'; Headers = @('変化状態','テーブル名','キーカラム','データ型','開始前キー数','開始後キー数','キー数増減','JMeter並列適性 (前→後)'); Rows = $s8JmRows; NumStyles = @{5=6;6=6;7=14}; SignedCols = @(7); CenterCols = @(1,4); WrapCols = @(8); StatusCol = 1 },
        @{ Title = '6. テスト前後比較 総合所見・チェックポイント'; Headers = @('No','区分','総合所見'); Rows = $s8InsRows; NumStyles = @{1=6}; CenterCols = @(1,2); WrapCols = @(3); StatusCol = 2 }
    )

    $sheetNames = @(
        '01_比較サマリー', '02_全テーブル比較一覧', '03_更新のあったテーブル',
        '04_追加のあったテーブル', '05_変更のなかったテーブル', '06_カラム・統計差分詳細',
        '07_全カラム比較一覧', '08_性能・DB・時系列・品質変化'
    )

    $sheetXmls = @(
        (Build-CmpSheetXml -Title '性能試験 テスト開始前・開始後 分析ファイル比較レポート — 比較サマリー' -Subtitle $subInfo -Headers @('判定区分','テーブル数','構成比','開始前レコード数','開始後レコード数','レコード増減','CSV増減(KB)','InnoDB増減(KB)','検出差分項目数','状態','判定基準・説明') -Rows $s1Rows -ColWidths @(22,12,10,15,15,14,13,15,14,10,48) -NumColStyles @{2=6;3=10;4=6;5=6;6=14;7=16;8=18;9=6} -StatusCol 1 -AlertCol 10 -SignedDeltaCols @(6,7,8) -CenterCols @(1,10) -WrapCols @(11) -TotalRowValues $s1Total -TotalColStyles @{2=35;3=41;4=35;5=35;6=36;7=38;8=40;9=35} -IncludeChartDrawing $true),
        (Build-CmpSheetXml -Title '全テーブル テスト開始前・開始後 比較一覧' -Subtitle $subInfo -Headers @('No','テーブル名','判定区分','推定役割','開始前レコード数','開始後レコード数','レコード増減','増減率','増減速度(件/分)','行変動内訳 (挿入/更新/削除)','開始前列数','開始後列数','列数増減','変更列数','不変列数','内容ハッシュ','開始前CSV(KB)','開始後CSV(KB)','CSV増減(KB)','開始前InnoDB(KB)','開始後InnoDB(KB)','InnoDB増減(KB)','主な変更内容・所見','警告・注意点') -Rows $s2Rows -ColWidths @(5,22,20,15,13,13,12,10,11,26,9,9,9,9,9,12,12,12,12,13,13,13,38,28) -NumColStyles @{1=6;5=6;6=6;7=14;8=20;9=18;11=6;12=6;13=14;14=6;15=6;17=8;18=8;19=16;20=12;21=12;22=18} -StatusCol 3 -AlertCol 24 -SignedDeltaCols @(7,9,13,19,22) -CenterCols @(1,3,4,11,12,13,14,15,16) -WrapCols @(10,23,24) -TotalRowValues $s2Total -TotalColStyles @{5=35;6=35;7=36;8=47;9=40;14=35;15=35;17=37;18=37;19=38;20=39;21=39;22=40}),
        (Build-CmpSheetXml -Title '更新のあったテーブル一覧 (既存データ更新・レコード減少・追加＋更新併用)' -Subtitle $subInfo -Headers @('No','テーブル名','判定区分','推定役割','開始前レコード数','開始後レコード数','レコード増減','行変動内訳 (挿入/更新/削除)','内容ハッシュ','開始前平均行長(B)','開始後平均行長(B)','平均行長増減(B)','CSV増減(KB)','変化列数','変更検出カラム','変更なしカラム','具体的な変更内容','警告・注意点') -Rows $s3Rows -ColWidths @(5,22,20,15,13,13,12,24,11,13,13,12,12,9,24,24,38,28) -NumColStyles @{1=6;5=6;6=6;7=14;10=12;11=12;12=18;13=16;14=6} -StatusCol 3 -AlertCol 18 -SignedDeltaCols @(7,12,13) -CenterCols @(1,3,4,9,14) -WrapCols @(8,15,16,17,18)),
        (Build-CmpSheetXml -Title '追加のあったテーブル一覧 (新規テーブル追加 ＆ レコード増加テーブル)' -Subtitle $subInfo -Headers @('No','テーブル名','追加種別','推定役割','開始前レコード数','開始後レコード数','追加レコード数','増加率','増加速度(件/分)','行変動内訳','列数','開始前CSV(KB)','開始後CSV(KB)','CSV増加(KB)','InnoDB増加(KB)','PK/UK候補','変更サマリー','警告・注意点') -Rows $s4Rows -ColWidths @(5,22,18,15,13,13,13,10,11,24,7,12,12,12,13,18,36,26) -NumColStyles @{1=6;5=6;6=6;7=14;8=20;9=18;11=6;12=8;13=8;14=16;15=18} -StatusCol 3 -AlertCol 18 -SignedDeltaCols @(7,9,14,15) -CenterCols @(1,3,4,11) -WrapCols @(10,17,18) -TotalRowValues $s4Total -TotalColStyles @{5=35;6=35;7=36;8=47;14=38;15=40}),
        (Build-CmpSheetXml -Title '変更のなかったテーブル一覧 (開始前後で件数・サイズ・統計・値ハッシュ完全一致)' -Subtitle $subInfo -Headers @('No','テーブル名','判定区分','推定役割','照合根拠','レコード数','列数','CSVサイズ(KB)','推定InnoDB(KB)','平均行長(B)','平均NULL率','PK/UK候補','日時カラム','既存警告・注意点') -Rows $s5Rows -ColWidths @(5,24,14,15,16,13,8,14,15,12,11,20,20,30) -NumColStyles @{1=6;6=6;7=6;8=8;9=12;10=12;11=10} -StatusCol 3 -AlertCol 14 -CenterCols @(1,3,4,5,7) -WrapCols @(12,13,14) -TotalRowValues $s5Total -TotalColStyles @{6=35;8=37;9=39}),
        (Build-CmpSheetXml -Title 'カラム・分布・数値・日時・DDL・ハッシュ 統計差分詳細一覧' -Subtitle $subInfo -Headers @('テーブル名','カラム名','差分カテゴリ','指標名','テスト開始前','テスト開始後','差分・変化内容','重要度') -Rows $s6Rows -ColWidths @(22,20,16,24,24,24,28,10) -AlertCol 8 -CenterCols @(3,8) -WrapCols @(5,6,7)),
        (Build-CmpSheetXml -Title '全カラム 変更あり・変更なし 比較マトリクス' -Subtitle $subInfo -Headers @('No','テーブル名','カラム名','テーブル判定','カラム判定','内容ハッシュ','推定データ型 (前→後)','DDL型 / キー・役割','有効件数増減','NULL件数 (前→後)','ユニーク数 (前→後)','バイト長 (平均/最大)','最小値/最古 (前→後)','最大値/最新 (前→後)','最頻値Top1 (前→後)','数値合計増減','差分件数','変更内容サマリー / 無変更根拠','重要度') -Rows $s7Rows -ColWidths @(5,20,20,18,22,12,20,22,12,18,18,22,20,20,22,14,9,36,10) -NumColStyles @{1=6;9=14;17=6} -StatusCol 5 -AlertCol 19 -SignedDeltaCols @(9) -CenterCols @(1,4,5,6,17,19) -WrapCols @(7,8,12,13,14,15,18) -TotalRowValues $s7Total -TotalColStyles @{17=35}),
        (Build-CmpMultiSectionSheetXml -Title 'テスト有意情報サマリー (性能警告推移・FK整合性・時系列増分・Aurora/MySQL負荷＆物理容量・JMeter適性)' -Subtitle $subInfo -ColWidths @(16,24,16,16,15,15,16,18,22,34,24) -Sections $s8Sections)
    )

    $stylesXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="8">
    <numFmt numFmtId="164" formatCode="#,##0"/>
    <numFmt numFmtId="165" formatCode="#,##0.00"/>
    <numFmt numFmtId="166" formatCode="0.0%"/>
    <numFmt numFmtId="167" formatCode="+#,##0;-#,##0;0"/>
    <numFmt numFmtId="168" formatCode="+#,##0.00;-#,##0.00;0.00"/>
    <numFmt numFmtId="169" formatCode="+0.0%;-0.0%;0.0%"/>
    <numFmt numFmtId="170" formatCode="#,##0.0"/>
    <numFmt numFmtId="171" formatCode="+#,##0.0;-#,##0.0;0.0"/>
  </numFmts>
  <fonts count="8">
    <font><sz val="9.5"/><color rgb="FF333333"/><name val="Meiryo UI"/></font>
    <font><b/><sz val="13.5"/><color rgb="FFFFFFFF"/><name val="Meiryo UI"/></font>
    <font><sz val="9"/><color rgb="FF333333"/><name val="Meiryo UI"/></font>
    <font><b/><sz val="9.5"/><color rgb="FFFFFFFF"/><name val="Meiryo UI"/></font>
    <font><b/><sz val="9.5"/><color rgb="FF800020"/><name val="Meiryo UI"/></font>
    <font><b/><sz val="9.5"/><color rgb="FF1A2B4C"/><name val="Meiryo UI"/></font>
    <font><b/><sz val="9.5"/><color rgb="FF333333"/><name val="Meiryo UI"/></font>
    <font><sz val="9.5"/><color rgb="FF595959"/><name val="Meiryo UI"/></font>
  </fonts>
  <fills count="8">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1A2B4C"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF7F7F7"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF5F5F5"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFEAEAEA"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/></patternFill></fill>
  </fills>
  <borders count="5">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFD9D9D9"/></left>
      <right style="thin"><color rgb="FFD9D9D9"/></right>
      <top style="thin"><color rgb="FFD9D9D9"/></top>
      <bottom style="thin"><color rgb="FFD9D9D9"/></bottom>
      <diagonal/>
    </border>
    <border>
      <left style="thin"><color rgb="FFD9D9D9"/></left>
      <right style="thin"><color rgb="FFD9D9D9"/></right>
      <top style="thin"><color rgb="FF8C8C8C"/></top>
      <bottom style="medium"><color rgb="FF1A2B4C"/></bottom>
      <diagonal/>
    </border>
    <border>
      <left/><right/><top/>
      <bottom style="thin"><color rgb="FF1A2B4C"/></bottom>
      <diagonal/>
    </border>
    <border>
      <left style="thin"><color rgb="FF4D5B75"/></left>
      <right style="thin"><color rgb="FF4D5B75"/></right>
      <top style="thin"><color rgb="FF1A2B4C"/></top>
      <bottom style="thin"><color rgb="FF1A2B4C"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="50">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center" indent="1"/></xf>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="3" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" indent="1"/></xf>
    <xf numFmtId="0" fontId="3" fillId="2" borderId="4" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="7" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="164" fontId="0" fillId="7" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="164" fontId="0" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="165" fontId="0" fillId="7" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="165" fontId="0" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="166" fontId="0" fillId="7" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="166" fontId="0" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="170" fontId="0" fillId="7" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="170" fontId="0" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="167" fontId="5" fillId="7" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="167" fontId="5" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="168" fontId="5" fillId="7" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="168" fontId="5" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="171" fontId="5" fillId="7" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="171" fontId="5" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="169" fontId="0" fillId="7" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="169" fontId="0" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="7" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="7" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="4" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="5" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="6" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="7" fillId="7" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="7" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="167" fontId="4" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="168" fontId="4" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="171" fontId="4" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="0" fontId="5" fillId="6" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="164" fontId="5" fillId="6" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="167" fontId="5" fillId="6" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="165" fontId="5" fillId="6" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="168" fontId="5" fillId="6" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="170" fontId="5" fillId="6" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="171" fontId="5" fillId="6" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="166" fontId="5" fillId="6" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="167" fontId="4" fillId="6" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="168" fontId="4" fillId="6" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="171" fontId="4" fillId="6" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="166" fontId="4" fillId="6" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="0" fontId="4" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="169" fontId="5" fillId="6" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="169" fontId="4" fillId="6" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="0" fontId="5" fillId="3" borderId="3" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" indent="1"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
'@

    if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }
    $zip = [System.IO.Compression.ZipFile]::Open($stagingPath, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        function Add-CmpZipText([string]$entryName, [string]$text) {
            if ($null -ne $PumpCallback) { & $PumpCallback }
            if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }
            $e = $zip.CreateEntry($entryName, [System.IO.Compression.CompressionLevel]::Optimal)
            $sw = New-Object System.IO.StreamWriter($e.Open(), (New-Object System.Text.UTF8Encoding($false)))
            try { $sw.Write($text) } finally { $sw.Dispose() }
        }

        $ctExtra = '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
        $sheetCt = ($sheetNames | ForEach-Object -Begin { $i = 0 } -Process { $i++; "<Override PartName=`"/xl/worksheets/sheet$i.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>" }) -join ''
        Add-CmpZipText '[Content_Types].xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' + $sheetCt + $ctExtra + '</Types>')
        Add-CmpZipText '_rels/.rels' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'

        $wbSheets = ''
        $wbRels = '<Relationship Id="rIdStyle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        for ($i = 0; $i -lt $sheetNames.Count; $i++) {
            $idx = $i + 1
            $wbSheets += "<sheet name=`"$(XmlEscCmp $sheetNames[$i])`" sheetId=`"$idx`" r:id=`"rId$idx`"/>"
            $wbRels += "<Relationship Id=`"rId$idx`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$idx.xml`"/>"
            Add-CmpZipText "xl/worksheets/sheet$idx.xml" $sheetXmls[$i]
        }

        Add-CmpZipText 'xl/workbook.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + $wbSheets + '</sheets></workbook>')
        Add-CmpZipText 'xl/_rels/workbook.xml.rels' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + $wbRels + '</Relationships>')
        Add-CmpZipText 'xl/styles.xml' $stylesXml

        $lastRow1 = 4 + $s1Rows.Count
        Add-CmpZipText 'xl/worksheets/_rels/sheet1.xml.rels' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>'
        Add-CmpZipText 'xl/drawings/_rels/drawing1.xml.rels' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>'
        Add-CmpZipText 'xl/drawings/drawing1.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + ($lastRow1 + 3) + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + ($lastRow1 + 18) + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Comparison Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>')
        $maxBucketCnt = 0.0
        foreach ($sr in $s1Rows) {
            if ($sr.Count -gt 1 -and [double]$sr[1] -gt $maxBucketCnt) { $maxBucketCnt = [double]$sr[1] }
        }
        $majorUnitXml = if ($maxBucketCnt -le 10) { '<c:majorUnit val="1"/>' } else { '' }
        Add-CmpZipText 'xl/charts/chart1.xml' ("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><c:chartSpace xmlns:c=`"http://schemas.openxmlformats.org/drawingml/2006/chart`" xmlns:a=`"http://schemas.openxmlformats.org/drawingml/2006/main`"><c:spPr><a:solidFill><a:srgbClr val=`"FFFFFF`"/></a:solidFill><a:ln w=`"9525`"><a:solidFill><a:srgbClr val=`"D9D9D9`"/></a:solidFill></a:ln></c:spPr><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr b=`"1`" sz=`"1050`"><a:solidFill><a:srgbClr val=`"1A2B4C`"/></a:solidFill><a:latin typeface=`"Meiryo UI`"/><a:ea typeface=`"Meiryo UI`"/></a:defRPr></a:pPr><a:r><a:rPr lang=`"ja-JP`" b=`"1`" sz=`"1050`"><a:solidFill><a:srgbClr val=`"1A2B4C`"/></a:solidFill><a:latin typeface=`"Meiryo UI`"/><a:ea typeface=`"Meiryo UI`"/></a:rPr><a:t>変更区分別テーブル数</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=`"0`"/></c:title><c:plotArea><c:spPr><a:solidFill><a:srgbClr val=`"FFFFFF`"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr><c:barChart><c:barDir val=`"bar`"/><c:grouping val=`"clustered`"/><c:gapWidth val=`"70`"/><c:ser><c:idx val=`"0`"/><c:order val=`"0`"/><c:spPr><a:solidFill><a:srgbClr val=`"1A2B4C`"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr><c:dLbls><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz=`"800`"><a:solidFill><a:srgbClr val=`"333333`"/></a:solidFill><a:latin typeface=`"Meiryo UI`"/><a:ea typeface=`"Meiryo UI`"/></a:defRPr></a:pPr><a:endParaRPr lang=`"ja-JP`"/></a:p></c:txPr><c:dLblPos val=`"outEnd`"/><c:showLegendKey val=`"0`"/><c:showVal val=`"1`"/><c:showCatName val=`"0`"/><c:showSerName val=`"0`"/><c:showPercent val=`"0`"/><c:showBubbleSize val=`"0`"/></c:dLbls><c:cat><c:strRef><c:f>'01_比較サマリー'!`$A`$5:`$A`$$lastRow1</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>'01_比較サマリー'!`$B`$5:`$B`$$lastRow1</c:f></c:numRef></c:val></c:ser><c:axId val=`"1`"/><c:axId val=`"2`"/></c:barChart><c:catAx><c:axId val=`"1`"/><c:scaling><c:orientation val=`"maxMin`"/></c:scaling><c:delete val=`"0`"/><c:axPos val=`"l`"/><c:tickLblPos val=`"low`"/><c:spPr><a:ln w=`"9525`"><a:solidFill><a:srgbClr val=`"D9D9D9`"/></a:solidFill></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz=`"850`"><a:solidFill><a:srgbClr val=`"333333`"/></a:solidFill><a:latin typeface=`"Meiryo UI`"/><a:ea typeface=`"Meiryo UI`"/></a:defRPr></a:pPr><a:endParaRPr lang=`"ja-JP`"/></a:p></c:txPr><c:crossAx val=`"2`"/></c:catAx><c:valAx><c:axId val=`"2`"/><c:scaling><c:orientation val=`"minMax`"/></c:scaling><c:delete val=`"0`"/><c:axPos val=`"b`"/><c:majorGridlines><c:spPr><a:ln w=`"9525`"><a:solidFill><a:srgbClr val=`"E6E6E6`"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:spPr><a:ln><a:noFill/></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz=`"850`"><a:solidFill><a:srgbClr val=`"333333`"/></a:solidFill><a:latin typeface=`"Meiryo UI`"/><a:ea typeface=`"Meiryo UI`"/></a:defRPr></a:pPr><a:endParaRPr lang=`"ja-JP`"/></a:p></c:txPr><c:crossAx val=`"1`"/>$majorUnitXml</c:valAx></c:plotArea></c:chart></c:chartSpace>")
        $zip.Dispose()
        $zip = $null
        if ($null -ne $LogCallback) { & $LogCallback "比較Excelファイル保存中: $OutputPath" 98 }
        if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }
        [AuroraAnalyzer.AtomicFile]::Publish($stagingPath, $OutputPath)
    }
    finally {
        try { if ($null -ne $zip) { $zip.Dispose() } }
        finally { try { if ([IO.File]::Exists($stagingPath)) { [IO.File]::Delete($stagingPath) } } catch { } }
    }
}

function Export-ComparisonReport {
    param(
        [Parameter(Mandatory = $true)]$Comparison,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [string]$Engine = 'Auto',
        [scriptblock]$LogCallback = $null,
        [AuroraAnalyzer.OperationCancelState]$CancelState = $null,
        [scriptblock]$PumpCallback = $null
    )

    if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }

    if ($Engine -eq 'OpenXml') {
        Export-ComparisonToOpenXml -Comparison $Comparison -OutputPath $OutputPath -LogCallback $LogCallback -CancelState $CancelState -PumpCallback $PumpCallback
        return 'OpenXml'
    }

    $comAvailable = $false
    if ($Engine -eq 'Com' -or $Engine -eq 'Auto') {
        try {
            $progId = [Type]::GetTypeFromProgID('Excel.Application')
            if ($null -ne $progId) { $comAvailable = $true }
        }
        catch { $comAvailable = $false }
    }

    if ($comAvailable) {
        try {
            Export-ComparisonToExcelCom -Comparison $Comparison -OutputPath $OutputPath -LogCallback $LogCallback -CancelState $CancelState -PumpCallback $PumpCallback
            return 'Com'
        }
        catch [AuroraAnalyzer.ExcelActivationException] {
            if ($Engine -eq 'Com') { throw }
        }
    }
    elseif ($Engine -eq 'Com') {
        throw "Microsoft Excel COM (Excel.Application) が利用できません。"
    }
    if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }
    if ($null -ne $LogCallback) { & $LogCallback "Excel COM を利用できないため OpenXML (.xlsx) フォールバックエンジンで比較レポートを出力します..." 45 }
    Export-ComparisonToOpenXml -Comparison $Comparison -OutputPath $OutputPath -LogCallback $LogCallback -CancelState $CancelState -PumpCallback $PumpCallback
    return 'OpenXml'
}

# ==============================================================================
#  Module 5: Core Pipeline Runner (Used by both GUI and CLI)
# ==============================================================================
function Invoke-AnalyzerPipeline {
    param(
        [Parameter(Mandatory = $true)][string]$RunMode,
        [hashtable]$ConnConfig = $null,
        [string]$DbLabel = '',
        [string]$CsvFolder = '',
        [string]$ExcelOutputPath = '',
        [int]$MaxRows = 0,
        [string[]]$IncTables = @(),
        [string[]]$ExcTables = @(),
        [string]$Engine = 'Auto',
        [scriptblock]$LogCallback = $null,
        [string]$BeforeAnalysisPath = '',
        [string]$AfterAnalysisPath = '',
        [switch]$Resume,
        [switch]$DetailedAnalysis,
        [AuroraAnalyzer.OperationCancelState]$CancelState = $null,
        [scriptblock]$PumpCallback = $null
    )

    function Log([string]$m, [int]$p = -1) {
        if ($null -ne $LogCallback) { & $LogCallback $m $p }
        else { Write-Host $m }
    }

    if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }

    if ($RunMode -eq 'CompareAnalysis') {
        if ([string]::IsNullOrWhiteSpace($BeforeAnalysisPath) -or [string]::IsNullOrWhiteSpace($AfterAnalysisPath)) {
            throw "テスト開始前およびテスト開始後の分析ファイル (.xlsx) を2ファイル指定してください。"
        }
        if ([string]::IsNullOrWhiteSpace($ExcelOutputPath)) {
            $stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
            $baseDir = Split-Path -Parent $AfterAnalysisPath
            if ([string]::IsNullOrWhiteSpace($baseDir)) { $baseDir = (Get-Location).Path }
            $ExcelOutputPath = Join-Path $baseDir "Aurora_DiffReport_${stamp}.xlsx"
        }

        Log "分析ファイル読込・比較開始: 開始前=$BeforeAnalysisPath / 開始後=$AfterAnalysisPath" 15
        $cmpHandle = [AuroraAnalyzer.BackgroundTaskLauncher]::StartCompareWorkbooks($BeforeAnalysisPath, $AfterAnalysisPath, $CancelState)
        $comparison = Wait-BackgroundTask -Handle $cmpHandle -LogCallback $LogCallback -PumpCallback $PumpCallback -CancelState $CancelState
        Log ("比較解析完了 ({0}): 全 {1} テーブル (変更 {2} [更新 {3} / 追加 {4} / 削除 {5}] / 変更なし {6}) | 全 {7} カラム (変更 {8} / 変更なし {9}) | 差分項目 {10} 件" -f `
            $comparison.PrecisionLabel, $comparison.TotalUniqueTableCount, $comparison.ChangedTableCount, `
            $comparison.UpdatedSheetCount, $comparison.AddedSheetCount, $comparison.DeletedTableCount, `
            $comparison.UnchangedTableCount, $comparison.TotalUniqueColumnCount, $comparison.ChangedColumnTotal, `
            $comparison.UnchangedColumnTotal, $comparison.AllColumnDiffs.Count) 40

        if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }
        $usedEngine = Export-ComparisonReport -Comparison $comparison -OutputPath $ExcelOutputPath -Engine $Engine -LogCallback $LogCallback -CancelState $CancelState -PumpCallback $PumpCallback
        Log ("比較レポート出力完了 (Engine=$usedEngine): $ExcelOutputPath") 100

        return [PSCustomObject]@{
            ExportedTables = @()
            Analysis       = $null
            Comparison     = $comparison
            OutputPath     = $ExcelOutputPath
            ExcelEngine    = $usedEngine
        }
    }

    if ($RunMode -eq 'DirectAnalyze') {
        if ($null -eq $ConnConfig) {
            throw "DB直接分析モードでは接続設定 (ConnConfig) を指定してください。"
        }
        if ([string]::IsNullOrWhiteSpace($ExcelOutputPath)) {
            $stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
            $ExcelOutputPath = Join-Path (Get-Location).Path "Aurora_PerfAnalysis_${stamp}.xlsx"
        }

        $analysis = Invoke-AuroraDirectAnalysis `
            -ConnConfig $ConnConfig `
            -DbLabel $DbLabel `
            -MaxRowsPerTable $MaxRows `
            -IncludeTables $IncTables `
            -ExcludeTables $ExcTables `
            -LogCallback $LogCallback `
            -CancelState $CancelState `
            -PumpCallback $PumpCallback

        if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }
        $usedEngine = Export-AnalysisReport -Analysis $analysis -OutputPath $ExcelOutputPath -DbLabel $DbLabel -Engine $Engine -DetailedAnalysis:$DetailedAnalysis -LogCallback $LogCallback -CancelState $CancelState -PumpCallback $PumpCallback
        Log ("完了 (DB直接分析 / Engine=$usedEngine): $ExcelOutputPath") 100

        return [PSCustomObject]@{
            ExportedTables = @()
            Analysis       = $analysis
            Comparison     = $null
            OutputPath     = $ExcelOutputPath
            ExcelEngine    = $usedEngine
        }
    }

    if ([string]::IsNullOrWhiteSpace($CsvFolder)) {
        throw "入力 / CSV保存フォルダを指定してください。"
    }

    $CsvFolder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($CsvFolder)
    $exported = @()
    $folderLock = $null
    $selectedFiles = $null
    $exportAnalysis = $null
    $analysis = $null
    $usedEngine = ''
    try {
        if ($RunMode -eq 'ExportAndAnalyze' -or $RunMode -eq 'ExportOnly') {
            $folderLock = Open-ExportFolderLock $CsvFolder
            $endPct = if ($RunMode -eq 'ExportOnly') { 100 } else { 45 }
            if ($RunMode -eq 'ExportAndAnalyze') { $exportAnalysis = [AuroraAnalyzer.CsvExportAnalysis]::new() }
            $exported = @(Invoke-AuroraBulkCsvExport `
                -ConnConfig $ConnConfig -ExportFolder $CsvFolder -FolderLock $folderLock `
                -MaxRowsPerTable $MaxRows -IncludeTables $IncTables -ExcludeTables $ExcTables `
                -LogCallback $LogCallback -Resume:$Resume -CancelState $CancelState -PumpCallback $PumpCallback -ExportEndPercent $endPct -ExportAnalysis $exportAnalysis)
            $selectedFiles = [string[]]$exported.CsvPath
        }
        if ($RunMode -ne 'ExportOnly') {
            if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }
            if ([string]::IsNullOrWhiteSpace($ExcelOutputPath)) {
                $stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
                $ExcelOutputPath = Join-Path $CsvFolder "Aurora_PerfAnalysis_${stamp}.xlsx"
            }
            Log "CSV解析開始: $CsvFolder" 45
            $anHandle = [AuroraAnalyzer.BackgroundTaskLauncher]::StartAnalyzeFolder($CsvFolder, $CancelState, $selectedFiles, $exportAnalysis)
            $analysis = Wait-BackgroundTask -Handle $anHandle -LogCallback $LogCallback -PumpCallback $PumpCallback -CancelState $CancelState
            Log ("CSV解析完了: {0} テーブル / 総レコード {1:N0} 件 / 検出警告 {2} 件 (高リスク {3} 件)" -f $analysis.TableCount, $analysis.TotalRows, $analysis.TotalWarningCount, $analysis.HighWarningCount) 58

            if ($null -ne $CancelState) { $CancelState.ThrowIfCancelled() }
            $usedEngine = Export-AnalysisReport -Analysis $analysis -OutputPath $ExcelOutputPath -DbLabel $DbLabel -Engine $Engine -DetailedAnalysis:$DetailedAnalysis -LogCallback $LogCallback -CancelState $CancelState -PumpCallback $PumpCallback
            if ($RunMode -eq 'ExportAndAnalyze') {
                $checkpoint = Read-ExportProgress $CsvFolder
                $checkpoint.Status = 'Completed'
                Save-ExportProgress $CsvFolder $checkpoint
            }
            Log ("完了 (Engine=$usedEngine): $ExcelOutputPath") 100
        }
    }
    finally { if ($null -ne $folderLock) { $folderLock.Dispose() } }

    return [PSCustomObject]@{
        ExportedTables = $exported
        Analysis       = $analysis
        Comparison     = $null
        OutputPath     = $ExcelOutputPath
        ExcelEngine    = $usedEngine
    }
}

# ==============================================================================
#  Module 6: Windows Forms GUI Dialog (Monotone + Midnight Navy + Meiryo UI)
# ==============================================================================
function New-AnalyzerGuiForm {
    param(
        [string]$InitialRegistryRoot = 'HKCU:\Software\mmatsubara',
        [string]$InitialInputFolder = '',
        [string]$InitialOutputPath = '',
        [string]$InitialA5DbName = '',
        [scriptblock]$ShowMessage = {
            param($owner, $text, $caption, $icon)
            [void][System.Windows.Forms.MessageBox]::Show($owner, $text, $caption, 'OK', $icon)
        },
        [string]$InitialA5DbListPath = '',
        [string]$InitialBeforeAnalysisPath = '',
        [string]$InitialAfterAnalysisPath = '',
        [string]$InitialMode = '',
        [switch]$InitialResume,
        [switch]$InitialDetailedAnalysis,
        [ValidateSet('Auto','Com','OpenXml')][string]$InitialExcelEngine = 'Auto'
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $cText   = ConvertTo-DrawingColor $script:THEME.TextHex
    $cMuted  = ConvertTo-DrawingColor $script:THEME.MutedHex
    $cNavy   = ConvertTo-DrawingColor $script:THEME.AccentHex
    $cWarn   = ConvertTo-DrawingColor $script:THEME.WarnHex
    $cBg     = ConvertTo-DrawingColor $script:THEME.ZebraHex
    $cWhite  = ConvertTo-DrawingColor $script:THEME.WhiteHex
    $cLight  = ConvertTo-DrawingColor $script:THEME.AccentLightHex

    $fontMain  = New-Object System.Drawing.Font('Meiryo UI', 9.0)
    $fontBold  = New-Object System.Drawing.Font('Meiryo UI', 9.0, [System.Drawing.FontStyle]::Bold)
    $fontTitle = New-Object System.Drawing.Font('Meiryo UI', 12.0, [System.Drawing.FontStyle]::Bold)
    $fontSmall = New-Object System.Drawing.Font('Meiryo UI', 8.0)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Aurora Serverless v2 (MySQL 8.4.7) A5:SQL Mk-2 CSV一括出力・性能試験データ分析ツール'
    $form.AutoScaleMode = 'Dpi'
    $form.Size = New-Object System.Drawing.Size(820, 788)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.BackColor = $cBg
    $form.ForeColor = $cText
    $form.Font = $fontMain

    # Header Banner
    $pnlHeader = New-Object System.Windows.Forms.Panel
    $pnlHeader.Dock = 'Top'
    $pnlHeader.Height = 58
    $pnlHeader.BackColor = $cNavy
    $form.Controls.Add($pnlHeader)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = 'Aurora Serverless v2 (MySQL 8.4.7) 性能試験データ分析ツール'
    $lblTitle.Font = $fontTitle
    $lblTitle.ForeColor = $cWhite
    $lblTitle.Location = New-Object System.Drawing.Point(16, 8)
    $lblTitle.AutoSize = $true
    $pnlHeader.Controls.Add($lblTitle)

    $lblSub = New-Object System.Windows.Forms.Label
    $lblSub.Text = 'A5:SQL Mk-2 登録DB直接分析・CSV一括出力・性能試験データ分析 ＆ テスト開始前/開始後 比較Excel出力'
    $lblSub.Font = $fontSmall
    $lblSub.ForeColor = $cLight
    $lblSub.Location = New-Object System.Drawing.Point(18, 34)
    $lblSub.AutoSize = $true
    $pnlHeader.Controls.Add($lblSub)

    # Mode GroupBox (3x2 clean layout: DB-connected modes on left, local file modes on right)
    $gbMode = New-Object System.Windows.Forms.GroupBox
    $gbMode.Text = '1. 実行モード選択'
    $gbMode.Font = $fontBold
    $gbMode.ForeColor = $cNavy
    $gbMode.Location = New-Object System.Drawing.Point(14, 64)
    $gbMode.Size = New-Object System.Drawing.Size(776, 102)
    $gbMode.BackColor = $cWhite
    $form.Controls.Add($gbMode)

    $rbDirect = New-Object System.Windows.Forms.RadioButton
    $rbDirect.Name = 'DirectAnalyzeMode'
    $rbDirect.Text = '登録DBから直接分析Excel生成 (CSV出力なし)'
    $rbDirect.Font = $fontMain; $rbDirect.ForeColor = $cText
    $rbDirect.Location = New-Object System.Drawing.Point(14, 22)
    $rbDirect.AutoSize = $true
    $gbMode.Controls.Add($rbDirect)

    $rbAnalyze = New-Object System.Windows.Forms.RadioButton
    $rbAnalyze.Text = '既存CSVフォルダを分析してExcel生成'
    $rbAnalyze.Font = $fontMain; $rbAnalyze.ForeColor = $cText
    $rbAnalyze.Location = New-Object System.Drawing.Point(395, 22)
    $rbAnalyze.AutoSize = $true
    $rbAnalyze.Checked = $true
    $gbMode.Controls.Add($rbAnalyze)

    $rbFull = New-Object System.Windows.Forms.RadioButton
    $rbFull.Text = 'A5M2登録DBからCSV一括出力 ＋ 分析Excel生成'
    $rbFull.Font = $fontMain; $rbFull.ForeColor = $cText
    $rbFull.Location = New-Object System.Drawing.Point(14, 48)
    $rbFull.AutoSize = $true
    $gbMode.Controls.Add($rbFull)

    $rbCompare = New-Object System.Windows.Forms.RadioButton
    $rbCompare.Name = 'CompareMode'
    $rbCompare.Text = 'テスト前後の分析ファイル2件を比較してExcel生成'
    $rbCompare.Font = $fontMain; $rbCompare.ForeColor = $cText
    $rbCompare.Location = New-Object System.Drawing.Point(395, 48)
    $rbCompare.AutoSize = $true
    $gbMode.Controls.Add($rbCompare)

    $rbExportOnly = New-Object System.Windows.Forms.RadioButton
    $rbExportOnly.Text = 'CSV一括出力のみ'
    $rbExportOnly.Font = $fontMain; $rbExportOnly.ForeColor = $cText
    $rbExportOnly.Location = New-Object System.Drawing.Point(14, 74)
    $rbExportOnly.AutoSize = $true
    $gbMode.Controls.Add($rbExportOnly)

    # Folder / Output GroupBox (Prominent & always active for input/output!)
    $gbPath = New-Object System.Windows.Forms.GroupBox
    $gbPath.Text = '2. 入力フォルダ (CSVフォルダ) ＆ 出力先Excelファイル指定'
    $gbPath.Font = $fontBold
    $gbPath.ForeColor = $cNavy
    $gbPath.Location = New-Object System.Drawing.Point(14, 172)
    $gbPath.Size = New-Object System.Drawing.Size(776, 94)
    $gbPath.BackColor = $cWhite
    $form.Controls.Add($gbPath)

    $lblIn = New-Object System.Windows.Forms.Label
    $lblIn.Text = '入力 / CSV保存フォルダ:'
    $lblIn.Font = $fontMain; $lblIn.ForeColor = $cText
    $lblIn.Location = New-Object System.Drawing.Point(14, 26)
    $lblIn.Size = New-Object System.Drawing.Size(155, 22)
    $gbPath.Controls.Add($lblIn)

    $txtInFolder = New-Object System.Windows.Forms.TextBox
    $txtInFolder.Font = $fontMain; $txtInFolder.ForeColor = $cText
    $txtInFolder.Location = New-Object System.Drawing.Point(174, 23)
    $txtInFolder.Size = New-Object System.Drawing.Size(495, 24)
    $txtInFolder.Text = $InitialInputFolder
    $gbPath.Controls.Add($txtInFolder)

    $btnBrowseIn = New-Object System.Windows.Forms.Button
    $btnBrowseIn.Text = 'フォルダ選択...'
    $btnBrowseIn.Font = $fontMain; $btnBrowseIn.ForeColor = $cText; $btnBrowseIn.BackColor = $cLight
    $btnBrowseIn.FlatStyle = 'Flat'
    $btnBrowseIn.Location = New-Object System.Drawing.Point(676, 22)
    $btnBrowseIn.Size = New-Object System.Drawing.Size(88, 26)
    $gbPath.Controls.Add($btnBrowseIn)

    $lblOut = New-Object System.Windows.Forms.Label
    $lblOut.Text = '出力先Excel (.xlsx):'
    $lblOut.Font = $fontMain; $lblOut.ForeColor = $cText
    $lblOut.Location = New-Object System.Drawing.Point(14, 59)
    $lblOut.Size = New-Object System.Drawing.Size(155, 22)
    $gbPath.Controls.Add($lblOut)

    $txtOutFile = New-Object System.Windows.Forms.TextBox
    $txtOutFile.Font = $fontMain; $txtOutFile.ForeColor = $cText
    $txtOutFile.Location = New-Object System.Drawing.Point(174, 56)
    $txtOutFile.Size = New-Object System.Drawing.Size(495, 24)
    $txtOutFile.Text = $InitialOutputPath
    $gbPath.Controls.Add($txtOutFile)

    $btnBrowseOut = New-Object System.Windows.Forms.Button
    $btnBrowseOut.Text = '保存先指定...'
    $btnBrowseOut.Font = $fontMain; $btnBrowseOut.ForeColor = $cText; $btnBrowseOut.BackColor = $cLight
    $btnBrowseOut.FlatStyle = 'Flat'
    $btnBrowseOut.Location = New-Object System.Drawing.Point(676, 55)
    $btnBrowseOut.Size = New-Object System.Drawing.Size(88, 26)
    $gbPath.Controls.Add($btnBrowseOut)

    # A5:SQL Mk-2 DB & SSH Tunnel GroupBox
    $gbA5 = New-Object System.Windows.Forms.GroupBox
    $gbA5.Text = '3. A5:SQL Mk-2 登録データベース接続設定 (EC2踏み台 SSHトンネル / Aurora MySQL 8.4.7)'
    $gbA5.Font = $fontBold
    $gbA5.ForeColor = $cNavy
    $gbA5.Location = New-Object System.Drawing.Point(14, 272)
    $gbA5.Size = New-Object System.Drawing.Size(776, 276)
    $gbA5.BackColor = $cWhite
    $form.Controls.Add($gbA5)

    # Compare Analysis Files GroupBox (swapped in when CompareMode is selected)
    $gbCompare = New-Object System.Windows.Forms.GroupBox
    $gbCompare.Name = 'CompareGroup'
    $gbCompare.Text = '3. 比較対象 分析ファイル指定 (テスト開始前 / テスト開始後 .xlsx)'
    $gbCompare.Font = $fontBold
    $gbCompare.ForeColor = $cNavy
    $gbCompare.Location = New-Object System.Drawing.Point(14, 272)
    $gbCompare.Size = New-Object System.Drawing.Size(776, 276)
    $gbCompare.BackColor = $cWhite
    $gbCompare.Visible = $false
    $form.Controls.Add($gbCompare)

    $lblBefore = New-Object System.Windows.Forms.Label
    $lblBefore.Text = 'テスト開始前 分析ファイル:'
    $lblBefore.Font = $fontMain; $lblBefore.ForeColor = $cText
    $lblBefore.Location = New-Object System.Drawing.Point(14, 34)
    $lblBefore.Size = New-Object System.Drawing.Size(165, 22)
    $gbCompare.Controls.Add($lblBefore)

    $txtBeforeFile = New-Object System.Windows.Forms.TextBox
    $txtBeforeFile.Name = 'BeforeAnalysisPath'
    $txtBeforeFile.Font = $fontMain; $txtBeforeFile.ForeColor = $cText
    $txtBeforeFile.Location = New-Object System.Drawing.Point(182, 31)
    $txtBeforeFile.Size = New-Object System.Drawing.Size(487, 24)
    $txtBeforeFile.Text = $InitialBeforeAnalysisPath
    $gbCompare.Controls.Add($txtBeforeFile)

    $btnBrowseBefore = New-Object System.Windows.Forms.Button
    $btnBrowseBefore.Name = 'BrowseBefore'
    $btnBrowseBefore.Text = 'ファイル選択...'
    $btnBrowseBefore.Font = $fontMain; $btnBrowseBefore.ForeColor = $cText; $btnBrowseBefore.BackColor = $cLight
    $btnBrowseBefore.FlatStyle = 'Flat'
    $btnBrowseBefore.Location = New-Object System.Drawing.Point(676, 30)
    $btnBrowseBefore.Size = New-Object System.Drawing.Size(88, 26)
    $gbCompare.Controls.Add($btnBrowseBefore)

    $lblAfter = New-Object System.Windows.Forms.Label
    $lblAfter.Text = 'テスト開始後 分析ファイル:'
    $lblAfter.Font = $fontMain; $lblAfter.ForeColor = $cText
    $lblAfter.Location = New-Object System.Drawing.Point(14, 72)
    $lblAfter.Size = New-Object System.Drawing.Size(165, 22)
    $gbCompare.Controls.Add($lblAfter)

    $txtAfterFile = New-Object System.Windows.Forms.TextBox
    $txtAfterFile.Name = 'AfterAnalysisPath'
    $txtAfterFile.Font = $fontMain; $txtAfterFile.ForeColor = $cText
    $txtAfterFile.Location = New-Object System.Drawing.Point(182, 69)
    $txtAfterFile.Size = New-Object System.Drawing.Size(487, 24)
    $txtAfterFile.Text = $InitialAfterAnalysisPath
    $gbCompare.Controls.Add($txtAfterFile)

    $btnBrowseAfter = New-Object System.Windows.Forms.Button
    $btnBrowseAfter.Name = 'BrowseAfter'
    $btnBrowseAfter.Text = 'ファイル選択...'
    $btnBrowseAfter.Font = $fontMain; $btnBrowseAfter.ForeColor = $cText; $btnBrowseAfter.BackColor = $cLight
    $btnBrowseAfter.FlatStyle = 'Flat'
    $btnBrowseAfter.Location = New-Object System.Drawing.Point(676, 68)
    $btnBrowseAfter.Size = New-Object System.Drawing.Size(88, 26)
    $gbCompare.Controls.Add($btnBrowseAfter)

    $btnSwapFiles = New-Object System.Windows.Forms.Button
    $btnSwapFiles.Name = 'SwapCompareFiles'
    $btnSwapFiles.Text = '開始前 / 開始後を入替'
    $btnSwapFiles.Font = $fontSmall; $btnSwapFiles.ForeColor = $cNavy; $btnSwapFiles.BackColor = $cLight
    $btnSwapFiles.FlatStyle = 'Flat'
    $btnSwapFiles.Location = New-Object System.Drawing.Point(182, 104)
    $btnSwapFiles.Size = New-Object System.Drawing.Size(150, 25)
    $gbCompare.Controls.Add($btnSwapFiles)

    $lblCmpInfo = New-Object System.Windows.Forms.Label
    $lblCmpInfo.Text = "【出力される比較Excelレポート構成 (8シート / CSV未出力のDB直接分析レポート完全対応・CustomXMLハッシュ照合)】`r`n・01_比較サマリー: 12種KPIカード(件数速度・行変動・警告・FK・DB状態)、変更区分サマリー、推移比較グラフ、総合所見`r`n・02_全テーブル比較一覧 / 03_更新 / 04_追加 / 05_変更なし: 増減率・増減速度・行変動内訳(挿入/更新/削除)・変更/不変列数・ハッシュ`r`n・06_カラム・統計差分詳細: ハッシュ・DDL・INDEX・FK・物理容量・AUTO_INCREMENT・数値合計/パーセンタイル・分布・日時差分`r`n・07_全カラム比較一覧: 全テーブルの全カラムを変更あり/変更なしで一覧化 (フィルタ抽出・値ハッシュ照合根拠付き)`r`n・08_性能・DB・時系列・品質変化: 警告推移、FK孤立行/ファンアウト、時間帯増分、MySQLサーバー負荷・物理InnoDB・JMeter適性"
    $lblCmpInfo.Font = $fontSmall; $lblCmpInfo.ForeColor = $cMuted
    $lblCmpInfo.BackColor = $cBg
    $lblCmpInfo.BorderStyle = 'FixedSingle'
    $lblCmpInfo.Padding = New-Object System.Windows.Forms.Padding(8, 6, 8, 6)
    $lblCmpInfo.Location = New-Object System.Drawing.Point(14, 140)
    $lblCmpInfo.Size = New-Object System.Drawing.Size(750, 124)
    $gbCompare.Controls.Add($lblCmpInfo)

    function Add-Lbl([System.Windows.Forms.Control]$Parent, [string]$Text, [int]$X, [int]$Y, [int]$W) {
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $Text; $l.Font = $fontMain; $l.ForeColor = $cText
        $l.Location = New-Object System.Drawing.Point($X, $Y)
        $l.Size = New-Object System.Drawing.Size($W, 22)
        $Parent.Controls.Add($l)
        return $l
    }

    function Add-Txt([System.Windows.Forms.Control]$Parent, [int]$X, [int]$Y, [int]$W, [bool]$Password = $false) {
        $t = New-Object System.Windows.Forms.TextBox
        $t.Font = $fontMain; $t.ForeColor = $cText
        $t.Location = New-Object System.Drawing.Point($X, $Y)
        $t.Size = New-Object System.Drawing.Size($W, 24)
        if ($Password) { $t.UseSystemPasswordChar = $true }
        $Parent.Controls.Add($t)
        return $t
    }

    [void](Add-Lbl $gbA5 '登録DBの読込元:' 14 26 115)
    $cmbA5Source = New-Object System.Windows.Forms.ComboBox
    $cmbA5Source.Name = 'A5Source'
    $cmbA5Source.Font = $fontMain
    $cmbA5Source.DropDownStyle = 'DropDownList'
    $cmbA5Source.Location = New-Object System.Drawing.Point(132, 23)
    $cmbA5Source.Size = New-Object System.Drawing.Size(467, 24)
    [void]$cmbA5Source.Items.Add('レジストリ')
    $cmbA5Source.SelectedIndex = 0
    if ($InitialA5DbListPath) { $cmbA5Source.SelectedIndex = $cmbA5Source.Items.Add($InitialA5DbListPath) }
    $gbA5.Controls.Add($cmbA5Source)

    $btnBrowseA5List = New-Object System.Windows.Forms.Button
    $btnBrowseA5List.Text = 'ポータブルDBリスト選択...'
    $btnBrowseA5List.Font = $fontMain; $btnBrowseA5List.BackColor = $cLight
    $btnBrowseA5List.FlatStyle = 'Flat'
    $btnBrowseA5List.Location = New-Object System.Drawing.Point(605, 22)
    $btnBrowseA5List.Size = New-Object System.Drawing.Size(159, 26)
    $gbA5.Controls.Add($btnBrowseA5List)

    [void](Add-Lbl $gbA5 'A5M2 登録DB名:' 14 56 115)
    $cmbA5Db = New-Object System.Windows.Forms.ComboBox
    $cmbA5Db.Name = 'A5DbName'
    $cmbA5Db.Font = $fontMain; $cmbA5Db.ForeColor = $cText
    $cmbA5Db.DropDownStyle = 'DropDown'
    $cmbA5Db.Location = New-Object System.Drawing.Point(132, 53)
    $cmbA5Db.Size = New-Object System.Drawing.Size(535, 24)
    $gbA5.Controls.Add($cmbA5Db)

    $btnRefreshA5 = New-Object System.Windows.Forms.Button
    $btnRefreshA5.Text = '再読込'
    $btnRefreshA5.Font = $fontMain; $btnRefreshA5.ForeColor = $cText; $btnRefreshA5.BackColor = $cLight
    $btnRefreshA5.FlatStyle = 'Flat'
    $btnRefreshA5.Location = New-Object System.Drawing.Point(676, 52)
    $btnRefreshA5.Size = New-Object System.Drawing.Size(88, 26)
    $gbA5.Controls.Add($btnRefreshA5)

    # Row 2: Aurora Host / Port / DB
    [void](Add-Lbl $gbA5 'Auroraホスト:' 14 88 115)
    $txtDbHost = Add-Txt $gbA5 132 85 290
    [void](Add-Lbl $gbA5 'Port:' 430 88 38)
    $txtDbPort = Add-Txt $gbA5 470 85 60; $txtDbPort.Text = '3306'
    [void](Add-Lbl $gbA5 'スキーマ:' 538 88 64)
    $txtDbName = Add-Txt $gbA5 604 85 160

    # Row 3: DB User / Password / SSL Mode
    [void](Add-Lbl $gbA5 'DBユーザー:' 14 120 115)
    $txtDbUser = Add-Txt $gbA5 132 117 170
    [void](Add-Lbl $gbA5 'DBパスワード:' 310 120 95)
    $txtDbPass = Add-Txt $gbA5 408 117 165 $true
    [void](Add-Lbl $gbA5 'MySQL SSL:' 582 120 78)
    $cmbSsl = New-Object System.Windows.Forms.ComboBox
    $cmbSsl.Font = $fontMain; $cmbSsl.ForeColor = $cText; $cmbSsl.DropDownStyle = 'DropDownList'
    [void]$cmbSsl.Items.AddRange(@('Preferred', 'Required', 'VerifyCA', 'Disabled'))
    $cmbSsl.SelectedIndex = 0
    $cmbSsl.Location = New-Object System.Drawing.Point(664, 117)
    $cmbSsl.Size = New-Object System.Drawing.Size(100, 24)
    $gbA5.Controls.Add($cmbSsl)

    # Row 4: SSH Tunnel Toggle + EC2 Host + Port + User
    $chkSsh = New-Object System.Windows.Forms.CheckBox
    $chkSsh.Text = 'SSHトンネル経由 (EC2踏み台)'
    $chkSsh.Font = $fontBold; $chkSsh.ForeColor = $cNavy
    $chkSsh.Location = New-Object System.Drawing.Point(14, 152)
    $chkSsh.AutoSize = $true
    $chkSsh.Checked = $true
    $gbA5.Controls.Add($chkSsh)

    [void](Add-Lbl $gbA5 '上限件数/表(0=全件):' 475 154 140)
    $txtMaxRows = Add-Txt $gbA5 620 151 144; $txtMaxRows.Text = '0'

    [void](Add-Lbl $gbA5 'EC2踏み台ホスト:' 14 182 115)
    $txtSshHost = Add-Txt $gbA5 132 179 270
    [void](Add-Lbl $gbA5 'SSH Port:' 410 182 65)
    $txtSshPort = Add-Txt $gbA5 478 179 55; $txtSshPort.Text = '22'
    [void](Add-Lbl $gbA5 'SSHユーザー:' 540 182 88)
    $txtSshUser = Add-Txt $gbA5 630 179 134

    # Row 5: SSH Key + Passphrase
    [void](Add-Lbl $gbA5 'SSH秘密鍵(.pem):' 14 214 115)
    $txtSshKey = Add-Txt $gbA5 132 211 310
    $btnBrowseKey = New-Object System.Windows.Forms.Button
    $btnBrowseKey.Text = '鍵選択...'
    $btnBrowseKey.Font = $fontSmall; $btnBrowseKey.ForeColor = $cText; $btnBrowseKey.BackColor = $cLight
    $btnBrowseKey.FlatStyle = 'Flat'
    $btnBrowseKey.Location = New-Object System.Drawing.Point(446, 210)
    $btnBrowseKey.Size = New-Object System.Drawing.Size(68, 25)
    $gbA5.Controls.Add($btnBrowseKey)

    [void](Add-Lbl $gbA5 '鍵パスフレーズ:' 522 214 98)
    $txtSshPass = Add-Txt $gbA5 622 211 142 $true

    $lblA5Hint = New-Object System.Windows.Forms.Label
    $lblA5Hint.Text = '※ A5:SQL Mk-2に保存されたパスワードは暗号化されているため、DB接続実行時はDBパスワード(および必要に応じ鍵パスフレーズ)をご入力ください。'
    $lblA5Hint.Font = $fontSmall; $lblA5Hint.ForeColor = $cMuted
    $lblA5Hint.Location = New-Object System.Drawing.Point(14, 242)
    $lblA5Hint.Size = New-Object System.Drawing.Size(750, 28)
    $gbA5.Controls.Add($lblA5Hint)

    # Progress & Log Area
    $pb = New-Object System.Windows.Forms.ProgressBar
    $pb.Location = New-Object System.Drawing.Point(14, 554)
    $pb.Size = New-Object System.Drawing.Size(600, 18)
    $pb.Style = 'Continuous'
    $form.Controls.Add($pb)
    $lblProgress = Add-Lbl $form '全体 0%' 618 552 172
    $lblProgress.Name = 'ProgressPercent'
    $lblProgress.TextAlign = 'MiddleRight'

    $txtLog = New-Object System.Windows.Forms.TextBox
    $txtLog.Multiline = $true
    $txtLog.ReadOnly = $true
    $txtLog.ScrollBars = 'Vertical'
    $txtLog.Font = New-Object System.Drawing.Font('Meiryo UI', 8.5)
    $txtLog.BackColor = $cWhite
    $txtLog.ForeColor = $cText
    $txtLog.Location = New-Object System.Drawing.Point(14, 578)
    $txtLog.Size = New-Object System.Drawing.Size(776, 80)
    $form.Controls.Add($txtLog)

    # Bottom Controls
    $chkDetailedAnalysis = New-Object System.Windows.Forms.CheckBox
    $chkDetailedAnalysis.Name = 'DetailedAnalysis'
    $chkDetailedAnalysis.Text = '詳細分析を出力する (分布・数値統計・時系列・関連・警告シートを追加)'
    $chkDetailedAnalysis.Font = $fontMain; $chkDetailedAnalysis.ForeColor = $cText
    $chkDetailedAnalysis.Location = New-Object System.Drawing.Point(16, 666)
    $chkDetailedAnalysis.AutoSize = $true
    $chkDetailedAnalysis.Checked = [bool]$InitialDetailedAnalysis
    $form.Controls.Add($chkDetailedAnalysis)

    $chkOpenExcel = New-Object System.Windows.Forms.CheckBox
    $chkOpenExcel.Text = '完了後にExcelレポートを自動で開く'
    $chkOpenExcel.Font = $fontMain; $chkOpenExcel.ForeColor = $cText
    $chkOpenExcel.Location = New-Object System.Drawing.Point(16, 688)
    $chkOpenExcel.AutoSize = $true
    $chkOpenExcel.Checked = $true
    $form.Controls.Add($chkOpenExcel)

    $chkResume = New-Object System.Windows.Forms.CheckBox
    $chkResume.Name = 'ResumeExport'
    $chkResume.Text = '出力済みCSVから続行 (中断後の再開・出力済みスキップ)'
    $chkResume.Font = $fontMain; $chkResume.ForeColor = $cText
    $chkResume.Location = New-Object System.Drawing.Point(16, 710)
    $chkResume.AutoSize = $true
    $chkResume.Checked = [bool]$InitialResume
    $form.Controls.Add($chkResume)

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = '実行開始'
    $btnRun.Font = $fontBold
    $btnRun.BackColor = $cNavy
    $btnRun.ForeColor = $cWhite
    $btnRun.FlatStyle = 'Flat'
    $btnRun.Location = New-Object System.Drawing.Point(466, 686)
    $btnRun.Size = New-Object System.Drawing.Size(112, 36)
    $form.Controls.Add($btnRun)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Name = 'CancelButton'
    $btnCancel.Text = '中断'
    $btnCancel.Font = $fontBold
    $btnCancel.BackColor = $cLight
    $btnCancel.ForeColor = $cMuted
    $btnCancel.FlatStyle = 'Flat'
    $btnCancel.Location = New-Object System.Drawing.Point(586, 686)
    $btnCancel.Size = New-Object System.Drawing.Size(102, 36)
    $btnCancel.Enabled = $false
    $btnCancel.Cursor = [System.Windows.Forms.Cursors]::Default
    $form.Controls.Add($btnCancel)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = '閉じる'
    $btnClose.Font = $fontMain
    $btnClose.BackColor = $cLight
    $btnClose.ForeColor = $cText
    $btnClose.FlatStyle = 'Flat'
    $btnClose.Location = New-Object System.Drawing.Point(696, 686)
    $btnClose.Size = New-Object System.Drawing.Size(94, 36)
    $btnClose.Cursor = [System.Windows.Forms.Cursors]::Default
    $btnClose.Add_Click({ $form.Close() }.GetNewClosure())
    $form.Controls.Add($btnClose)

    # Each closure keeps this form's controls after the factory returns. Capture the
    # script functions too, since GetNewClosure() runs callbacks in a new module scope.
    $getRegisteredDatabases = ${function:Get-A5RegisteredDatabases}
    $runPipeline = ${function:Invoke-AnalyzerPipeline}
    $getResumeSummary = ${function:Get-ExportFolderResumeSummary}
    $guiA5Map = @{}
    $guiRunState = @{
        IsRunning        = $false
        CancelState      = $null
        CloseAfterCancel = $false
        InspectedFolder  = $null
    }

    $pumpUi = {
        [System.Windows.Forms.Application]::DoEvents()
    }.GetNewClosure()

    $appendLog = {
        param([string]$msg, [int]$pct = -1, [int]$csvPct = -1)
        $ts = (Get-Date).ToString('HH:mm:ss')
        $txtLog.AppendText("[$ts] $msg`r`n")
        if ($pct -ge 0 -and $pct -le 100) {
            $pb.Value = $pct
            $lblProgress.Text = if ($csvPct -ge 0) { "全体 $pct%   CSV $csvPct%" } else { "全体 $pct%" }
        }
        [System.Windows.Forms.Application]::DoEvents()
    }.GetNewClosure()

    $inspectResumeFolder = {
        $dir = $txtInFolder.Text.Trim()
        if ($guiRunState.IsRunning -or -not ($rbFull.Checked -or $rbExportOnly.Checked) -or $guiRunState.InspectedFolder -eq $dir) { return }
        if ($null -ne $guiRunState.InspectedFolder) { $chkResume.Checked = $false }
        $guiRunState.InspectedFolder = $dir
        if ([string]::IsNullOrWhiteSpace($dir)) { return }
        $sum = & $getResumeSummary $dir
        if ($null -eq $sum) { return }
        if ($sum.HasProgressFile -and $sum.ProgressStatus -in 'Interrupted','InProgress','CsvCompleted') {
            $chkResume.Checked = $true
            & $appendLog ("前回の出力は {0}/{1} テーブル完了しています。「出力済みCSVから続行」を選択しました。" -f $sum.RecordedDone, $sum.TotalTables) -1
        }
        elseif ($sum.ProgressStatus -eq 'Invalid') {
            & $appendLog '進捗記録を読み込めません。このフォルダからの続行には正常な進捗記録が必要です。' -1
        }
        elseif ($sum.ExistingCsvCount -gt 0 -and ($rbFull.Checked -or $rbExportOnly.Checked)) {
            & $appendLog ("指定フォルダに既存のCSVファイルを {0} 件検出しました。「出力済みCSVから続行」にチェックを入れると出力済みテーブルをスキップして続行できます。" -f $sum.ExistingCsvCount) -1
        }
    }.GetNewClosure()
    $txtInFolder.Add_Leave($inspectResumeFolder)

    $btnCancel.Add_Click({
        if ($guiRunState.IsRunning -and $null -ne $guiRunState.CancelState -and -not $guiRunState.CancelState.IsCancelled) {
            $btnCancel.Enabled = $false
            $btnCancel.BackColor = $cLight
            $btnCancel.ForeColor = $cMuted
            $guiRunState.CancelState.Cancel()
            & $appendLog '中断要求を受け付けました。接続・処理を安全に停止しています...' -1
        }
    }.GetNewClosure())

    $form.Add_FormClosing({
        param($sender, $e)
        if ($guiRunState.IsRunning) {
            $e.Cancel = $true
            $guiRunState.CloseAfterCancel = $true
            if ($null -ne $guiRunState.CancelState -and -not $guiRunState.CancelState.IsCancelled) {
                $btnCancel.Enabled = $false
                $btnCancel.BackColor = $cLight
                $btnCancel.ForeColor = $cMuted
                $btnClose.Enabled = $false
                $guiRunState.CancelState.Cancel()
                & $appendLog '終了要求を受け付けました。現在の処理を安全に中断して終了します...' -1
            }
        }
    }.GetNewClosure())

    $refreshA5List = {
        $cmbA5Db.Tag = $true
        $previousName = [string]$cmbA5Db.SelectedItem
        $cmbA5Db.Items.Clear()
        $cmbA5Db.Text = ''
        $guiA5Map.Clear()
        # A failed reload or an empty list must not leave the previous connection usable.
        foreach ($field in $txtDbHost, $txtDbName, $txtDbUser, $txtDbPass, $txtSshHost, $txtSshUser, $txtSshKey, $txtSshPass) {
            $field.Clear()
        }
        $txtDbPort.Text = '3306'
        $txtSshPort.Text = '22'
        $chkSsh.Checked = $true
        $cmbSsl.SelectedIndex = 0
        $cmbSsl.Tag = $null

        $listPath = if ($cmbA5Source.SelectedIndex -gt 0) { [string]$cmbA5Source.SelectedItem } else { '' }
        try {
            $dbs = @(& $getRegisteredDatabases -RegistryRoot $InitialRegistryRoot -DbListPath $listPath)
        }
        catch {
            & $appendLog ("登録DBの読込に失敗しました: {0}" -f $_.Exception.Message) -1
            return
        }
        foreach ($d in $dbs) {
            $guiA5Map[$d.RegisteredName] = $d
            [void]$cmbA5Db.Items.Add($d.RegisteredName)
        }
        if ($cmbA5Db.Items.Count -gt 0) {
            if ($guiA5Map.ContainsKey($previousName)) {
                $cmbA5Db.SelectedItem = $guiA5Map[$previousName].RegisteredName
            }
            elseif ($InitialA5DbName -and $guiA5Map.ContainsKey($InitialA5DbName)) {
                $cmbA5Db.SelectedItem = $guiA5Map[$InitialA5DbName].RegisteredName
            }
            elseif ($listPath -and $InitialA5DbName) {
                & $appendLog "指定された登録DB名 '$InitialA5DbName' がこのDBリストにありません。登録DB名を選択してください。" -1
            }
            else {
                # Prefer MySQL entry if present
                $mysqlEntry = $dbs | Where-Object { $_.ProviderName -match 'MySQL' } | Select-Object -First 1
                if ($null -ne $mysqlEntry) { $cmbA5Db.SelectedItem = $mysqlEntry.RegisteredName }
                else { $cmbA5Db.SelectedIndex = 0 }
            }
            & $appendLog ("A5:SQL Mk-2 登録データベースを {0} 件検出しました。" -f $dbs.Count) -1
        }
        else {
            & $appendLog "選択した読込元に登録DBがありません。ポータブルDBリストの選択、手動入力、既存CSV分析が利用できます。" -1
        }
    }.GetNewClosure()

    $cmbA5Db.Add_SelectedIndexChanged({
        $sel = [string]$cmbA5Db.SelectedItem
        if (-not [string]::IsNullOrWhiteSpace($sel) -and $guiA5Map.ContainsKey($sel)) {
            $d = $guiA5Map[$sel]
            $txtDbHost.Text  = $d.ServerName
            $txtDbPort.Text  = [string]$d.Port
            $txtDbName.Text  = $d.Database
            $txtDbUser.Text  = $d.UserName
            $txtDbPass.Text  = $d.PlainPassword
            $chkSsh.Checked  = [bool]$d.UseSshTunnel
            $txtSshHost.Text = $d.SshHost
            $txtSshPort.Text = [string]$d.SshPort
            $txtSshUser.Text = $d.SshUser
            $txtSshKey.Text  = $d.SshKeyFile
            $txtSshPass.Clear()
            $cmbSsl.Tag = $d.SslCaCert
            $idx = $cmbSsl.FindStringExact($d.SslMode)
            if ($idx -ge 0) { $cmbSsl.SelectedIndex = $idx }
            & $appendLog ("A5M2設定読込: [{0}] Provider={1}, Host={2}:{3}, DB={4}, SSH={5}" -f $d.RegisteredName, $d.ProviderName, $d.ServerName, $d.Port, $d.Database, $d.UseSshTunnel) -1
        }
    }.GetNewClosure())

    $updateEnableState = {
        $isCompare = $rbCompare.Checked
        $isDirect  = $rbDirect.Checked
        $needDb = $rbFull.Checked -or $rbExportOnly.Checked -or $isDirect
        $needOut = $rbFull.Checked -or $rbAnalyze.Checked -or $isCompare -or $isDirect
        $needInFolder = $rbFull.Checked -or $rbAnalyze.Checked -or $rbExportOnly.Checked
        $needResume = $rbFull.Checked -or $rbExportOnly.Checked
        $needAnalysisOut = $rbFull.Checked -or $rbAnalyze.Checked -or $isDirect

        $gbA5.Visible = -not $isCompare
        $gbA5.Enabled = $needDb -and (-not $guiRunState.IsRunning)
        $gbCompare.Visible = $isCompare
        $gbCompare.Enabled = $isCompare -and (-not $guiRunState.IsRunning)

        $txtInFolder.Enabled = $needInFolder
        $btnBrowseIn.Enabled = $needInFolder
        $txtOutFile.Enabled = $needOut
        $btnBrowseOut.Enabled = $needOut
        $chkDetailedAnalysis.Enabled = $needAnalysisOut -and (-not $guiRunState.IsRunning)
        $chkOpenExcel.Enabled = $needOut -and (-not $guiRunState.IsRunning)
        $chkResume.Enabled = $needResume -and (-not $guiRunState.IsRunning)
        if ($needResume -and -not [string]::IsNullOrWhiteSpace($txtInFolder.Text)) { & $inspectResumeFolder }
        $txtSshHost.Enabled = $chkSsh.Checked
        $txtSshPort.Enabled = $chkSsh.Checked
        $txtSshUser.Enabled = $chkSsh.Checked
        $txtSshKey.Enabled  = $chkSsh.Checked
        $btnBrowseKey.Enabled = $chkSsh.Checked
        $txtSshPass.Enabled = $chkSsh.Checked
    }.GetNewClosure()

    $rbDirect.Add_CheckedChanged($updateEnableState)
    $rbFull.Add_CheckedChanged($updateEnableState)
    $rbAnalyze.Add_CheckedChanged($updateEnableState)
    $rbExportOnly.Add_CheckedChanged($updateEnableState)
    $rbCompare.Add_CheckedChanged($updateEnableState)
    $chkSsh.Add_CheckedChanged($updateEnableState)
    $cmbA5Source.Add_SelectedIndexChanged($refreshA5List)
    $cmbA5Db.Add_DropDown({
        if (-not $cmbA5Db.Tag) { & $refreshA5List }
    }.GetNewClosure())
    $btnRefreshA5.Add_Click($refreshA5List)

    $btnBrowseBefore.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        try {
            $ofd.Title = 'テスト開始前の分析Excelファイル (.xlsx) を選択してください'
            $ofd.Filter = 'Excel ブック (*.xlsx)|*.xlsx|すべてのファイル (*.*)|*.*'
            $ofd.RestoreDirectory = $true
            if (-not [string]::IsNullOrWhiteSpace($txtBeforeFile.Text)) {
                $d = Split-Path -Parent $txtBeforeFile.Text
                if ($d -and (Test-Path -LiteralPath $d)) { $ofd.InitialDirectory = $d }
            }
            if ($ofd.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
                $txtBeforeFile.Text = $ofd.FileName
                if ([string]::IsNullOrWhiteSpace($txtOutFile.Text)) {
                    $stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
                    $txtOutFile.Text = Join-Path (Split-Path -Parent $ofd.FileName) "Aurora_DiffReport_${stamp}.xlsx"
                }
            }
        }
        finally { $ofd.Dispose() }
    }.GetNewClosure())

    $btnBrowseAfter.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        try {
            $ofd.Title = 'テスト開始後の分析Excelファイル (.xlsx) を選択してください'
            $ofd.Filter = 'Excel ブック (*.xlsx)|*.xlsx|すべてのファイル (*.*)|*.*'
            $ofd.RestoreDirectory = $true
            if (-not [string]::IsNullOrWhiteSpace($txtAfterFile.Text)) {
                $d = Split-Path -Parent $txtAfterFile.Text
                if ($d -and (Test-Path -LiteralPath $d)) { $ofd.InitialDirectory = $d }
            }
            elseif (-not [string]::IsNullOrWhiteSpace($txtBeforeFile.Text)) {
                $d = Split-Path -Parent $txtBeforeFile.Text
                if ($d -and (Test-Path -LiteralPath $d)) { $ofd.InitialDirectory = $d }
            }
            if ($ofd.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
                $txtAfterFile.Text = $ofd.FileName
                if ([string]::IsNullOrWhiteSpace($txtOutFile.Text)) {
                    $stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
                    $txtOutFile.Text = Join-Path (Split-Path -Parent $ofd.FileName) "Aurora_DiffReport_${stamp}.xlsx"
                }
            }
        }
        finally { $ofd.Dispose() }
    }.GetNewClosure())

    $btnSwapFiles.Add_Click({
        $tmp = $txtBeforeFile.Text
        $txtBeforeFile.Text = $txtAfterFile.Text
        $txtAfterFile.Text = $tmp
    }.GetNewClosure())

    $btnBrowseA5List.Add_Click({
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        try {
            $dialog.Title = 'A5:SQL Mk-2 からエクスポートしたDBリストを選択'
            $dialog.Filter = 'A5:SQL Mk-2 DBリスト (*.a5dblist;*.dblist)|*.a5dblist;*.dblist'
            $dialog.RestoreDirectory = $true
            if ($dialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
                $idx = $cmbA5Source.Items.IndexOf($dialog.FileName)
                if ($idx -lt 0) { $idx = $cmbA5Source.Items.Add($dialog.FileName) }
                if ($idx -eq $cmbA5Source.SelectedIndex) { & $refreshA5List }
                else { $cmbA5Source.SelectedIndex = $idx }
            }
        }
        finally { $dialog.Dispose() }
    }.GetNewClosure())

    $btnBrowseIn.Add_Click({
        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
        try {
            $fbd.Description = 'CSVファイルが格納された入力フォルダ (またはエクスポート先フォルダ) を選択してください'
            if (-not [string]::IsNullOrWhiteSpace($txtInFolder.Text) -and (Test-Path -LiteralPath $txtInFolder.Text)) {
                $fbd.SelectedPath = $txtInFolder.Text
            }
            if ($fbd.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
                $txtInFolder.Text = $fbd.SelectedPath
                if ([string]::IsNullOrWhiteSpace($txtOutFile.Text)) {
                    $stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
                    $txtOutFile.Text = Join-Path $fbd.SelectedPath "Aurora_PerfAnalysis_${stamp}.xlsx"
                }
                & $inspectResumeFolder
            }
        }
        finally { $fbd.Dispose() }
    }.GetNewClosure())

    $btnBrowseOut.Add_Click({
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        try {
            $sfd.Title = if ($rbCompare.Checked) { '比較結果Excelファイルの保存先を指定してください' } else { '分析結果Excelファイルの保存先を指定してください' }
            $sfd.Filter = 'Excel ブック (*.xlsx)|*.xlsx|すべてのファイル (*.*)|*.*'
            $sfd.DefaultExt = 'xlsx'
            $stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
            if ($rbCompare.Checked) {
                $sfd.FileName = "Aurora_DiffReport_${stamp}.xlsx"
            }
            elseif ($rbDirect.Checked) {
                $rawTag = if (-not [string]::IsNullOrWhiteSpace($cmbA5Db.Text)) { $cmbA5Db.Text.Trim() } elseif (-not [string]::IsNullOrWhiteSpace($txtDbName.Text)) { $txtDbName.Text.Trim() } else { '' }
                $safeTag = ($rawTag -replace '[\\/:*?"<>|\s]+', '_').Trim('_')
                $sfd.FileName = if ($safeTag) { "Aurora_PerfAnalysis_${safeTag}_${stamp}.xlsx" } else { "Aurora_PerfAnalysis_${stamp}.xlsx" }
            }
            else {
                $sfd.FileName = "Aurora_PerfAnalysis_${stamp}.xlsx"
            }
            if (-not [string]::IsNullOrWhiteSpace($txtOutFile.Text)) {
                $dir = Split-Path -Parent $txtOutFile.Text
                if (-not [string]::IsNullOrWhiteSpace($dir) -and (Test-Path -LiteralPath $dir)) { $sfd.InitialDirectory = $dir }
                $sfd.FileName = Split-Path -Leaf $txtOutFile.Text
            }
            elseif ($rbCompare.Checked -and -not [string]::IsNullOrWhiteSpace($txtAfterFile.Text)) {
                $dir = Split-Path -Parent $txtAfterFile.Text
                if (-not [string]::IsNullOrWhiteSpace($dir) -and (Test-Path -LiteralPath $dir)) { $sfd.InitialDirectory = $dir }
            }
            elseif ($rbDirect.Checked -and $cmbA5Source.SelectedIndex -gt 0) {
                $dir = Split-Path -Parent ([string]$cmbA5Source.SelectedItem)
                if (-not [string]::IsNullOrWhiteSpace($dir) -and (Test-Path -LiteralPath $dir)) { $sfd.InitialDirectory = $dir }
            }
            elseif (-not [string]::IsNullOrWhiteSpace($txtInFolder.Text) -and (Test-Path -LiteralPath $txtInFolder.Text)) {
                $sfd.InitialDirectory = $txtInFolder.Text
            }
            if ($sfd.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
                $txtOutFile.Text = $sfd.FileName
            }
        }
        finally { $sfd.Dispose() }
    }.GetNewClosure())

    $btnBrowseKey.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        try {
            $ofd.Title = 'EC2踏み台用 SSH秘密鍵ファイル (.pem / id_rsa) を選択してください'
            $ofd.Filter = 'SSH秘密鍵 (*.pem;*.key;id_*)|*.pem;*.key;id_*|すべてのファイル (*.*)|*.*'
            if ($ofd.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
                $txtSshKey.Text = $ofd.FileName
            }
        }
        finally { $ofd.Dispose() }
    }.GetNewClosure())

    $btnRun.Add_Click({
        if ($guiRunState.IsRunning) { return }
        $runMode = if ($rbCompare.Checked) { 'CompareAnalysis' } elseif ($rbDirect.Checked) { 'DirectAnalyze' } elseif ($rbFull.Checked) { 'ExportAndAnalyze' } elseif ($rbExportOnly.Checked) { 'ExportOnly' } else { 'AnalyzeOnly' }
        $inDir = $txtInFolder.Text.Trim()
        $outXlsx = $txtOutFile.Text.Trim()
        $beforeFile = $txtBeforeFile.Text.Trim()
        $afterFile = $txtAfterFile.Text.Trim()

        if ($runMode -eq 'CompareAnalysis') {
            if ([string]::IsNullOrWhiteSpace($beforeFile) -or [string]::IsNullOrWhiteSpace($afterFile)) {
                & $ShowMessage $form 'テスト開始前およびテスト開始後の分析ファイル (.xlsx) を2ファイル指定してください。' '入力チェック' 'Warning'
                return
            }
            if (-not (Test-Path -LiteralPath $beforeFile -PathType Leaf)) {
                & $ShowMessage $form "テスト開始前の分析ファイルが見つかりません:`n$beforeFile" '入力チェック' 'Warning'
                return
            }
            if (-not (Test-Path -LiteralPath $afterFile -PathType Leaf)) {
                & $ShowMessage $form "テスト開始後の分析ファイルが見つかりません:`n$afterFile" '入力チェック' 'Warning'
                return
            }
            if ([string]::IsNullOrWhiteSpace($outXlsx)) {
                $stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
                $outXlsx = Join-Path (Split-Path -Parent $afterFile) "Aurora_DiffReport_${stamp}.xlsx"
                $txtOutFile.Text = $outXlsx
            }
        }
        elseif ($runMode -eq 'DirectAnalyze') {
            if ([string]::IsNullOrWhiteSpace($outXlsx)) {
                & $ShowMessage $form '出力先Excelファイル (.xlsx) を指定してください。' '入力チェック' 'Warning'
                return
            }
        }
        else {
            if ([string]::IsNullOrWhiteSpace($inDir)) {
                & $ShowMessage $form '入力 / CSV保存フォルダを指定してください。' '入力チェック' 'Warning'
                return
            }
            if ($runMode -eq 'AnalyzeOnly' -and -not (Test-Path -LiteralPath $inDir)) {
                & $ShowMessage $form "指定された入力フォルダが存在しません:`n$inDir" '入力チェック' 'Warning'
                return
            }
            if (($runMode -eq 'AnalyzeOnly' -or $runMode -eq 'ExportAndAnalyze') -and [string]::IsNullOrWhiteSpace($outXlsx)) {
                $stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
                $outXlsx = Join-Path $inDir "Aurora_PerfAnalysis_${stamp}.xlsx"
                $txtOutFile.Text = $outXlsx
            }
        }

        $connCfg = $null
        if ($runMode -eq 'ExportAndAnalyze' -or $runMode -eq 'ExportOnly' -or $runMode -eq 'DirectAnalyze') {
            $dbPortVal = 3306; [void][int]::TryParse($txtDbPort.Text, [ref]$dbPortVal)
            $sshPortVal = 22; [void][int]::TryParse($txtSshPort.Text, [ref]$sshPortVal)
            $connCfg = @{
                ServerName    = $txtDbHost.Text.Trim()
                Port          = $dbPortVal
                Database      = $txtDbName.Text.Trim()
                UserName      = $txtDbUser.Text.Trim()
                Password      = $txtDbPass.Text
                UseSshTunnel  = $chkSsh.Checked
                SshHost       = $txtSshHost.Text.Trim()
                SshPort       = $sshPortVal
                SshUser       = $txtSshUser.Text.Trim()
                SshKeyPath    = $txtSshKey.Text.Trim()
                SshPassphrase = $txtSshPass.Text
                SslMode       = [string]$cmbSsl.SelectedItem
                SslCaCert     = [string]$cmbSsl.Tag
            }
            if ([string]::IsNullOrWhiteSpace($connCfg.ServerName) -or [string]::IsNullOrWhiteSpace($connCfg.Database) -or [string]::IsNullOrWhiteSpace($connCfg.UserName)) {
                & $ShowMessage $form 'Auroraホスト、スキーマ名、DBユーザー名を入力してください。' '接続設定チェック' 'Warning'
                return
            }
            if ($connCfg.UseSshTunnel -and ([string]::IsNullOrWhiteSpace($connCfg.SshHost) -or [string]::IsNullOrWhiteSpace($connCfg.SshUser))) {
                & $ShowMessage $form 'SSHトンネルが有効ですが、EC2踏み台ホストまたはSSHユーザーが未入力です。' 'SSH設定チェック' 'Warning'
                return
            }
        }

        $maxR = 0; [void][int]::TryParse($txtMaxRows.Text, [ref]$maxR)
        $useResume = $chkResume.Checked -and ($runMode -eq 'ExportAndAnalyze' -or $runMode -eq 'ExportOnly')
        $useDetailed = $chkDetailedAnalysis.Checked -and ($runMode -eq 'AnalyzeOnly' -or $runMode -eq 'ExportAndAnalyze' -or $runMode -eq 'DirectAnalyze')

        $cancelState = New-Object AuroraAnalyzer.OperationCancelState
        $guiRunState.IsRunning = $true
        $guiRunState.CancelState = $cancelState
        $guiRunState.CloseAfterCancel = $false

        $gbMode.Enabled = $false
        $gbPath.Enabled = $false
        $gbA5.Enabled = $false
        $gbCompare.Enabled = $false
        $chkDetailedAnalysis.Enabled = $false
        $chkOpenExcel.Enabled = $false
        $chkResume.Enabled = $false
        $btnRun.Enabled = $false
        $btnCancel.Enabled = $true
        $btnCancel.BackColor = $cWarn
        $btnCancel.ForeColor = $cWhite
        $btnClose.Enabled = $true
        $pb.Value = 0
        $lblProgress.Text = '全体 0%'
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

        try {
            $res = & $runPipeline `
                -RunMode $runMode `
                -ConnConfig $connCfg `
                -DbLabel $cmbA5Db.Text.Trim() `
                -CsvFolder $inDir `
                -ExcelOutputPath $outXlsx `
                -MaxRows $maxR `
                -Engine $InitialExcelEngine `
                -LogCallback $appendLog `
                -BeforeAnalysisPath $beforeFile `
                -AfterAnalysisPath $afterFile `
                -Resume:$useResume `
                -DetailedAnalysis:$useDetailed `
                -CancelState $cancelState `
                -PumpCallback $pumpUi

            $sheetModeLabel = if ($useDetailed) { '詳細分析 (8シート)' } else { '標準出力 (3シート)' }
            $msgDone = if ($runMode -eq 'ExportOnly') {
                $skipCnt = @($res.ExportedTables | Where-Object { $_.Skipped }).Count
                $newCnt = $res.ExportedTables.Count - $skipCnt
                if ($useResume -and $skipCnt -gt 0) {
                    "CSV一括エクスポート（続行）が完了しました。`n保存先: $inDir`nテーブル数: $($res.ExportedTables.Count) (新規出力: $newCnt / 既存スキップ: $skipCnt)"
                }
                else {
                    "CSV一括エクスポートが完了しました。`n保存先: $inDir`nテーブル数: $($res.ExportedTables.Count)"
                }
            }
            elseif ($runMode -eq 'CompareAnalysis') {
                $cmp = $res.Comparison
                "テスト前後の分析ファイル比較レポートの出力が完了しました。`n`n出力ファイル: $($res.OutputPath)`n照合方式: $($cmp.PrecisionLabel)`n比較対象テーブル: $($cmp.TotalUniqueTableCount) 件 (変更 $($cmp.ChangedTableCount) [更新 $($cmp.UpdatedSheetCount) / 追加 $($cmp.AddedSheetCount) / 削除 $($cmp.DeletedTableCount)] / 変更なし $($cmp.UnchangedTableCount))`n比較対象カラム: $($cmp.TotalUniqueColumnCount) 列 (変更あり $($cmp.ChangedColumnTotal) 列 / 変更なし $($cmp.UnchangedColumnTotal) 列)`n総レコード増減: $(if ($cmp.TotalRowDelta -gt 0) { '+' } else { '' })$($cmp.TotalRowDelta.ToString('N0')) 件"
            }
            elseif ($runMode -eq 'DirectAnalyze') {
                "DB直接接続による性能試験用データ分析レポートの出力が完了しました（CSV未出力）。`n`n出力ファイル: $($res.OutputPath)`n出力シート構成: $sheetModeLabel`n対象テーブル数: $($res.Analysis.TableCount) / 総レコード数: $($res.Analysis.TotalRows.ToString('N0')) 件`n検出警告: 高 $($res.Analysis.HighWarningCount) 件 / 合計 $($res.Analysis.TotalWarningCount) 件"
            }
            else {
                "性能試験用データ分析レポートの出力が完了しました。`n`n出力ファイル: $($res.OutputPath)`n出力シート構成: $sheetModeLabel`n対象テーブル数: $($res.Analysis.TableCount) / 総レコード数: $($res.Analysis.TotalRows.ToString('N0')) 件`n検出警告: 高 $($res.Analysis.HighWarningCount) 件 / 合計 $($res.Analysis.TotalWarningCount) 件"
            }
            & $ShowMessage $form $msgDone '完了' 'Information'

            if ($chkOpenExcel.Checked -and $runMode -ne 'ExportOnly' -and (Test-Path -LiteralPath $res.OutputPath)) {
                Start-Process -FilePath $res.OutputPath
            }
        }
        catch {
            $isCancelled = ($cancelState.IsCancelled) -or ($_.Exception -is [System.OperationCanceledException]) -or ($null -ne $_.Exception.InnerException -and $_.Exception.InnerException -is [System.OperationCanceledException])
            if ($isCancelled) {
                & $appendLog '[中断] 処理を安全に中断しました。' -1
                $csvRun = $runMode -in 'ExportAndAnalyze','ExportOnly'
                if ($csvRun) {
                    $summary = & $getResumeSummary $inDir
                    $canResume = $null -ne $summary -and $summary.ProgressStatus -in 'Interrupted','InProgress','CsvCompleted'
                    $chkResume.Checked = $useResume -or $canResume
                }
                if (-not $guiRunState.CloseAfterCancel) {
                    $cancelMsg = if ($csvRun -and $canResume) {
                        "処理を中断しました。完了済みCSVは $($summary.RecordedDone)/$($summary.TotalTables) テーブルです。`n`n同じフォルダで「出力済みCSVから続行」を選び、実行開始を押してください。中断したテーブルは先頭から取得します。"
                    }
                    elseif ($csvRun) {
                        "今回のCSV出力を開始する前に中断しました。設定を確認して再実行してください。"
                    }
                    else {
                        "処理を中断しました。`n接続およびバックグラウンド処理は安全に終了しました。"
                    }
                    & $ShowMessage $form $cancelMsg '中断' 'Information'
                }
            }
            else {
                $errMsg = $_.Exception.Message
                & $appendLog ("[エラー] $errMsg") 0
                & $ShowMessage $form "処理中にエラーが発生しました:`n`n$errMsg" 'エラー' 'Error'
            }
        }
        finally {
            $shouldClose = [bool]$guiRunState.CloseAfterCancel
            $guiRunState.IsRunning = $false
            $guiRunState.CancelState = $null
            $guiRunState.CloseAfterCancel = $false
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $gbMode.Enabled = $true
            $gbPath.Enabled = $true
            $btnRun.Enabled = $true
            $btnCancel.Enabled = $false
            $btnCancel.BackColor = $cLight
            $btnCancel.ForeColor = $cMuted
            $btnClose.Enabled = $true
            & $updateEnableState
            if ($shouldClose -and -not $form.IsDisposed) {
                $form.Close()
            }
        }
    }.GetNewClosure())

    # With no explicit source or DB, defer discovery until the user opens the DB list.
    if ($InitialA5DbListPath -or $InitialA5DbName) { & $refreshA5List }
    if ($InitialMode -eq 'DirectAnalyze') {
        $rbDirect.Checked = $true
    }
    elseif ($InitialMode -eq 'ExportAndAnalyze') {
        $rbFull.Checked = $true
    }
    elseif ($InitialMode -eq 'ExportOnly') {
        $rbExportOnly.Checked = $true
    }
    elseif ($InitialMode -eq 'AnalyzeOnly') {
        $rbAnalyze.Checked = $true
    }
    elseif ($InitialMode -eq 'CompareAnalysis' -or -not [string]::IsNullOrWhiteSpace($InitialBeforeAnalysisPath) -or -not [string]::IsNullOrWhiteSpace($InitialAfterAnalysisPath)) {
        $rbCompare.Checked = $true
    }
    elseif (-not [string]::IsNullOrWhiteSpace($InitialA5DbName)) {
        $rbFull.Checked = $true
    }
    & $updateEnableState
    if (-not [string]::IsNullOrWhiteSpace($InitialInputFolder)) { & $inspectResumeFolder }
    $form.Add_Disposed({
        foreach ($f in @($fontMain, $fontBold, $fontTitle, $fontSmall)) {
            if ($null -ne $f) { try { $f.Dispose() } catch { } }
        }
    }.GetNewClosure())
    return $form
}

# ==============================================================================
#  Main Entry Point
# ==============================================================================
if ($MyInvocation.InvocationName -ne '.') {
    if ($Mode -eq 'Gui' -and -not $NoGui) {
        $frm = New-AnalyzerGuiForm `
            -InitialRegistryRoot $A5RegistryRoot `
            -InitialA5DbListPath $A5DbListPath `
            -InitialInputFolder $InputFolder `
            -InitialOutputPath $OutputPath `
            -InitialA5DbName $A5DbName `
            -InitialBeforeAnalysisPath $BeforeAnalysisPath `
            -InitialAfterAnalysisPath $AfterAnalysisPath `
            -InitialResume:$Resume -InitialDetailedAnalysis:$DetailedAnalysis -InitialExcelEngine $ExcelEngine
        [void]$frm.ShowDialog()
        $frm.Dispose()
        exit 0
    }

    # Headless / CLI Mode
    $resolvedMode = if ($Mode -eq 'Gui') {
        if (-not [string]::IsNullOrWhiteSpace($BeforeAnalysisPath) -or -not [string]::IsNullOrWhiteSpace($AfterAnalysisPath)) {
            'CompareAnalysis'
        }
        elseif ([string]::IsNullOrWhiteSpace($InputFolder) -and (
                -not [string]::IsNullOrWhiteSpace($A5DbName) -or
                -not [string]::IsNullOrWhiteSpace($A5DbListPath) -or
                -not [string]::IsNullOrWhiteSpace($ConnectionStringOverride) -or
                -not [string]::IsNullOrWhiteSpace($DbHost))) {
            'DirectAnalyze'
        }
        else {
            'AnalyzeOnly'
        }
    }
    else {
        $Mode
    }
    $connCfg = $null
    $dbLabel = $A5DbName

    $cliCancelState = New-Object AuroraAnalyzer.OperationCancelState
    $consoleCancelReg = $cliCancelState.AttachConsoleCancel()
    $consolePump = $null
    $previousInputMode = $null
    try {
        # Windows PowerShell can stop its pipeline before a Ctrl+C signal reaches script cleanup.
        # Handle keyboard input cooperatively and keep CancelKeyPress for console control events.
        if (-not [Console]::IsInputRedirected) {
            try {
                $previousInputMode = [Console]::TreatControlCAsInput
                [Console]::TreatControlCAsInput = $true
                $consolePump = {
                    while ([Console]::KeyAvailable) {
                        $key = [Console]::ReadKey($true)
                        if ($key.KeyChar -eq [char]3 -or ($key.Key -eq 'C' -and ($key.Modifiers -band [ConsoleModifiers]::Control))) {
                            $cliCancelState.Cancel()
                            break
                        }
                    }
                }
            }
            catch [IO.IOException] { $previousInputMode = $null }
        }
        if ($resolvedMode -eq 'CompareAnalysis') {
            if ([string]::IsNullOrWhiteSpace($BeforeAnalysisPath) -or [string]::IsNullOrWhiteSpace($AfterAnalysisPath)) {
                throw "-BeforeAnalysisPath および -AfterAnalysisPath に比較対象の分析ファイル (.xlsx) を2ファイル指定してください。"
            }
            $res = Invoke-AnalyzerPipeline `
                -RunMode 'CompareAnalysis' `
                -ExcelOutputPath $OutputPath `
                -Engine $ExcelEngine `
                -BeforeAnalysisPath $BeforeAnalysisPath `
                -AfterAnalysisPath $AfterAnalysisPath `
                -CancelState $cliCancelState -PumpCallback $consolePump

            if ($OpenAfterCreate -and (Test-Path -LiteralPath $res.OutputPath)) {
                Start-Process -FilePath $res.OutputPath
            }
            exit 0
        }

        if ($resolvedMode -eq 'ExportAndAnalyze' -or $resolvedMode -eq 'ExportOnly' -or $resolvedMode -eq 'DirectAnalyze') {
            $a5Entry = $null
            if (-not [string]::IsNullOrWhiteSpace($ConnectionStringOverride)) {
                $a5Entry = Parse-A5ConnectionString -DbName $(if ($A5DbName) { $A5DbName } else { 'Override' }) -RawConnectionString $ConnectionStringOverride
            }
            elseif ($A5DbListPath -or -not [string]::IsNullOrWhiteSpace($A5DbName)) {
                if ([string]::IsNullOrWhiteSpace($A5DbName)) { throw 'DBリスト内の登録DB名を -A5DbName で指定してください。' }
                $allA5 = @(Get-A5RegisteredDatabases -RegistryRoot $A5RegistryRoot -DbListPath $A5DbListPath)
                $a5Entry = $allA5 | Where-Object { $_.RegisteredName -eq $A5DbName } | Select-Object -First 1
                if ($null -eq $a5Entry) {
                    $avail = ($allA5 | ForEach-Object { $_.RegisteredName }) -join ', '
                    throw "A5:SQL Mk-2 登録データベース名 '$A5DbName' が見つかりません。検出済み: [$avail]"
                }
            }

            $effHost    = if (-not [string]::IsNullOrWhiteSpace($DbHost)) { $DbHost } elseif ($a5Entry) { $a5Entry.ServerName } else { '' }
            $effPort    = if ($DbPort -gt 0) { $DbPort } elseif ($a5Entry) { $a5Entry.Port } else { 3306 }
            $effDb      = if (-not [string]::IsNullOrWhiteSpace($DbName)) { $DbName } elseif ($a5Entry) { $a5Entry.Database } else { '' }
            $effUser    = if (-not [string]::IsNullOrWhiteSpace($DbUser)) { $DbUser } elseif ($a5Entry) { $a5Entry.UserName } else { '' }
            $effPass    = if (-not [string]::IsNullOrEmpty($DbPassword)) { $DbPassword } elseif ($a5Entry) { $a5Entry.PlainPassword } else { '' }
            $effUseSsh  = if ($DisableSsh) { $false } elseif (-not [string]::IsNullOrWhiteSpace($SshHost)) { $true } elseif ($a5Entry) { [bool]$a5Entry.UseSshTunnel } else { $false }
            $effSshHost = if (-not [string]::IsNullOrWhiteSpace($SshHost)) { $SshHost } elseif ($a5Entry) { $a5Entry.SshHost } else { '' }
            $effSshPort = if ($SshPort -gt 0) { $SshPort } elseif ($a5Entry) { $a5Entry.SshPort } else { 22 }
            $effSshUser = if (-not [string]::IsNullOrWhiteSpace($SshUser)) { $SshUser } elseif ($a5Entry) { $a5Entry.SshUser } else { '' }
            $effSshKey  = if (-not [string]::IsNullOrWhiteSpace($SshKeyPath)) { $SshKeyPath } elseif ($a5Entry) { $a5Entry.SshKeyFile } else { '' }
            $effSsl     = if ($SslMode -ne 'Auto') { $SslMode } elseif ($a5Entry) { $a5Entry.SslMode } else { 'Preferred' }
            $effSslCa   = if (-not [string]::IsNullOrWhiteSpace($SslCaCertPath)) { $SslCaCertPath } elseif ($a5Entry) { $a5Entry.SslCaCert } else { '' }

            $connCfg = @{
                ServerName    = $effHost
                Port          = $effPort
                Database      = $effDb
                UserName      = $effUser
                Password      = $effPass
                UseSshTunnel  = $effUseSsh
                SshHost       = $effSshHost
                SshPort       = $effSshPort
                SshUser       = $effSshUser
                SshKeyPath    = $effSshKey
                SshPassphrase = $SshPassphrase
                SslMode       = $effSsl
                SslCaCert     = $effSslCa
            }
        }

        if ($resolvedMode -ne 'DirectAnalyze' -and [string]::IsNullOrWhiteSpace($InputFolder)) {
            throw "-InputFolder を指定してください。"
        }

        $res = Invoke-AnalyzerPipeline `
            -RunMode $resolvedMode `
            -ConnConfig $connCfg `
            -DbLabel $dbLabel `
            -CsvFolder $InputFolder `
            -ExcelOutputPath $OutputPath `
            -MaxRows $MaxRowsPerTable `
            -IncTables $IncludeTables `
            -ExcTables $ExcludeTables `
            -Engine $ExcelEngine `
            -Resume:$Resume `
            -DetailedAnalysis:$DetailedAnalysis `
            -CancelState $cliCancelState -PumpCallback $consolePump

        if ($OpenAfterCreate -and $resolvedMode -ne 'ExportOnly' -and (Test-Path -LiteralPath $res.OutputPath)) {
            Start-Process -FilePath $res.OutputPath
        }
    }
    catch {
        $isCancelled = ($cliCancelState.IsCancelled) -or ($_.Exception -is [System.OperationCanceledException]) -or ($null -ne $_.Exception.InnerException -and $_.Exception.InnerException -is [System.OperationCanceledException])
        if ($isCancelled) {
            Write-Warning "処理を中断して安全に終了しました。CSV出力モードの場合は同じ -InputFolder を指定し -Resume を付けて再実行すると続行できます。"
            exit 130
        }
        throw
    }
    finally {
        if ($null -ne $previousInputMode) { [Console]::TreatControlCAsInput = $previousInputMode }
        if ($null -ne $consoleCancelReg) { $consoleCancelReg.Dispose() }
    }
}
