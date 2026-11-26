#requires -Version 5.1
<#
.SYNOPSIS
    End-to-End Verification Suite for AuroraCsvAnalyzer.ps1 and Run_AuroraCsvAnalyzer.bat
#>
param([switch]$PortableOnly)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$rootDir  = Split-Path -Parent $PSScriptRoot
$batPath  = Join-Path $rootDir 'Run_AuroraCsvAnalyzer.bat'
$ps1Path  = Join-Path $rootDir 'AuroraCsvAnalyzer.ps1'
$workDir  = Join-Path ([System.IO.Path]::GetTempPath()) ("AuroraAnalyzerVerify_" + [Guid]::NewGuid().ToString('N').Substring(0, 8))
$testRegRoot = 'HKCU:\Software\mmatsubara_verify_tmp_' + [Guid]::NewGuid().ToString('N')

$passCount = 0
$failCount = 0

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if ($Condition) {
        $script:passCount++
        Write-Host "  [PASS] $Message" -ForegroundColor Green
    }
    else {
        $script:failCount++
        Write-Host "  [FAIL] $Message" -ForegroundColor Red
    }
}

try {
    New-Item -ItemType Directory -Path $workDir -Force | Out-Null
    Write-Host "=== 1. Encoding & Syntax Checks ===" -ForegroundColor Cyan

    $ps1Bytes = [System.IO.File]::ReadAllBytes($ps1Path)
    Assert-True ($ps1Bytes.Length -ge 3 -and $ps1Bytes[0] -eq 0xEF -and $ps1Bytes[1] -eq 0xBB -and $ps1Bytes[2] -eq 0xBF) "AuroraCsvAnalyzer.ps1 has UTF-8 BOM (required for PowerShell 5.1 Japanese strings)"

    $batBytes = [System.IO.File]::ReadAllBytes($batPath)
    Assert-True (-not ($batBytes.Length -ge 3 -and $batBytes[0] -eq 0xEF -and $batBytes[1] -eq 0xBB -and $batBytes[2] -eq 0xBF)) "Run_AuroraCsvAnalyzer.bat has NO UTF-8 BOM (cmd.exe safe)"

    $tokens = $null; $parseErrors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($ps1Path, [ref]$tokens, [ref]$parseErrors)
    Assert-True ($parseErrors.Count -eq 0) "AuroraCsvAnalyzer.ps1 parses with 0 PowerShell 5.1 syntax errors (found: $($parseErrors.Count))"
    if ($parseErrors.Count -gt 0) {
        foreach ($pe in $parseErrors) { Write-Host "    ParseError: $($pe.Message) at Line $($pe.Extent.StartLineNumber)" -ForegroundColor Red }
    }

    Write-Host "`n=== 2. Dot-Source & Embedded C# 5.0 Compilation ===" -ForegroundColor Cyan
    . $ps1Path
    Assert-True ($null -ne ([System.Management.Automation.PSTypeName]'AuroraAnalyzer.MySql84Client').Type) "AuroraAnalyzer.MySql84Client compiled cleanly"
    Assert-True ($null -ne ([System.Management.Automation.PSTypeName]'AuroraAnalyzer.OperationCancelState').Type) "AuroraAnalyzer.OperationCancelState compiled cleanly"
    Assert-True ($null -ne ([System.Management.Automation.PSTypeName]'AuroraAnalyzer.BackgroundTaskLauncher').Type) "AuroraAnalyzer.BackgroundTaskLauncher compiled cleanly"
    Assert-True ($null -ne ([System.Management.Automation.PSTypeName]'AuroraAnalyzer.CsvAnalysisEngine').Type) "AuroraAnalyzer.CsvAnalysisEngine compiled cleanly"
    Assert-True ($null -ne ([System.Management.Automation.PSTypeName]'AuroraAnalyzer.DirectAnalysisSession').Type) "AuroraAnalyzer.DirectAnalysisSession compiled cleanly"
    Assert-True ($null -ne ([System.Management.Automation.PSTypeName]'AuroraAnalyzer.AnalysisComparisonEngine').Type) "AuroraAnalyzer.AnalysisComparisonEngine compiled cleanly"

    # Verify OLE color conversion (BGR!)
    $oleNavy = ConvertTo-ExcelOleColor '#1A2B4C'
    $oleBurg = ConvertTo-ExcelOleColor '#800020'
    $oleGray = ConvertTo-ExcelOleColor '#333333'
    Assert-True ($oleNavy -eq (0x1A + 0x2B * 256 + 0x4C * 65536)) "ConvertTo-ExcelOleColor('#1A2B4C') produces exact BGR ($oleNavy)"
    Assert-True ($oleBurg -eq (0x80 + 0x00 * 256 + 0x20 * 65536)) "ConvertTo-ExcelOleColor('#800020') produces exact BGR ($oleBurg)"
    Assert-True ($oleGray -eq (0x33 + 0x33 * 256 + 0x33 * 65536)) "ConvertTo-ExcelOleColor('#333333') produces exact BGR ($oleGray)"

    Write-Host "`n=== 3. A5:SQL Mk-2 portable DB list discovery ===" -ForegroundColor Cyan
    $listPath = Join-Path $workDir '接続先 [検証].a5dblist'
    $portableLines = @(
        '# A5:SQL Mk-2 Database list'
        '[DatabaseList]'
        '検証/東京=__ConnectionType=Internal;ProviderName=MySQL;ServerName=primary.example;Port=3307;Database="shop;east";UserName=user=qa;Password=********;__SSH=True;__SSHHostName=bastion.example;__SSHPort=2222;__SSHUserName=tester;__SSHKeyFile=keys\検証.pem;__SSHPassword=********;SSLMode=Verify_CA;SSLCACert=certs\ca.pem'
        '予備=ProviderName=MySQL;ServerName=second.example;Database=other;UserName=tester;__SSH=False;__SSHHostName=unused.example;UseSSL=False'
        '[Other]'
        'ignored=ServerName=other.example'
    )
    foreach ($enc in [Text.Encoding]::UTF8, [Text.UTF8Encoding]::new($false), [Text.Encoding]::Unicode, [Text.Encoding]::GetEncoding(932)) {
        [IO.File]::WriteAllLines($listPath, $portableLines, $enc)
        $dbs = @(Get-A5RegisteredDatabases -RegistryRoot '' -DbListPath $listPath)
        Assert-True ($dbs.Count -eq 2 -and $dbs[0].RegisteredName -ceq '検証/東京' -and $dbs[1].RegisteredName -ceq '予備') "Read Japanese aliases in file order ($($enc.WebName), BOM bytes=$($enc.GetPreamble().Length))"
    }
    $entry = $dbs[0]
    Assert-True ($entry.ServerName -eq 'primary.example' -and $entry.Port -eq 3307 -and $entry.Database -eq 'shop;east' -and $entry.UserName -ceq 'user=qa') "Portable DB endpoint and quoted or embedded delimiters are preserved"
    Assert-True ($entry.UseSshTunnel -and $entry.SshHost -eq 'bastion.example' -and $entry.SshPort -eq 2222 -and $entry.SshUser -eq 'tester' -and $entry.SshKeyFile -eq (Join-Path $workDir 'keys\検証.pem')) "Portable SSH settings and key path resolve against the list folder"
    Assert-True ($entry.SslMode -eq 'VerifyCA' -and $entry.SslCaCert -eq (Join-Path $workDir 'certs\ca.pem') -and -not $dbs[1].UseSshTunnel) "Portable certificate path and disabled SSH setting are respected"
    Assert-True ($entry.PlainPassword -eq '' -and $entry.SourceFile -eq $listPath -and $entry.SourceRegistry -eq '') "Exported password masks are never used for login"

    $otherList = Join-Path $workDir 'other.dblist'
    [IO.File]::WriteAllText($otherList, '[DatabaseList]')
    Assert-True (@(Get-A5RegisteredDatabases -RegistryRoot '' -DbListPath $otherList).Count -eq 0) "Empty exported DB list is valid and does not fall back to the registry"
    foreach ($invalid in @(
        @{ Text = "[Other]`nentry=ServerName=x"; Error = '\[DatabaseList\]' }
        @{ Text = "[DatabaseList]`ninvalid"; Error = '2 行目' }
        @{ Text = "[DatabaseList]`nDemo=Server=x;malformed-value"; Error = '2 行目' }
        @{ Text = "[DatabaseList]`nDemo=Server=x`ndemo=Server=y"; Error = '重複' }
    )) {
        [IO.File]::WriteAllText($otherList, $invalid.Text)
        $message = ''
        try { $null = Get-A5RegisteredDatabases -DbListPath $otherList }
        catch { $message = $_.Exception.Message }
        Assert-True ($message -match $invalid.Error -and $message.Contains($otherList) -and $message -notmatch 'malformed-value') "Invalid DB list reports its location without echoing connection data ($($invalid.Error))"
    }

    $portableForm = New-AnalyzerGuiForm -InitialA5DbListPath $listPath -InitialA5DbName '予備' -InitialRegistryRoot (Join-Path $workDir 'no-registry')
    try {
        $source = $portableForm.Controls.Find('A5Source', $true)[0]
        $names = $portableForm.Controls.Find('A5DbName', $true)[0]
        $hostField = $names.Parent.Controls | Where-Object Text -eq 'second.example'
        Assert-True ($source.Text -eq $listPath -and $names.Text -ceq '予備' -and $null -ne $hostField) "GUI loads the requested portable profile and host"
        $names.SelectedIndex = 0
        Assert-True ($hostField.Text -eq 'primary.example') "Changing the selected DB updates its connection"
        $source.SelectedIndex = 0
        Assert-True ($names.Items.Count -eq 0 -and $hostField.Text -eq '') "Switching to an empty source clears the previous connection"
        $source.SelectedIndex = 1
        Assert-True ($names.Items.Count -eq 2 -and $hostField.Text -eq 'second.example') "Switching back reloads the selected file"
        $passwordFields = @($names.Parent.Controls | Where-Object { $_ -is [Windows.Forms.TextBox] -and $_.UseSystemPasswordChar })
        foreach ($field in $passwordFields) { $field.Text = [guid]::NewGuid().ToString() }
        $source.SelectedIndex = $source.Items.Add($otherList)
        Assert-True ($names.Items.Count -eq 0 -and $hostField.Text -eq '' -and @($passwordFields | Where-Object Text -ne '').Count -eq 0) "Invalid selections clear the previous endpoint and credentials"
    }
    finally { $portableForm.Dispose() }

    if (-not $PortableOnly) {
        Write-Host "`n=== A5:SQL Mk-2 Registry Discovery & SSH/SSL Connection String Parsing ===" -ForegroundColor Cyan
        $dsKey = Join-Path $testRegRoot 'A5M2(x64)\DataSources'
        New-Item -Path $dsKey -Force | Out-Null
        $sampleConn = '__ConnectionType=Internal;ProviderName=MySQL;ServerName=aurora-perf.cluster-xyz.ap-northeast-1.rds.amazonaws.com;Port=3306;Database=perf_shop;UserName=readonly_user;UseSSL=True;__UseSSH=True;__SSHHost=ec2-bastion.example.com;__SSHPort=2222;__SSHUser=ec2-user;__SSHKeyFile=C:\keys\bastion.pem'
        New-ItemProperty -Path $dsKey -Name 'Aurora_Perf_Prod' -Value $sampleConn -PropertyType String -Force | Out-Null
        $paramsSub = Join-Path $dsKey 'Aurora_Perf_Prod.params'
        New-Item -Path $paramsSub -Force | Out-Null
        New-ItemProperty -Path $paramsSub -Name 'Password' -Value ('A1B2C3D4' * 20) -PropertyType String -Force | Out-Null

        $foundDbs = @(Get-A5RegisteredDatabases -RegistryRoot $testRegRoot)
        Assert-True ($foundDbs.Count -eq 1) "Discovered test A5M2 DB entry from registry (count=$($foundDbs.Count))"
        $entry = $foundDbs[0]
        Assert-True ($entry.RegisteredName -eq 'Aurora_Perf_Prod') "RegisteredName matches"
        Assert-True ($entry.ServerName -eq 'aurora-perf.cluster-xyz.ap-northeast-1.rds.amazonaws.com') "Aurora ServerName parsed"
        Assert-True ($entry.Port -eq 3306 -and $entry.Database -eq 'perf_shop' -and $entry.UserName -eq 'readonly_user') "Port/Database/UserName parsed"
        Assert-True ($entry.UseSshTunnel -eq $true -and $entry.SshHost -eq 'ec2-bastion.example.com' -and $entry.SshPort -eq 2222 -and $entry.SshUser -eq 'ec2-user' -and $entry.SshKeyFile -eq 'C:\keys\bastion.pem') "SSH2 Tunnel EC2 bastion config parsed accurately"
        Assert-True ($entry.SslMode -eq 'Required') "MySQL SSLMode=Required parsed from UseSSL=True"
        Assert-True ($entry.PassIsEncrypted -eq $true -and $entry.PlainPassword -eq '') "Encrypted A5M2 password safely flagged and not treated as plaintext"

        # Also read real host A5M2 registry (read-only smoke check)
        $realDbs = @(Get-A5RegisteredDatabases)
        Assert-True ($realDbs.Count -ge 0) "Read-only discovery against real HKCU:\Software\mmatsubara succeeded ($($realDbs.Count) entries found)"
    }

    Write-Host "`n=== 4. Live MySQL 8.4.7 Wire Protocol Mock Server End-to-End Export Test ===" -ForegroundColor Cyan
    $mockServerCs = @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace AuroraTestMock
{
    public class MockMySql84Server : IDisposable
    {
        private TcpListener _listener;
        private Thread _thread;
        private volatile bool _stop;
        private string _authMode; // "fast", "rsa", "native"
        private string _expectedPassword;
        private RSACryptoServiceProvider _rsa;
        private string _rsaPem;
        public int Port { get; private set; }
        public bool AuthSucceeded { get; private set; }
        public int QueriesServed { get; private set; }
        public int RowDelayMs { get; set; }
        public int DataRepeats { get; set; }
        public string RowEstimate { get; set; }
        public bool BinaryMemo { get; set; }
        public string LastError { get; private set; }
        private readonly object _sqlSync = new object();
        private readonly List<string> _executedSqls = new List<string>();

        public string[] GetExecutedSqls()
        {
            lock (_sqlSync) { return _executedSqls.ToArray(); }
        }

        public void ClearExecutedSqls()
        {
            lock (_sqlSync) { _executedSqls.Clear(); }
        }

        public int CountTableSelects(string tableName)
        {
            int count = 0;
            lock (_sqlSync)
            {
                for (int i = 0; i < _executedSqls.Count; i++)
                {
                    if (_executedSqls[i].IndexOf("SELECT * FROM", StringComparison.OrdinalIgnoreCase) >= 0 &&
                        _executedSqls[i].IndexOf(tableName, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        count++;
                    }
                }
            }
            return count;
        }

        public MockMySql84Server(string authMode, string expectedPassword)
        {
            _authMode = authMode;
            _expectedPassword = expectedPassword;
            DataRepeats = 1;
            RowEstimate = "3";
            _rsa = new RSACryptoServiceProvider(2048);
            _rsaPem = ExportPublicKeyPem(_rsa.ExportParameters(false));
            _listener = new TcpListener(IPAddress.Loopback, 0);
            _listener.Start();
            Port = ((IPEndPoint)_listener.LocalEndpoint).Port;
            _thread = new Thread(RunLoop);
            _thread.IsBackground = true;
            _thread.Start();
        }

        private static string ExportPublicKeyPem(RSAParameters p)
        {
            byte[] modInt = DerInt(p.Modulus);
            byte[] expInt = DerInt(p.Exponent);
            byte[] rsaSeq = DerSeq(Concat(modInt, expInt));
            byte[] bitStr = DerBitString(rsaSeq);
            // rsaEncryption OID 1.2.840.113549.1.1.1 + NULL
            byte[] algId = new byte[] { 0x30, 0x0D, 0x06, 0x09, 0x2A, 0x86, 0x48, 0x86, 0xF7, 0x0D, 0x01, 0x01, 0x01, 0x05, 0x00 };
            byte[] spki = DerSeq(Concat(algId, bitStr));
            return "-----BEGIN PUBLIC KEY-----\n" + Convert.ToBase64String(spki) + "\n-----END PUBLIC KEY-----\n";
        }

        private static byte[] Concat(byte[] a, byte[] b)
        {
            byte[] r = new byte[a.Length + b.Length];
            Array.Copy(a, 0, r, 0, a.Length);
            Array.Copy(b, 0, r, a.Length, b.Length);
            return r;
        }
        private static byte[] DerLen(int len)
        {
            if (len < 128) return new byte[] { (byte)len };
            if (len < 256) return new byte[] { 0x81, (byte)len };
            return new byte[] { 0x82, (byte)(len >> 8), (byte)(len & 0xFF) };
        }
        private static byte[] DerWrap(byte tag, byte[] content)
        {
            byte[] l = DerLen(content.Length);
            byte[] r = new byte[1 + l.Length + content.Length];
            r[0] = tag;
            Array.Copy(l, 0, r, 1, l.Length);
            Array.Copy(content, 0, r, 1 + l.Length, content.Length);
            return r;
        }
        private static byte[] DerSeq(byte[] content) { return DerWrap(0x30, content); }
        private static byte[] DerBitString(byte[] content)
        {
            byte[] b = new byte[content.Length + 1];
            b[0] = 0x00;
            Array.Copy(content, 0, b, 1, content.Length);
            return DerWrap(0x03, b);
        }
        private static byte[] DerInt(byte[] val)
        {
            if (val[0] >= 0x80)
            {
                byte[] padded = new byte[val.Length + 1];
                padded[0] = 0x00;
                Array.Copy(val, 0, padded, 1, val.Length);
                return DerWrap(0x02, padded);
            }
            return DerWrap(0x02, val);
        }

        private void RunLoop()
        {
            try
            {
                while (!_stop)
                {
                    if (!_listener.Pending()) { Thread.Sleep(20); continue; }
                    using (TcpClient client = _listener.AcceptTcpClient())
                    {
                        try { HandleClient(client); }
                        catch (IOException) { }
                        catch (SocketException) { }
                    }
                }
            }
            catch (Exception ex) { LastError = ex.ToString(); }
        }

        private void HandleClient(TcpClient client)
        {
            NetworkStream ns = client.GetStream();
            byte seq = 0;
            byte[] nonce = Encoding.ASCII.GetBytes("12345678901234567890");

            // Send Initial Handshake v10 (Aurora MySQL 8.4.7)
            using (MemoryStream ms = new MemoryStream())
            {
                ms.WriteByte(0x0A);
                byte[] ver = Encoding.ASCII.GetBytes("8.4.7-aurora-mock\0");
                ms.Write(ver, 0, ver.Length);
                ms.Write(new byte[] { 1, 0, 0, 0 }, 0, 4); // conn id
                ms.Write(nonce, 0, 8);
                ms.WriteByte(0x00);
                // lower caps: LONG_PASSWORD | LONG_FLAG | CONNECT_WITH_DB | PROTOCOL_41 | TRANSACTIONS | SECURE_CONNECTION
                ushort lowCaps = 0xA20D;
                ms.WriteByte((byte)(lowCaps & 0xFF)); ms.WriteByte((byte)(lowCaps >> 8));
                ms.WriteByte(255); // charset
                ms.Write(new byte[] { 2, 0 }, 0, 2); // status
                // upper caps: PLUGIN_AUTH (0x0008) | MULTI_RESULTS (0x0002)
                ushort hiCaps = 0x000A;
                ms.WriteByte((byte)(hiCaps & 0xFF)); ms.WriteByte((byte)(hiCaps >> 8));
                ms.WriteByte(21); // auth plugin data len
                ms.Write(new byte[10], 0, 10);
                ms.Write(nonce, 8, 12);
                ms.WriteByte(0x00);
                byte[] plug = Encoding.ASCII.GetBytes("caching_sha2_password\0");
                ms.Write(plug, 0, plug.Length);
                WritePkt(ns, ms.ToArray(), ref seq);
            }

            // Read HandshakeResponse41
            byte[] resp = ReadPkt(ns, ref seq);
            if (_authMode == "fast")
            {
                WritePkt(ns, new byte[] { 0x01, 0x03 }, ref seq); // fast_auth_success
                WritePkt(ns, new byte[] { 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00 }, ref seq); // OK
                AuthSucceeded = true;
            }
            else if (_authMode == "rsa")
            {
                WritePkt(ns, new byte[] { 0x01, 0x04 }, ref seq); // perform_full_authentication
                byte[] reqKey = ReadPkt(ns, ref seq);
                if (reqKey.Length != 1 || reqKey[0] != 0x02) throw new Exception("Expected 0x02 public key request");
                byte[] pemBytes = Encoding.ASCII.GetBytes(_rsaPem);
                byte[] keyPkt = new byte[1 + pemBytes.Length];
                keyPkt[0] = 0x01;
                Array.Copy(pemBytes, 0, keyPkt, 1, pemBytes.Length);
                WritePkt(ns, keyPkt, ref seq);

                byte[] encPw = ReadPkt(ns, ref seq);
                byte[] dec = _rsa.Decrypt(encPw, true);
                for (int i = 0; i < dec.Length; i++) dec[i] = (byte)(dec[i] ^ nonce[i % nonce.Length]);
                string gotPw = Encoding.UTF8.GetString(dec).TrimEnd('\0');
                if (gotPw != _expectedPassword)
                {
                    WriteErr(ns, 1045, "Access denied: RSA password mismatch", ref seq);
                    return;
                }
                WritePkt(ns, new byte[] { 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00 }, ref seq);
                AuthSucceeded = true;
            }
            else if (_authMode == "native")
            {
                byte[] newNonce = Encoding.ASCII.GetBytes("abcdefghij0123456789");
                using (MemoryStream ms = new MemoryStream())
                {
                    ms.WriteByte(0xFE);
                    byte[] p = Encoding.ASCII.GetBytes("mysql_native_password\0");
                    ms.Write(p, 0, p.Length);
                    ms.Write(newNonce, 0, newNonce.Length);
                    ms.WriteByte(0x00);
                    WritePkt(ns, ms.ToArray(), ref seq);
                }
                byte[] swResp = ReadPkt(ns, ref seq);
                using (SHA1 sha1 = SHA1.Create())
                {
                    byte[] s1 = sha1.ComputeHash(Encoding.UTF8.GetBytes(_expectedPassword));
                    byte[] s2 = sha1.ComputeHash(s1);
                    byte[] comb = Concat(newNonce, s2);
                    byte[] s3 = sha1.ComputeHash(comb);
                    byte[] exp = new byte[20];
                    for (int i = 0; i < 20; i++) exp[i] = (byte)(s1[i] ^ s3[i]);
                    bool ok = (swResp.Length == 20);
                    for (int i = 0; ok && i < 20; i++) if (swResp[i] != exp[i]) ok = false;
                    if (!ok) { WriteErr(ns, 1045, "Native scramble mismatch", ref seq); return; }
                }
                WritePkt(ns, new byte[] { 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00 }, ref seq);
                AuthSucceeded = true;
            }

            // Command loop
            while (!_stop)
            {
                byte[] cmd;
                try { cmd = ReadPkt(ns, ref seq); }
                catch { break; }
                if (cmd.Length == 0 || cmd[0] == 0x01) break; // COM_QUIT
                if (cmd[0] == 0x03)
                {
                    QueriesServed++;
                    string sql = Encoding.UTF8.GetString(cmd, 1, cmd.Length - 1).Trim();
                    lock (_sqlSync) { _executedSqls.Add(sql); }
                    if (sql.StartsWith("SET ", StringComparison.OrdinalIgnoreCase))
                    {
                        WritePkt(ns, new byte[] { 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00 }, ref seq);
                    }
                    else if (sql.IndexOf("SHOW FULL TABLES", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        SendResultSet(ns, new string[] { "Tables_in_mockdb", "Table_type" }, new string[][] {
                            new string[] { "mock_users", "BASE TABLE" },
                            new string[] { "mock_orders", "BASE TABLE" }
                        }, ref seq, false);
                    }
                    else if (sql.IndexOf("information_schema.tables", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        SendResultSet(ns, new string[] { "TABLE_NAME", "TABLE_ROWS" }, new string[][] {
                            new string[] { "mock_users", RowEstimate },
                            new string[] { "mock_orders", RowEstimate }
                        }, ref seq, false);
                    }
                    else if (sql.IndexOf("mock_users", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        SendResultSet(ns, new string[] { "id", "name", "plan", "memo" }, new string[][] {
                            new string[] { "1", "山田, 太郎", "PRO", null },
                            new string[] { "2", "佐藤 \"花子\"", "FREE", "line1\nline2" },
                            new string[] { "3", "鈴木 一郎", "PRO", "" }
                        }, ref seq, true);
                    }
                    else
                    {
                        SendResultSet(ns, new string[] { "order_id", "user_id", "amount", "status", "created_at" }, new string[][] {
                            new string[] { "101", "1", "12000", "PAID", "2026-10-01 10:15:00" },
                            new string[] { "102", "1", "3400", "PAID", "2026-10-02 14:30:00" },
                            new string[] { "103", "2", "500", "CANCELLED", "2026-10-03 14:45:00" }
                        }, ref seq, true);
                    }
                }
            }
        }

        private void SendResultSet(NetworkStream ns, string[] cols, string[][] rows, ref byte seq, bool applyDelay)
        {
            WritePkt(ns, new byte[] { (byte)cols.Length }, ref seq);
            foreach (string col in cols)
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    WriteLenStr(ms, "def"); WriteLenStr(ms, "mockdb"); WriteLenStr(ms, "t"); WriteLenStr(ms, "t");
                    WriteLenStr(ms, col); WriteLenStr(ms, col);
                    ms.WriteByte(0x0C);
                    ms.Write(new byte[] { 0xFF, 0x00, 0xFF, 0x00, 0x00, 0x00, 0xFD, 0x00, 0x00, 0x00, 0x00, 0x00 }, 0, 12);
                    WritePkt(ns, ms.ToArray(), ref seq);
                }
            }
            WritePkt(ns, new byte[] { 0xFE, 0x00, 0x00, 0x02, 0x00 }, ref seq);
            int count = rows.Length * (applyDelay ? DataRepeats : 1);
            for (int i = 0; i < count; i++)
            {
                string[] r = rows[i % rows.Length];
                if (applyDelay && RowDelayMs > 0) Thread.Sleep(RowDelayMs);
                using (MemoryStream ms = new MemoryStream())
                {
                    foreach (string cell in r)
                    {
                        if (cell == null && BinaryMemo) ms.Write(new byte[] { 2, 0xFF, 0x2C }, 0, 3);
                        else if (cell == null) ms.WriteByte(0xFB);
                        else WriteLenStr(ms, cell);
                    }
                    WritePkt(ns, ms.ToArray(), ref seq);
                }
            }
            WritePkt(ns, new byte[] { 0xFE, 0x00, 0x00, 0x02, 0x00 }, ref seq);
        }

        private static void WriteLenStr(Stream s, string v)
        {
            byte[] b = Encoding.UTF8.GetBytes(v);
            s.WriteByte((byte)b.Length);
            s.Write(b, 0, b.Length);
        }

        private static void WriteErr(NetworkStream ns, int code, string msg, ref byte seq)
        {
            byte[] mb = Encoding.UTF8.GetBytes("#28000" + msg);
            byte[] pkt = new byte[3 + mb.Length];
            pkt[0] = 0xFF; pkt[1] = (byte)(code & 0xFF); pkt[2] = (byte)(code >> 8);
            Array.Copy(mb, 0, pkt, 3, mb.Length);
            WritePkt(ns, pkt, ref seq);
        }

        private static byte[] ReadPkt(NetworkStream ns, ref byte seq)
        {
            byte[] hdr = new byte[4];
            ReadFull(ns, hdr, 4);
            int len = hdr[0] | (hdr[1] << 8) | (hdr[2] << 16);
            seq = (byte)(hdr[3] + 1);
            byte[] body = new byte[len];
            if (len > 0) ReadFull(ns, body, len);
            return body;
        }
        private static void ReadFull(NetworkStream ns, byte[] buf, int count)
        {
            int off = 0;
            while (off < count) { int r = ns.Read(buf, off, count - off); if (r <= 0) throw new EndOfStreamException(); off += r; }
        }
        private static void WritePkt(NetworkStream ns, byte[] body, ref byte seq)
        {
            byte[] hdr = new byte[] { (byte)(body.Length & 0xFF), (byte)((body.Length >> 8) & 0xFF), (byte)((body.Length >> 16) & 0xFF), seq++ };
            ns.Write(hdr, 0, 4);
            ns.Write(body, 0, body.Length);
            ns.Flush();
        }
        public void Dispose()
        {
            _stop = true;
            try { _listener.Stop(); } catch { }
            try { _rsa.Dispose(); } catch { }
        }
    }
}
'@
    if (-not ([System.Management.Automation.PSTypeName]'AuroraTestMock.MockMySql84Server').Type) {
        Add-Type -TypeDefinition $mockServerCs -ReferencedAssemblies @('System', 'System.Core')
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    foreach ($authMode in @('fast', 'rsa', 'native')) {
        $mock = New-Object AuroraTestMock.MockMySql84Server($authMode, 'SuperSecret84!')
        try {
            $mock.BinaryMemo = $authMode -eq 'fast'
            $expDir = Join-Path $workDir "mock_export_$authMode"
            $cfg = @{
                ServerName = '127.0.0.1'; Port = $mock.Port; Database = 'mockdb'; UserName = 'tester'; Password = 'SuperSecret84!'
                UseSshTunnel = $false; SshHost = ''; SshPort = 22; SshUser = ''; SshKeyPath = ''; SshPassphrase = ''
                SslMode = 'Disabled'; SslCaCert = ''
            }
            $exFiles = @(Invoke-AuroraBulkCsvExport -ConnConfig $cfg -ExportFolder $expDir)
            Assert-True ($mock.AuthSucceeded -and $exFiles.Count -eq 2) "MySQL 8.4.7 protocol ($authMode auth) connected and bulk-exported 2 tables cleanly"

            # Direct DB streaming analysis without exporting any CSV files
            $directAn = Invoke-AuroraDirectAnalysis -ConnConfig $cfg -DbLabel "Direct_$authMode"
            $csvAn = [AuroraAnalyzer.CsvAnalysisEngine]::AnalyzeFolder($expDir)
            $directUsers = $directAn.Tables | Where-Object { $_.TableName -eq 'mock_users' }
            $csvUsers = $csvAn.Tables | Where-Object { $_.TableName -eq 'mock_users' }
            Assert-True ($directAn.IsDirectDb -and $directAn.SourceMode -eq 'DirectDB' -and $directAn.TableCount -eq $csvAn.TableCount -and $directAn.TotalRows -eq $csvAn.TotalRows -and $directAn.TotalColumns -eq $csvAn.TotalColumns) "Direct DB analysis ($authMode auth) matches CSV analysis table/row/column counts exactly"
            Assert-True ($directAn.TotalFileBytes -eq $csvAn.TotalFileBytes -and $directAn.TotalEstInnoDbBytes -eq $csvAn.TotalEstInnoDbBytes -and $directUsers.FileSizeBytes -eq $csvUsers.FileSizeBytes -and $directUsers.EncodingName -eq 'UTF-8 (DB直接)' -and $directUsers.FileName -eq 'mockdb.mock_users') "Direct DB analysis ($authMode auth) matches CSV-equivalent byte size ($($directAn.TotalFileBytes)B) and InnoDB size with zero CSV files"

            if ($authMode -eq 'rsa') {
                $usersCsv = Get-Content -LiteralPath (Join-Path $expDir 'mock_users.csv') -Raw -Encoding UTF8
                Assert-True ($usersCsv -match '"山田, 太郎"' -and $usersCsv -match '"佐藤 ""花子"""') "Exported CSV properly RFC-4180 quotes Japanese commas & quotes"
            }
            if ($authMode -eq 'rsa' -and -not $PortableOnly) {
                # Full two-stage .bat -> .ps1 ExportAndAnalyze by A5M2 Registered DB Name!
                $mockRegConn = "__ConnectionType=Internal;ProviderName=MySQL;ServerName=127.0.0.1;Port=$($mock.Port);Database=mockdb;UserName=tester;UseSSL=False"
                New-ItemProperty -Path $dsKey -Name 'Mock_Aurora_84' -Value $mockRegConn -PropertyType String -Force | Out-Null
                $fullExpDir = Join-Path $workDir 'full_bat_export'
                $fullOutXlsx = Join-Path $workDir 'full_bat_report.xlsx'
                $fullOut = cmd.exe /c "`"$batPath`" -Mode ExportAndAnalyze -NoGui -A5DbName `"Mock_Aurora_84`" -A5RegistryRoot `"$testRegRoot`" -DbPassword `"SuperSecret84!`" -InputFolder `"$fullExpDir`" -OutputPath `"$fullOutXlsx`" -ExcelEngine OpenXml" 2>&1
                Assert-True ($LASTEXITCODE -eq 0 -and (Test-Path -LiteralPath (Join-Path $fullExpDir 'mock_users.csv')) -and (Test-Path -LiteralPath $fullOutXlsx)) "Full .bat ExportAndAnalyze via A5M2 registered DB name succeeded end-to-end"
                if ($LASTEXITCODE -ne 0) { Write-Host ($fullOut -join "`n") -ForegroundColor Red }

                # Full two-stage .bat -> .ps1 DirectAnalyze (NO -InputFolder, NO CSV files exported!)
                $directOutDir = Join-Path $workDir 'direct_bat_out'
                New-Item -ItemType Directory -Path $directOutDir -Force | Out-Null
                $directOutXlsx = Join-Path $directOutDir 'direct_bat_report.xlsx'
                $directOut = cmd.exe /c "`"$batPath`" -Mode DirectAnalyze -NoGui -A5DbName `"Mock_Aurora_84`" -A5RegistryRoot `"$testRegRoot`" -DbPassword `"SuperSecret84!`" -OutputPath `"$directOutXlsx`" -ExcelEngine OpenXml" 2>&1
                $directCsvCount = @(Get-ChildItem -LiteralPath $directOutDir -Filter '*.csv' -Recurse).Count
                Assert-True ($LASTEXITCODE -eq 0 -and (Test-Path -LiteralPath $directOutXlsx) -and $directCsvCount -eq 0) "Full .bat DirectAnalyze via A5M2 registered DB name succeeded with 0 CSV files created"
                if ($LASTEXITCODE -ne 0) { Write-Host ($directOut -join "`n") -ForegroundColor Red }

                # Verify DirectAnalyze workbook OOXML metadata & cross-mode parity with ExportAndAnalyze
                $dz = [System.IO.Compression.ZipFile]::OpenRead($directOutXlsx)
                try {
                    $ds1Sr = New-Object System.IO.StreamReader($dz.GetEntry('xl/worksheets/sheet1.xml').Open()); $ds1Text = $ds1Sr.ReadToEnd(); $ds1Sr.Dispose()
                    $ds2Sr = New-Object System.IO.StreamReader($dz.GetEntry('xl/worksheets/sheet2.xml').Open()); $ds2Text = $ds2Sr.ReadToEnd(); $ds2Sr.Dispose()
                    Assert-True ($ds1Text -match 'DB直接分析' -and $ds1Text -match 'Mock_Aurora_84' -and $ds2Text -match 'mockdb\.mock_users' -and $ds2Text -match 'UTF-8 \(DB直接\)') "DirectAnalyze workbook records DirectDB metadata and A5M2 registered DB label"
                }
                finally { $dz.Dispose() }

                $crossModeCmp = [AuroraAnalyzer.AnalysisComparisonEngine]::CompareAnalysisWorkbooks($fullOutXlsx, $directOutXlsx)
                Assert-True ($crossModeCmp.TotalUniqueTableCount -eq 2 -and $crossModeCmp.UnchangedTableCount -eq 2 -and $crossModeCmp.UpdatedSheetCount -eq 0 -and $crossModeCmp.AddedSheetCount -eq 0 -and $crossModeCmp.DeletedTableCount -eq 0) "CompareAnalysisWorkbooks(CSV ExportAndAnalyze vs DirectAnalyze) yields 0 false positives (2/2 unchanged)"
            }
            if ($authMode -eq 'native') {
                $mockList = Join-Path $workDir 'mock portable.a5dblist'
                [IO.File]::WriteAllLines($mockList, @('[DatabaseList]', "検証DB=ProviderName=MySQL;ServerName=127.0.0.1;Port=$($mock.Port);Database=mockdb;UserName=tester;UseSSL=False"), [Text.Encoding]::UTF8)
                $portableCsv = Join-Path $workDir 'portable_csv'
                $portableReport = Join-Path $workDir 'portable.xlsx'
                $portableOut = & $batPath -Mode ExportAndAnalyze -NoGui -A5DbListPath $mockList -A5DbName '検証DB' -DbPassword $cfg.Password -InputFolder $portableCsv -OutputPath $portableReport -ExcelEngine OpenXml
                Assert-True ($LASTEXITCODE -eq 0 -and @(Get-ChildItem -LiteralPath $portableCsv -Filter '*.csv').Count -eq 2 -and (Test-Path -LiteralPath $portableReport)) "Portable .bat export authenticates and creates two CSVs and an Excel report"
                if ($LASTEXITCODE -ne 0) { Write-Host ($portableOut -join "`n") -ForegroundColor Red }

                $portableDirectDir = Join-Path $workDir 'portable_direct_out'
                New-Item -ItemType Directory -Path $portableDirectDir -Force | Out-Null
                $portableDirectReport = Join-Path $portableDirectDir 'portable_direct.xlsx'
                $portableDirectOut = & $batPath -Mode DirectAnalyze -NoGui -A5DbListPath $mockList -A5DbName '検証DB' -DbPassword $cfg.Password -OutputPath $portableDirectReport -ExcelEngine OpenXml
                $portableDirectCsvs = @(Get-ChildItem -LiteralPath $portableDirectDir -Filter '*.csv' -Recurse).Count
                Assert-True ($LASTEXITCODE -eq 0 -and $portableDirectCsvs -eq 0 -and (Test-Path -LiteralPath $portableDirectReport)) "Portable .bat DirectAnalyze creates Excel report without InputFolder or CSV export"
                if ($LASTEXITCODE -ne 0) { Write-Host ($portableDirectOut -join "`n") -ForegroundColor Red }
                $crossMode = [AuroraAnalyzer.AnalysisComparisonEngine]::CompareAnalysisWorkbooks($portableDirectReport,$portableReport)
                Assert-True ($crossMode.TotalUniqueTableCount -eq 2 -and $crossMode.UnchangedTableCount -eq 2) "Direct and CSV-backed Excel reports have identical table and column statistics"
            }
        }
        finally { $mock.Dispose() }
    }

    Write-Host "`n=== 4B. Interruption (Cancel), Atomic .csv.part Cleanup & CSV Resume Tests ===" -ForegroundColor Cyan
    $mockResume = New-Object AuroraTestMock.MockMySql84Server('rsa', 'SuperSecret84!')
    try {
        $resumeCfg = @{
            ServerName = '127.0.0.1'; Port = $mockResume.Port; Database = 'mockdb'; UserName = 'tester'; Password = 'SuperSecret84!'
            UseSshTunnel = $false; SshHost = ''; SshPort = 22; SshUser = ''; SshKeyPath = ''; SshPassphrase = ''
            SslMode = 'Disabled'; SslCaCert = ''
        }

        # Baseline fresh export+analyze for parity verification
        $freshDir = Join-Path $workDir 'resume_fresh_csv'
        $freshXlsx = Join-Path $workDir 'resume_fresh.xlsx'
        $null = Invoke-AnalyzerPipeline -RunMode 'ExportAndAnalyze' -ConnConfig $resumeCfg -CsvFolder $freshDir -ExcelOutputPath $freshXlsx -Engine 'OpenXml'

        # Case A: Cancel mid-stream during 2nd table (mock_orders) after 1st table (mock_users) completes
        $resumeDir = Join-Path $workDir 'resume_interrupted_csv'
        $resumeXlsx = Join-Path $workDir 'resume_continued.xlsx'
        $mockResume.RowDelayMs = 80
        $mockResume.ClearExecutedSqls()
        $cancelState = New-Object AuroraAnalyzer.OperationCancelState
        $stateBox = @{ SawSecondTable = $false }
        $logCancel = {
            param($m, $p)
            if ($m -like '*エクスポート中: mock_orders*') { $stateBox.SawSecondTable = $true }
        }.GetNewClosure()
        $pumpCancel = {
            if ($stateBox.SawSecondTable -and (Test-Path -LiteralPath (Join-Path $resumeDir 'mock_orders.csv.part'))) {
                $cancelState.Cancel()
            }
        }.GetNewClosure()

        $caughtCancel = $false
        try {
            $null = Invoke-AnalyzerPipeline -RunMode 'ExportAndAnalyze' -ConnConfig $resumeCfg -CsvFolder $resumeDir -ExcelOutputPath $resumeXlsx -Engine 'OpenXml' -LogCallback $logCancel -CancelState $cancelState -PumpCallback $pumpCancel
        }
        catch {
            $caughtCancel = ($cancelState.IsCancelled) -or ($_.Exception -is [System.OperationCanceledException]) -or ($null -ne $_.Exception.InnerException -and $_.Exception.InnerException -is [System.OperationCanceledException])
        }
        Assert-True $caughtCancel "Mid-stream CSV export cancellation raises OperationCanceledException cleanly"
        $leftoverParts = @(Get-ChildItem -LiteralPath $resumeDir -Filter '*.csv.part' -ErrorAction SilentlyContinue)
        Assert-True ((Test-Path -LiteralPath (Join-Path $resumeDir 'mock_users.csv')) -and -not (Test-Path -LiteralPath (Join-Path $resumeDir 'mock_orders.csv')) -and $leftoverParts.Count -eq 0 -and -not (Test-Path -LiteralPath $resumeXlsx)) "Interrupted export preserves completed CSV, deletes incomplete .csv.part, and leaves no partial Excel"

        $progInterrupted = Read-ExportProgress -ExportFolder $resumeDir
        $sumInterrupted = Get-ExportFolderResumeSummary -ExportFolder $resumeDir
        Assert-True ($null -ne $progInterrupted -and $progInterrupted.Status -eq 'Interrupted' -and $progInterrupted.TotalTables -eq 2 -and $progInterrupted.CompletedTables.ContainsKey('mock_users') -and -not $progInterrupted.CompletedTables.ContainsKey('mock_orders')) "Export progress checkpoint records Interrupted status with mock_users completed"
        Assert-True ($null -ne $sumInterrupted -and $sumInterrupted.HasProgressFile -and $sumInterrupted.ProgressStatus -eq 'Interrupted' -and $sumInterrupted.RecordedDone -eq 1 -and $sumInterrupted.ExistingCsvCount -eq 1 -and $sumInterrupted.IncompletePartCount -eq 0) "Get-ExportFolderResumeSummary accurately detects resumable folder state"

        # Resume from interrupted folder via ExportAndAnalyze -Resume
        $mockResume.RowDelayMs = 0
        $mockResume.ClearExecutedSqls()
        [IO.File]::WriteAllText((Join-Path $resumeDir 'mock_orders.csv'), "old`r`nstale`r`n")
        [IO.File]::WriteAllText((Join-Path $resumeDir 'unrelated.csv'), "id`r`n9`r`n")
        $resumeLog = [Collections.Generic.List[string]]::new()
        $resumedRes = Invoke-AnalyzerPipeline -RunMode 'ExportAndAnalyze' -ConnConfig $resumeCfg -CsvFolder $resumeDir -ExcelOutputPath $resumeXlsx -Engine 'OpenXml' -Resume -LogCallback { param($message) $resumeLog.Add($message) }
        Assert-True (($resumeLog -match 'CSV解析中.*mock_users.csv').Count -eq 1 -and ($resumeLog -match '出力時の集計を使用.*mock_orders.csv').Count -eq 1) "Resume reads only skipped CSVs and uses the streamed profile for newly exported tables"
        $usersExpEntry = $resumedRes.ExportedTables | Where-Object { $_.TableName -eq 'mock_users' }
        $ordersExpEntry = $resumedRes.ExportedTables | Where-Object { $_.TableName -eq 'mock_orders' }
        Assert-True ($resumedRes.ExportedTables.Count -eq 2 -and $usersExpEntry.Skipped -and -not $ordersExpEntry.Skipped) "Resume skips already-exported mock_users and exports only remaining mock_orders"
        Assert-True ($mockResume.CountTableSelects('mock_users') -eq 0 -and $mockResume.CountTableSelects('mock_orders') -eq 1) "Resume issues 0 SELECT queries for skipped table and 1 SELECT for remaining table"
        $progCompleted = Read-ExportProgress -ExportFolder $resumeDir
        Assert-True ($null -ne $progCompleted -and $progCompleted.Status -eq 'Completed' -and $progCompleted.CompletedTables.Count -eq 2) "Progress checkpoint updates to Completed with all 2 tables after resume"
        $resumeParity = [AuroraAnalyzer.AnalysisComparisonEngine]::CompareAnalysisWorkbooks($freshXlsx, $resumeXlsx)
        Assert-True ((Test-Path -LiteralPath $resumeXlsx) -and $resumeParity.TotalUniqueTableCount -eq 2 -and $resumeParity.UnchangedTableCount -eq 2 -and $resumeParity.UpdatedSheetCount -eq 0 -and $resumeParity.AddedSheetCount -eq 0) "Resumed ExportAndAnalyze produces analysis Excel identical to fresh run (2/2 unchanged)"

        $mockResume.RowDelayMs = 80
        $mockResume.DataRepeats = 4
        $progressLog = New-Object 'System.Collections.Generic.List[string]'
        $percentTrace = [Collections.Generic.List[int]]::new()
        $recordProgress = {
            param($message, $pct, $csvPct = -1)
            $progressLog.Add($message)
            if ($csvPct -ge 0) { $percentTrace.Add($csvPct) }
        }
        $singleDir = Join-Path $workDir 'single_table'
        $single = Invoke-AnalyzerPipeline -RunMode ExportOnly -ConnConfig $resumeCfg -CsvFolder $singleDir -IncTables mock_users -LogCallback $recordProgress
        $progressText = $progressLog -join "`n"
        Assert-True ($progressText -match 'CSV出力\s+[1-9]\d?%.*件' -and $progressText -match 'CSV出力\s+100%') "A single slow table reports intermediate CSV percentages and reaches 100% after publication"
        Assert-True (($percentTrace -join ',') -ceq (($percentTrace | Sort-Object) -join ',') -and $percentTrace[0] -eq 0 -and @($percentTrace -eq 100).Count -eq 1 -and $percentTrace[-1] -eq 100) "An underestimated table advances monotonically and reports completion exactly once"
        Assert-True ($single.ExportedTables.Count -eq 1 -and $single.ExportedTables[0].RowCount -eq 12) "Single-table export returns a table collection and all rows despite a low estimate"
        $mockResume.RowDelayMs = 0
        $mockResume.DataRepeats = 1

        $beforeCheckpoint = [IO.File]::ReadAllText((Join-Path $resumeDir '_export_progress.json'))
        $mismatch = ''
        try { $null = Invoke-AuroraBulkCsvExport -ConnConfig $resumeCfg -ExportFolder $resumeDir -Resume -MaxRowsPerTable 1 }
        catch { $mismatch = $_.Exception.Message }
        Assert-True ($mismatch -match '上限' -and [IO.File]::ReadAllText((Join-Path $resumeDir '_export_progress.json')) -ceq $beforeCheckpoint) "Incompatible resume preserves the checkpoint and files"

        $oldFile = Join-Path $singleDir 'mock_users.csv'
        $oldHash = (Get-FileHash -LiteralPath $oldFile).Hash
        $stopOverwrite = [AuroraAnalyzer.OperationCancelState]::new()
        $stopAtPart = { if ([IO.File]::Exists($oldFile + '.part')) { $stopOverwrite.Cancel() } }.GetNewClosure()
        $mockResume.RowDelayMs = 120
        $stopped = $false
        try { $null = Invoke-AuroraBulkCsvExport -ConnConfig $resumeCfg -ExportFolder $singleDir -IncludeTables mock_users -CancelState $stopOverwrite -PumpCallback $stopAtPart }
        catch { $stopped = $stopOverwrite.IsCancelled }
        Assert-True ($stopped -and (Get-FileHash -LiteralPath $oldFile).Hash -ceq $oldHash -and -not [IO.File]::Exists($oldFile + '.part') -and (Get-ExportFolderResumeSummary $singleDir).RecordedDone -eq 0) "Interrupted overwrite preserves the previous CSV without marking it complete for this run"

        $mockResume.RowEstimate = $null
        $mockResume.RowDelayMs = 120
        $progressLog.Clear()
        $unknown = Invoke-AnalyzerPipeline -RunMode ExportOnly -ConnConfig $resumeCfg -CsvFolder $singleDir -IncTables mock_users -MaxRows 2 -LogCallback $recordProgress
        Assert-True (($progressLog -join "`n") -match '総件数不明' -and $unknown.ExportedTables[0].RowCount -eq 2) "Unknown estimate shows live row count and respects the row limit"
        $mockResume.RowDelayMs = 0
        $mockResume.RowEstimate = '3'
        $mockResume.DataRepeats = 0
        $emptyDir = Join-Path $workDir 'empty_table'
        $empty = Invoke-AnalyzerPipeline -RunMode ExportOnly -ConnConfig $resumeCfg -CsvFolder $emptyDir -IncTables mock_users
        $emptyDirect = Invoke-AuroraDirectAnalysis -ConnConfig $resumeCfg -IncludeTables mock_users
        $emptySize = [IO.FileInfo]::new($empty.ExportedTables[0].CsvPath).Length
        Assert-True ($emptySize -gt 3 -and $emptyDirect.TotalRows -eq 0 -and $emptyDirect.TotalFileBytes -eq $emptySize -and $emptyDirect.TotalColumns -eq 4) "Header-only CSV and direct analysis agree for an empty table"
        $mockResume.DataRepeats = 1

        $version1 = $beforeCheckpoint | ConvertFrom-Json
        $version1.Version = 1
        $version1.PSObject.Properties.Remove('Port')
        $manifest = Join-Path $resumeDir '_export_progress.json'
        [IO.File]::WriteAllText($manifest, ($version1 | ConvertTo-Json -Depth 6))
        [IO.File]::AppendAllText((Join-Path $resumeDir 'mock_orders.csv'), 'truncated')
        $repair = @(Invoke-AuroraBulkCsvExport -ConnConfig $resumeCfg -ExportFolder $resumeDir -Resume)
        Assert-True ($repair[0].Skipped -and -not $repair[1].Skipped -and (Read-ExportProgress $resumeDir).Version -eq 2) "Version 1 checkpoints remain usable and size mismatches are fetched again"
        [IO.File]::WriteAllText($manifest, '{broken')
        $invalid = ''
        try { $null = Invoke-AuroraBulkCsvExport -ConnConfig $resumeCfg -ExportFolder $resumeDir -Resume }
        catch { $invalid = $_.Exception.Message }
        Assert-True ($invalid -match '進捗記録' -and [IO.File]::ReadAllText($manifest) -ceq '{broken') "Damaged checkpoint is preserved and cannot silently adopt existing files"

        $reportCancel = [AuroraAnalyzer.OperationCancelState]::new()
        $cancelAtSave = { param($message, $pct) if ($pct -eq 98) { $reportCancel.Cancel() } }.GetNewClosure()
        $originalHash = (Get-FileHash -LiteralPath $freshXlsx).Hash
        $cancelledSave = $false
        try {
            $null = Invoke-AnalyzerPipeline -RunMode DirectAnalyze -ConnConfig $resumeCfg -ExcelOutputPath $freshXlsx -Engine OpenXml -LogCallback $cancelAtSave -CancelState $reportCancel
        }
        catch { $cancelledSave = $reportCancel.IsCancelled }
        Assert-True ($cancelledSave -and (Get-FileHash -LiteralPath $freshXlsx).Hash -ceq $originalHash -and @(Get-ChildItem -LiteralPath $workDir -Filter '.aurora-*.xlsx').Count -eq 0) "Cancelling DirectAnalyze during workbook save preserves the previous workbook and cleans staging"

        $reportCancel = [AuroraAnalyzer.OperationCancelState]::new()
        $cancelAtSave = { param($message, $pct) if ($pct -eq 98) { $reportCancel.Cancel() } }.GetNewClosure()
        $cancelledSave = $false
        try {
            $null = Invoke-AnalyzerPipeline -RunMode ExportAndAnalyze -ConnConfig $resumeCfg -CsvFolder $freshDir -ExcelOutputPath $freshXlsx -Engine OpenXml -Resume -LogCallback $cancelAtSave -CancelState $reportCancel
        }
        catch { $cancelledSave = $reportCancel.IsCancelled }
        Assert-True ($cancelledSave -and (Read-ExportProgress $freshDir).Status -eq 'CsvCompleted' -and (Get-FileHash -LiteralPath $freshXlsx).Hash -ceq $originalHash) "An interrupted Excel phase leaves a resumable checkpoint with all CSVs complete"
        $mockResume.ClearExecutedSqls()
        $null = Invoke-AnalyzerPipeline -RunMode ExportAndAnalyze -ConnConfig $resumeCfg -CsvFolder $freshDir -ExcelOutputPath $freshXlsx -Engine OpenXml -Resume
        Assert-True ((Read-ExportProgress $freshDir).Status -eq 'Completed' -and $mockResume.CountTableSelects('mock_users') -eq 0 -and $mockResume.CountTableSelects('mock_orders') -eq 0) "Resuming after report interruption publishes Excel without fetching any table again"

        # Case B: Resume from legacy folder WITHOUT _export_progress.json + 0-byte CSV + stray .csv.part
        $legacyDir = Join-Path $workDir 'resume_legacy_csv'
        New-Item -ItemType Directory -Path $legacyDir -Force | Out-Null
        Copy-Item -LiteralPath (Join-Path $freshDir 'mock_users.csv') -Destination (Join-Path $legacyDir 'mock_users.csv')
        [System.IO.File]::WriteAllBytes((Join-Path $legacyDir 'mock_orders.csv'), [byte[]]@())
        [System.IO.File]::WriteAllText((Join-Path $legacyDir 'mock_orders.csv.part'), 'partial,data')
        $mockResume.ClearExecutedSqls()
        $legacyExp = @(Invoke-AuroraBulkCsvExport -ConnConfig $resumeCfg -ExportFolder $legacyDir -Resume)
        $legacyUsers = $legacyExp | Where-Object { $_.TableName -eq 'mock_users' }
        $legacyOrders = $legacyExp | Where-Object { $_.TableName -eq 'mock_orders' }
        $legacyPartsAfter = @(Get-ChildItem -LiteralPath $legacyDir -Filter '*.csv.part' -ErrorAction SilentlyContinue)
        Assert-True ($legacyUsers.Skipped -and -not $legacyOrders.Skipped -and $mockResume.CountTableSelects('mock_users') -eq 0 -and $mockResume.CountTableSelects('mock_orders') -eq 1 -and $legacyPartsAfter.Count -eq 0 -and (Get-Item (Join-Path $legacyDir 'mock_orders.csv')).Length -gt 0) "Resume without checkpoint skips valid existing CSV, cleans .part, and re-exports 0-byte/incomplete CSV"

        # Case C: CLI two-stage .bat -Mode ExportAndAnalyze -NoGui -Resume
        $batResumeList = Join-Path $workDir 'bat_resume.a5dblist'
        [IO.File]::WriteAllLines($batResumeList, @('[DatabaseList]', "続行検証DB=ProviderName=MySQL;ServerName=127.0.0.1;Port=$($mockResume.Port);Database=mockdb;UserName=tester;UseSSL=False"), [Text.Encoding]::UTF8)
        $batResumeDir = Join-Path $workDir 'bat_resume_csv'
        $batResumeXlsx = Join-Path $workDir 'bat_resume.xlsx'
        New-Item -ItemType Directory -Path $batResumeDir -Force | Out-Null
        Copy-Item -LiteralPath (Join-Path $freshDir 'mock_users.csv') -Destination (Join-Path $batResumeDir 'mock_users.csv')
        $mockResume.ClearExecutedSqls()
        $batResOut = & $batPath -Mode ExportAndAnalyze -NoGui -A5DbListPath $batResumeList -A5DbName '続行検証DB' -DbPassword $resumeCfg.Password -InputFolder $batResumeDir -OutputPath $batResumeXlsx -ExcelEngine OpenXml -Resume
        Assert-True ($LASTEXITCODE -eq 0 -and $mockResume.CountTableSelects('mock_users') -eq 0 -and $mockResume.CountTableSelects('mock_orders') -eq 1 -and (Test-Path -LiteralPath $batResumeXlsx)) "CLI .bat -Mode ExportAndAnalyze -NoGui -Resume skips existing CSV and completes analysis"
        if ($LASTEXITCODE -ne 0) { Write-Host ($batResOut -join "`n") -ForegroundColor Red }
        Assert-True (($batResOut -join "`n") -match 'CSV出力 100%') "CLI prints CSV percentage even with redirected output"

        $guiDir = Join-Path $workDir 'gui_interrupt'
        $guiReport = Join-Path $workDir 'gui_interrupt.xlsx'
        $guiNotice = @{ Caption = ''; Text = ''; Progress = ''; LockDenied = $false }
        $recordNotice = { param($owner,$text,$caption,$icon) $guiNotice.Caption = $caption; $guiNotice.Text = $text }.GetNewClosure()
        $guiParams = @{ InitialMode = 'ExportAndAnalyze'; InitialExcelEngine = 'OpenXml'; InitialA5DbListPath = $batResumeList; InitialA5DbName = '続行検証DB'
            InitialInputFolder = $guiDir; InitialOutputPath = $guiReport; ShowMessage = $recordNotice }
        $gui = New-AnalyzerGuiForm @guiParams
        $timer = [Windows.Forms.Timer]::new()
        try {
            $gui.ShowInTaskbar = $false
            $gui.Opacity = 0
            $gui.Show()
            $dbGroup = $gui.Controls | Where-Object Text -like '3. A5:*'
            $passBox = $dbGroup.Controls | Where-Object { $_ -is [Windows.Forms.TextBox] -and $_.UseSystemPasswordChar } | Sort-Object Left | Select-Object -First 1
            $passBox.Text = $resumeCfg.Password
            ($gui.Controls | Where-Object Text -like '完了後にExcel*').Checked = $false
            $mockResume.RowDelayMs = 100
            $mockResume.DataRepeats = 3
            $cancelButton = $gui.Controls.Find('CancelButton',$true)[0]
            $progressLabel = $gui.Controls.Find('ProgressPercent',$true)[0]
            $timer.Interval = 25
            $timer.Add_Tick({
                if ($cancelButton.Enabled -and $progressLabel.Text -match 'CSV (?:[5-9]\d)%' -and [IO.File]::Exists((Join-Path $guiDir 'mock_orders.csv.part'))) {
                    $timer.Stop()
                    $guiNotice.Progress = $progressLabel.Text
                    try { $otherLock = Open-ExportFolderLock $guiDir; $otherLock.Dispose() }
                    catch { $guiNotice.LockDenied = $_.Exception.Message -match 'ロック' }
                    $cancelButton.PerformClick()
                }
            }.GetNewClosure())
            $timer.Start()
            ($gui.Controls | Where-Object Text -eq '実行開始').PerformClick()
            Assert-True ($guiNotice.Caption -eq '中断' -and $guiNotice.Progress -match 'CSV' -and $guiNotice.LockDenied -and (Get-ExportFolderResumeSummary $guiDir).RecordedDone -eq 1) "GUI cancel button interrupts a live transfer while the folder is locked"
        }
        finally { $timer.Dispose(); $gui.Dispose() }
        $mockResume.RowDelayMs = 0
        $mockResume.ClearExecutedSqls()
        $resumeForm = New-AnalyzerGuiForm @guiParams
        try {
            Assert-True ($resumeForm.Controls.Find('ResumeExport',$true)[0].Checked -and -not [IO.File]::Exists($guiReport) -and -not [IO.File]::Exists((Join-Path $guiDir 'mock_orders.csv.part'))) "Reopening the interrupted folder selects resume and leaves only completed data"
            $resumeForm.ShowInTaskbar = $false
            $resumeForm.Opacity = 0
            $resumeForm.Show()
            $dbGroup = $resumeForm.Controls | Where-Object Text -like '3. A5:*'
            $passBox = $dbGroup.Controls | Where-Object { $_ -is [Windows.Forms.TextBox] -and $_.UseSystemPasswordChar } | Sort-Object Left | Select-Object -First 1
            $passBox.Text = $resumeCfg.Password
            ($resumeForm.Controls | Where-Object Text -like '完了後にExcel*').Checked = $false
            ($resumeForm.Controls | Where-Object Text -eq '実行開始').PerformClick()
            Assert-True ($guiNotice.Caption -eq '完了' -and [IO.File]::Exists($guiReport) -and $mockResume.CountTableSelects('mock_users') -eq 0 -and $mockResume.CountTableSelects('mock_orders') -eq 1) "GUI resumes the selected folder, fetches only the unfinished table and publishes Excel ($($guiNotice.Text))"
            $resumeForm.Controls.Find('DirectAnalyzeMode',$true)[0].Checked = $true
            $paths = $resumeForm.Controls | Where-Object Text -like '2. 入力フォルダ*'
            $boxes = @($paths.Controls | Where-Object { $_ -is [Windows.Forms.TextBox] } | Sort-Object Top)
            $directDir = Join-Path $workDir 'gui_direct_only'
            $boxes[0].Text = ''
            $boxes[1].Text = Join-Path $directDir 'direct.xlsx'
            ($resumeForm.Controls | Where-Object Text -eq '実行開始').PerformClick()
            Assert-True ($guiNotice.Caption -eq '完了' -and [IO.File]::Exists($boxes[1].Text) -and @(Get-ChildItem -LiteralPath $directDir -File).Count -eq 1) "GUI DirectAnalyze publishes just the workbook with no input folder"
        }
        finally { $resumeForm.Dispose() }
        $mockResume.DataRepeats = 1

        # Case D: DirectAnalyze & AnalyzeOnly cancellation safety
        $cancelDirect = New-Object AuroraAnalyzer.OperationCancelState
        $cancelDirect.Cancel()
        $directCancelCaught = $false
        $cancelDirectXlsx = Join-Path $workDir 'cancelled_direct.xlsx'
        try {
            $null = Invoke-AnalyzerPipeline -RunMode 'DirectAnalyze' -ConnConfig $resumeCfg -ExcelOutputPath $cancelDirectXlsx -Engine 'OpenXml' -CancelState $cancelDirect
        }
        catch {
            $directCancelCaught = ($_.Exception -is [System.OperationCanceledException]) -or ($null -ne $_.Exception.InnerException -and $_.Exception.InnerException -is [System.OperationCanceledException])
        }
        Assert-True ($directCancelCaught -and -not (Test-Path -LiteralPath $cancelDirectXlsx)) "DirectAnalyze immediately honors OperationCancelState without creating output"

        $cancelAnalyze = New-Object AuroraAnalyzer.OperationCancelState
        $cancelAnalyze.Cancel()
        $analyzeCancelCaught = $false
        $cancelAnalyzeXlsx = Join-Path $workDir 'cancelled_analyze.xlsx'
        try {
            $null = Invoke-AnalyzerPipeline -RunMode 'AnalyzeOnly' -CsvFolder $freshDir -ExcelOutputPath $cancelAnalyzeXlsx -Engine 'OpenXml' -CancelState $cancelAnalyze
        }
        catch {
            $analyzeCancelCaught = ($_.Exception -is [System.OperationCanceledException]) -or ($null -ne $_.Exception.InnerException -and $_.Exception.InnerException -is [System.OperationCanceledException])
        }
        Assert-True ($analyzeCancelCaught -and -not (Test-Path -LiteralPath $cancelAnalyzeXlsx)) "AnalyzeOnly immediately honors OperationCancelState without creating output"
    }
    finally { $mockResume.Dispose() }

    if ($PortableOnly) {
        Write-Host "Portable verification: PASS = $passCount, FAIL = $failCount"
        if ($failCount -gt 0) { exit 1 }
        return
    }

    Write-Host "`n=== 5. Multi-Table Performance CSV Dataset Generation & Bat->PS1 End-to-End Execution ===" -ForegroundColor Cyan
    $csvDir = Join-Path $workDir 'aurora_csv_fixture'
    New-Item -ItemType Directory -Path $csvDir -Force | Out-Null

    # 5.1 users.csv (UTF-8 BOM, 200 rows)
    $uLines = New-Object System.Collections.Generic.List[string]
    $uLines.Add('id,user_code,plan,region,email,created_at')
    $plans = @('STANDARD','STANDARD','STANDARD','STANDARD','STANDARD','STANDARD','STANDARD','PRO','PRO','ENTERPRISE')
    $regions = @('Tokyo','Osaka','Nagoya','Fukuoka','Sapporo')
    for ($i = 1; $i -le 200; $i++) {
        $pl = $plans[($i - 1) % $plans.Count]
        $rg = $regions[($i - 1) % $regions.Count]
        $dt = (Get-Date '2026-01-01 09:00:00').AddDays($i % 180).AddHours($i % 12).ToString('yyyy-MM-dd HH:mm:ss')
        $uLines.Add("$i,USR$('{0:D5}' -f $i),$pl,$rg,user$i@example.jp,$dt")
    }
    [System.IO.File]::WriteAllLines((Join-Path $csvDir 'users.csv'), $uLines, (New-Object System.Text.UTF8Encoding($true)))

    # 5.2 orders.csv (UTF-8, 1200 rows: Whale user_id=1 has 350 orders, 5 orphan user_ids 9901..9905, peak hour 14:00, burst day, high NULL note, large JSON)
    $oLines = New-Object System.Collections.Generic.List[string]
    $oLines.Add('id,user_id,status,amount,created_at,note,payload_json')
    $statuses = @('PAID','PAID','PAID','PAID','PAID','PAID','PAID','PAID','PAID','PAID','PAID','PAID',
                  'SHIPPED','SHIPPED','SHIPPED','SHIPPED','SHIPPED',
                  'PENDING','PENDING',
                  'CANCELLED')
    $bigJson = '{"items":[' + ('{"sku":"A123","desc":"高性能テスト商品データ"},' * 120) + '{"sku":"END"}]}'
    $rng = New-Object System.Random 42
    for ($i = 1; $i -le 1200; $i++) {
        if ($i -le 350) { $uid = 1 }                    # Super-User / Whale! (350 orders -> severe skew + fanout!)
        elseif ($i -le 355) { $uid = 9900 + ($i - 350) } # 5 Orphan FK rows!
        else { $uid = 2 + (($i * 7) % 199) }

        $st = $statuses[($i - 1) % $statuses.Count]
        if ($i -le 80) { $amt = 95000 + ($i * 150) }    # Outliers
        elseif ($i -le 95) { $amt = -500 }              # Negative refunds
        elseif ($i -le 115) { $amt = 0 }                # Zero amounts
        else { $amt = 1000 + $rng.Next(0, 6000) }

        # Burst day 2026-08-15 (80 orders on one day!), peak hour 14:00
        if ($i -le 80) {
            $dtStr = (Get-Date '2026-08-15 14:00:00').AddMinutes($i % 59).ToString('yyyy-MM-dd HH:mm:ss')
        }
        else {
            $hr = if (($i % 3) -eq 0) { 14 } else { 8 + ($i % 12) }
            $dtStr = (Get-Date '2026-04-01 00:00:00').AddDays($i % 150).AddHours($hr).AddMinutes($i % 60).ToString('yyyy-MM-dd HH:mm:ss')
        }

        $note = if (($i % 10) -eq 0) { '特記事項あり' } else { '' } # 90% NULL
        $json = if ($i -eq 1) { '"' + ($bigJson -replace '"', '""') + '"' } else { '"{""ok"":true}"' }
        $oLines.Add("$i,$uid,$st,$amt,$dtStr,$note,$json")
    }
    [System.IO.File]::WriteAllLines((Join-Path $csvDir 'orders.csv'), $oLines, (New-Object System.Text.UTF8Encoding($false)))

    # 5.3 order_items.csv (Shift-JIS CP932! Multiline quoted fields + Japanese!)
    $oiLines = New-Object System.Collections.Generic.List[string]
    $oiLines.Add('id,order_id,product_id,qty,item_memo')
    for ($i = 1; $i -le 300; $i++) {
        $oid = 1 + (($i - 1) % 300)
        $prodId = 1 + (($i - 1) % 25)
        $qty = 1 + ($i % 5)
        $memo = if ($i -eq 5) { "`"配送先注意、`r`n午後指定`"" } else { '通常配送' }
        $oiLines.Add("$i,$oid,$prodId,$qty,$memo")
    }
    [System.IO.File]::WriteAllLines((Join-Path $csvDir 'order_items.csv'), $oiLines, [System.Text.Encoding]::GetEncoding(932))

    # 5.4 m_products.csv (Master table, 25 rows)
    $pLines = New-Object System.Collections.Generic.List[string]
    $pLines.Add('id,product_name,category,price,is_active')
    $cats = @('電子機器','書籍','食品','家電','雑貨')
    for ($i = 1; $i -le 25; $i++) {
        $pLines.Add("$i,商品_$i,$($cats[$i % 5]),$(500 * $i),$(if ($i -le 22) { 1 } else { 0 })")
    }
    [System.IO.File]::WriteAllLines((Join-Path $csvDir 'm_products.csv'), $pLines, (New-Object System.Text.UTF8Encoding($true)))

    # 5.5 audit_logs.csv (100 rows: 100% NULL col + Constant col + 1 ragged row!)
    $aLines = New-Object System.Collections.Generic.List[string]
    $aLines.Add('id,env_name,action_type,legacy_trace_id,logged_at')
    for ($i = 1; $i -le 100; $i++) {
        $act = if ($i % 4 -eq 0) { 'UPDATE' } else { 'SELECT' }
        $dt = (Get-Date '2026-06-01 10:00:00').AddHours($i).ToString('yyyy-MM-dd HH:mm:ss')
        if ($i -eq 50) { $aLines.Add("$i,prod,$act") } # ragged row!
        else { $aLines.Add("$i,prod,$act,,$dt") }
    }
    [System.IO.File]::WriteAllLines((Join-Path $csvDir 'audit_logs.csv'), $aLines, (New-Object System.Text.UTF8Encoding($true)))

    # 5.6 empty_staging.csv (0 rows -> Empty table high warning!)
    [System.IO.File]::WriteAllLines((Join-Path $csvDir 'empty_staging.csv'), @('id,batch_key,created_at'), (New-Object System.Text.UTF8Encoding($true)))

    # Direct analyzer ground-truth verification
    $an = [AuroraAnalyzer.CsvAnalysisEngine]::AnalyzeFolder($csvDir)
    Assert-True ($an.TableCount -eq 6) "Analyzed exact 6 CSV tables"
    Assert-True ($an.TotalRows -eq 1825) "Total rows exact match: $($an.TotalRows) == 1825"
    $oiTbl = $an.Tables | Where-Object { $_.TableName -eq 'order_items' }
    Assert-True ($oiTbl.EncodingName -eq 'Shift-JIS (CP932)' -and $oiTbl.RowCount -eq 300) "Auto-detected Shift-JIS (CP932) and parsed multiline quoted CSV accurately ($($oiTbl.EncodingName), $($oiTbl.RowCount) rows)"
    $audTbl = $an.Tables | Where-Object { $_.TableName -eq 'audit_logs' }
    Assert-True ($audTbl.RaggedRowCount -eq 1) "Detected exact 1 ragged CSV row in audit_logs"
    $ordRel = $an.Relationships | Where-Object { $_.ParentTable -eq 'users' -and $_.ChildTable -eq 'orders' } | Select-Object -First 1
    Assert-True ($null -ne $ordRel -and $ordRel.OrphanRows -eq 5 -and $ordRel.FanOutMax -eq 350) "Detected FK users.id->orders.user_id with 5 orphan rows & Whale FanOutMax=350"
    Assert-True (@($an.Warnings | Where-Object { $_.Category -eq 'データ不足' -and $_.TableName -eq 'empty_staging' }).Count -ge 1) "Detected Empty Table warning for empty_staging"
    Assert-True (@($an.Warnings | Where-Object { $_.Category -eq 'ホットキー・偏り' -and $_.TableName -eq 'orders' -and $_.ColumnName -eq 'user_id' }).Count -ge 1) "Detected Hot-Key Skew warning on orders.user_id"
    Assert-True (@($an.Warnings | Where-Object { $_.Category -eq '大容量ペイロード' -and $_.TableName -eq 'orders' -and $_.ColumnName -eq 'payload_json' }).Count -ge 1) "Detected Large Payload warning on orders.payload_json"
    Assert-True (@($an.Warnings | Where-Object { $_.Category -eq '全NULLカラム' -and $_.TableName -eq 'audit_logs' -and $_.ColumnName -eq 'legacy_trace_id' }).Count -ge 1) "Detected 100% NULL column warning on audit_logs.legacy_trace_id"

    # Execute TWO-STAGE .bat -> .ps1 invocation (Excel COM)
    $excelBefore = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $outCom = Join-Path $workDir 'Aurora_Report_COM.xlsx'
    $batCmd = "`"$batPath`" -Mode AnalyzeOnly -NoGui -InputFolder `"$csvDir`" -OutputPath `"$outCom`" -ExcelEngine Com"
    $batOut = cmd.exe /c $batCmd 2>&1
    $batExit = $LASTEXITCODE
    Assert-True ($batExit -eq 0 -and (Test-Path -LiteralPath $outCom)) "Two-stage Run_AuroraCsvAnalyzer.bat -> .ps1 (Excel COM) exited 0 and created report ($outCom)"
    if ($batExit -ne 0) { Write-Host ($batOut -join "`n") -ForegroundColor Red }

    Start-Sleep -Milliseconds 500
    $excelAfter = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $leakedExcel = @($excelAfter | Where-Object { $excelBefore -notcontains $_ })
    Assert-True ($leakedExcel.Count -eq 0) "Zero lingering EXCEL.EXE COM processes after generation"

    # Also test OpenXML Fallback engine via .bat
    $outXml = Join-Path $workDir 'Aurora_Report_OpenXML.xlsx'
    $batOut2 = cmd.exe /c "`"$batPath`" -Mode AnalyzeOnly -NoGui -InputFolder `"$csvDir`" -OutputPath `"$outXml`" -ExcelEngine OpenXml" 2>&1
    Assert-True ($LASTEXITCODE -eq 0 -and (Test-Path -LiteralPath $outXml)) "Two-stage Run_AuroraCsvAnalyzer.bat -> .ps1 (OpenXML fallback) exited 0 and created report"

    Write-Host "`n=== 6. Deep OOXML Verification (Sheets, Font Meiryo UI, Exact RGB #333333 / #1A2B4C / #800020, Charts) ===" -ForegroundColor Cyan
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    foreach ($pair in @(@{ Name = 'Excel COM'; Path = $outCom; MinCharts = 8 }, @{ Name = 'OpenXML Fallback'; Path = $outXml; MinCharts = 1 })) {
        $z = [System.IO.Compression.ZipFile]::OpenRead($pair.Path)
        try {
            $wbEntry = $z.GetEntry('xl/workbook.xml')
            $wbSr = New-Object System.IO.StreamReader($wbEntry.Open()); $wbText = $wbSr.ReadToEnd(); $wbSr.Dispose()
            $stEntry = $z.GetEntry('xl/styles.xml')
            $stSr = New-Object System.IO.StreamReader($stEntry.Open()); $stText = $stSr.ReadToEnd(); $stSr.Dispose()

            $expectedSheets = @('01_サマリー','02_テーブル一覧','03_カラム分析','04_分布・割合分析','05_数値統計・ヒストグラム','06_時系列分析','07_リレーション・ファンアウト','08_性能試験ガイド・警告')
            $allSheetsPresent = $true
            foreach ($sn in $expectedSheets) { if ($wbText -notmatch [regex]::Escape($sn)) { $allSheetsPresent = $false } }
            Assert-True $allSheetsPresent "[$($pair.Name)] All 8 Japanese worksheets exist in xl/workbook.xml"

            Assert-True ($stText -match 'Meiryo UI') "[$($pair.Name)] Meiryo UI font verified in xl/styles.xml"
            Assert-True ($stText -match '333333') "[$($pair.Name)] Exact Dark Gray (#333333) verified in xl/styles.xml"
            Assert-True ($stText -match '1A2B4C') "[$($pair.Name)] Exact Midnight Navy (#1A2B4C) verified in xl/styles.xml (NO BGR inversion!)"
            Assert-True ($stText -match '800020') "[$($pair.Name)] Exact Burgundy (#800020) verified in xl/styles.xml (NO BGR inversion!)"

            $chartEntries = @($z.Entries | Where-Object { $_.FullName -match '^xl/charts/chart\d+\.xml$' })
            Assert-True ($chartEntries.Count -ge $pair.MinCharts) "[$($pair.Name)] Native Excel charts embedded (count=$($chartEntries.Count) >= $($pair.MinCharts))"
            if ($chartEntries.Count -gt 0) {
                $cSr = New-Object System.IO.StreamReader($chartEntries[0].Open()); $cText = $cSr.ReadToEnd(); $cSr.Dispose()
                Assert-True ($cText -match 'Meiryo UI') "[$($pair.Name)] Embedded chart XML uses Meiryo UI font"
            }
        }
        finally { $z.Dispose() }
    }

    Write-Host "`n=== 7. Windows Forms GUI Mode Changes & Run Button ===" -ForegroundColor Cyan
    $guiReport = Join-Path $workDir 'Aurora_Report_GUI.xlsx'
    $notice = @{ Caption = ''; Text = '' }
    $recordNotice = {
        param($owner, $text, $caption, $icon)
        $notice.Caption = $caption
        $notice.Text = $text
    }.GetNewClosure()
    $guiForm = New-AnalyzerGuiForm -InitialRegistryRoot $testRegRoot -InitialInputFolder $csvDir -InitialOutputPath $guiReport -ShowMessage $recordNotice
    try {
        Assert-True ($null -ne $guiForm -and $guiForm.Font.Name -eq 'Meiryo UI') "GUI dialog instantiates cleanly with Meiryo UI font"
        Assert-True ($guiForm.Controls.Count -ge 6) "GUI dialog contains all mode, path, A5M2 DB, progress, and action controls"

        $modeGroup = $guiForm.Controls | Where-Object Text -eq '1. 実行モード選択'
        $directOption = $guiForm.Controls.Find('DirectAnalyzeMode', $true)[0]
        $fullOption = $modeGroup.Controls | Where-Object Text -like 'A5M2*'
        $exportOption = $modeGroup.Controls | Where-Object Text -eq 'CSV一括出力のみ'
        $analyzeOption = $modeGroup.Controls | Where-Object Text -like '既存CSV*'
        $pathGroup = $guiForm.Controls | Where-Object Text -like '2. 入力フォルダ*'
        $inputField = $pathGroup.Controls | Where-Object Text -eq $csvDir
        $outputField = $pathGroup.Controls | Where-Object Text -eq $guiReport
        $dbGroup = $guiForm.Controls | Where-Object Text -like '3. A5:*'
        $openReport = $guiForm.Controls | Where-Object Text -like '完了後にExcel*'
        $resumeChk = $guiForm.Controls.Find('ResumeExport', $true)[0]
        $runButton = $guiForm.Controls | Where-Object Text -eq '実行開始'
        $cancelButton = $guiForm.Controls.Find('CancelButton', $true)[0]

        $guiForm.Show() # Button.PerformClick() requires a visible, selectable control.
        $openReport.Checked = $false
        Assert-True ($null -ne $resumeChk -and $null -ne $cancelButton -and -not $cancelButton.Enabled) "GUI includes ResumeExport checkbox and initially disabled CancelButton"
        Assert-True ($analyzeOption.Checked -and -not $dbGroup.Enabled -and $outputField.Enabled -and -not $resumeChk.Enabled) "AnalyzeOnly is selected initially with ResumeExport disabled"
        $directOption.Checked = $true
        Assert-True ($dbGroup.Visible -and $dbGroup.Enabled -and -not $inputField.Enabled -and $outputField.Enabled -and $openReport.Enabled -and -not $resumeChk.Enabled) "DirectAnalyze enables DB and report settings while disabling CSV folder input and ResumeExport"
        $fullOption.Checked = $true
        Assert-True ($dbGroup.Enabled -and $inputField.Enabled -and $outputField.Enabled -and $openReport.Enabled -and $resumeChk.Enabled) "ExportAndAnalyze enables DB, folder, report settings, and ResumeExport"
        $exportOption.Checked = $true
        Assert-True ($dbGroup.Enabled -and $inputField.Enabled -and -not $outputField.Enabled -and -not $openReport.Enabled -and $resumeChk.Enabled) "ExportOnly disables report settings while keeping ResumeExport enabled"
        $analyzeOption.Checked = $true
        Assert-True (-not $dbGroup.Enabled -and $inputField.Enabled -and $outputField.Enabled -and $openReport.Enabled -and -not $resumeChk.Enabled) "Switching back to AnalyzeOnly restores report settings and disables ResumeExport"

        $inputField.Text = ''
        $runButton.PerformClick()
        Assert-True ($notice.Caption -eq '入力チェック') "Run with an empty input folder shows the validation message"
        $inputField.Text = $csvDir
        $runButton.PerformClick()
        Assert-True ($notice.Caption -eq '完了' -and (Test-Path -LiteralPath $guiReport) -and $runButton.Enabled -and -not $cancelButton.Enabled) "Run analyzes the CSV fixture and creates the report ($($notice.Text))"
    }
    finally { $guiForm.Dispose() }

    # GUI DirectAnalyze end-to-end test against live mock MySQL 8.4.7 server (no CSV output!)
    $guiDirectDir = Join-Path $workDir 'gui_direct_out'
    New-Item -ItemType Directory -Path $guiDirectDir -Force | Out-Null
    $guiDirectReport = Join-Path $guiDirectDir 'Aurora_Direct_GUI.xlsx'
    $guiDirectMock = New-Object AuroraTestMock.MockMySql84Server('rsa', 'SuperSecret84!')
    try {
        $guiDirectList = Join-Path $workDir 'gui_direct.a5dblist'
        [IO.File]::WriteAllLines($guiDirectList, @('[DatabaseList]', "直接検証DB=ProviderName=MySQL;ServerName=127.0.0.1;Port=$($guiDirectMock.Port);Database=mockdb;UserName=tester;UseSSL=False"), [Text.Encoding]::UTF8)
        $directNotice = @{ Caption = ''; Text = '' }
        $recordDirectNotice = {
            param($owner, $text, $caption, $icon)
            $directNotice.Caption = $caption
            $directNotice.Text = $text
        }.GetNewClosure()
        $directForm = New-AnalyzerGuiForm -InitialRegistryRoot $testRegRoot -InitialA5DbListPath $guiDirectList -InitialA5DbName '直接検証DB' -InitialOutputPath '' -InitialMode 'DirectAnalyze' -ShowMessage $recordDirectNotice
        try {
            $dRadio = $directForm.Controls.Find('DirectAnalyzeMode', $true)[0]
            $dPathGroup = $directForm.Controls | Where-Object Text -like '2. 入力フォルダ*'
            $dPathBoxes = @($dPathGroup.Controls | Where-Object { $_ -is [Windows.Forms.TextBox] } | Sort-Object Top)
            $dInBox = $dPathBoxes[0]
            $dOutBox = $dPathBoxes[1]
            $dDbGroup = $directForm.Controls | Where-Object Text -like '3. A5:*'
            $dDbPassBox = $dDbGroup.Controls | Where-Object { $_ -is [Windows.Forms.TextBox] -and $_.UseSystemPasswordChar } | Sort-Object Left | Select-Object -First 1
            $dOpenChk = $directForm.Controls | Where-Object Text -like '完了後にExcel*'
            $dRunBtn = $directForm.Controls | Where-Object Text -eq '実行開始'

            $directForm.Show()
            $dOpenChk.Checked = $false
            Assert-True ($dRadio.Checked -and $dDbGroup.Enabled -and -not $dInBox.Enabled -and $dOutBox.Enabled) "GUI initializes cleanly in DirectAnalyze mode with CSV folder disabled"

            $dRunBtn.PerformClick()
            Assert-True ($directNotice.Caption -eq '入力チェック') "GUI DirectAnalyze with empty output path shows validation message"

            $dOutBox.Text = $guiDirectReport
            $dDbPassBox.Text = 'SuperSecret84!'
            $dRunBtn.PerformClick()
            $guiDirectCsvs = @(Get-ChildItem -LiteralPath $guiDirectDir -Filter '*.csv' -Recurse).Count
            Assert-True ($directNotice.Caption -eq '完了' -and (Test-Path -LiteralPath $guiDirectReport) -and $guiDirectCsvs -eq 0 -and $dRunBtn.Enabled) "GUI DirectAnalyze connects via A5M2 profile and creates Excel report with 0 CSV files ($($directNotice.Text))"
        }
        finally { $directForm.Dispose() }
    }
    finally { $guiDirectMock.Dispose() }

    Write-Host "`n=== 8. Before/After Analysis File Comparison Engine, CLI (.bat), OOXML & GUI Tests ===" -ForegroundColor Cyan
    # Build Before fixture (add legacy_temp.csv so we also test deleted table detection)
    $cmpBeforeDir = Join-Path $workDir 'cmp_before_csv'
    $cmpAfterDir  = Join-Path $workDir 'cmp_after_csv'
    New-Item -ItemType Directory -Path $cmpBeforeDir -Force | Out-Null
    New-Item -ItemType Directory -Path $cmpAfterDir  -Force | Out-Null

    Copy-Item -LiteralPath (Join-Path $csvDir 'users.csv')         -Destination (Join-Path $cmpBeforeDir 'users.csv')
    Copy-Item -LiteralPath (Join-Path $csvDir 'orders.csv')        -Destination (Join-Path $cmpBeforeDir 'orders.csv')
    Copy-Item -LiteralPath (Join-Path $csvDir 'order_items.csv')   -Destination (Join-Path $cmpBeforeDir 'order_items.csv')
    Copy-Item -LiteralPath (Join-Path $csvDir 'm_products.csv')    -Destination (Join-Path $cmpBeforeDir 'm_products.csv')
    Copy-Item -LiteralPath (Join-Path $csvDir 'audit_logs.csv')    -Destination (Join-Path $cmpBeforeDir 'audit_logs.csv')
    Copy-Item -LiteralPath (Join-Path $csvDir 'empty_staging.csv') -Destination (Join-Path $cmpBeforeDir 'empty_staging.csv')
    [System.IO.File]::WriteAllLines((Join-Path $cmpBeforeDir 'legacy_temp.csv'), @('id,code', '1,OLD_A', '2,OLD_B'), (New-Object System.Text.UTF8Encoding($true)))

    # 8.1 Unchanged tables in After: m_products.csv (25 rows) & empty_staging.csv (0 rows)
    Copy-Item -LiteralPath (Join-Path $cmpBeforeDir 'm_products.csv')    -Destination (Join-Path $cmpAfterDir 'm_products.csv')
    Copy-Item -LiteralPath (Join-Path $cmpBeforeDir 'empty_staging.csv') -Destination (Join-Path $cmpAfterDir 'empty_staging.csv')

    # 8.2 In-place UPDATED table (same 200 rows, plan & created_at modified -> 更新（データ更新）)
    $uAfterLines = [System.IO.File]::ReadAllLines((Join-Path $cmpBeforeDir 'users.csv'), (New-Object System.Text.UTF8Encoding($true)))
    for ($i = 1; $i -le 20; $i++) {
        $uAfterLines[$i] = $uAfterLines[$i] -replace ',STANDARD,', ',PRO,'
        $uAfterLines[$i] = $uAfterLines[$i] -replace '2026-\d\d-\d\d \d\d:\d\d:\d\d$', '2026-11-24 23:55:00'
    }
    [System.IO.File]::WriteAllLines((Join-Path $cmpAfterDir 'users.csv'), $uAfterLines, (New-Object System.Text.UTF8Encoding($true)))

    # 8.3 Pure row-ADDED table (order_items: 300 -> 400 rows, +100 pure INSERTs -> 追加（レコード増加）)
    $oiAfterLines = New-Object System.Collections.Generic.List[string]
    foreach ($ln in [System.IO.File]::ReadAllLines((Join-Path $cmpBeforeDir 'order_items.csv'), [System.Text.Encoding]::GetEncoding(932))) {
        $oiAfterLines.Add($ln)
    }
    for ($i = 301; $i -le 400; $i++) {
        $oiAfterLines.Add("$i,$((1 + ($i % 300))),$((1 + ($i % 25))),$((1 + ($i % 5))),通常配送")
    }
    [System.IO.File]::WriteAllLines((Join-Path $cmpAfterDir 'order_items.csv'), $oiAfterLines, [System.Text.Encoding]::GetEncoding(932))

    # 8.4 Row-ADDED + In-Place UPDATED table (orders: 1200 -> 1250 rows AND existing PENDING updated to PAID -> 追加・更新)
    $oAfterLines = New-Object System.Collections.Generic.List[string]
    $pendingConverted = 0
    foreach ($ln in [System.IO.File]::ReadAllLines((Join-Path $cmpBeforeDir 'orders.csv'), (New-Object System.Text.UTF8Encoding($false)))) {
        if ($pendingConverted -lt 30 -and $ln -match ',PENDING,') {
            $oAfterLines.Add(($ln -replace ',PENDING,', ',PAID,'))
            $pendingConverted++
        }
        else {
            $oAfterLines.Add($ln)
        }
    }
    for ($i = 1201; $i -le 1250; $i++) {
        $oAfterLines.Add("$i,10,PAID,3200,2026-11-24 15:00:00,,`"{`"`"ok`"`":true}`"")
    }
    [System.IO.File]::WriteAllLines((Join-Path $cmpAfterDir 'orders.csv'), $oAfterLines, (New-Object System.Text.UTF8Encoding($false)))

    # 8.5 Row-DECREASED table (audit_logs: 100 -> 70 rows -> 更新（レコード減少） + Alert!)
    $aBeforeLines = [System.IO.File]::ReadAllLines((Join-Path $cmpBeforeDir 'audit_logs.csv'), (New-Object System.Text.UTF8Encoding($true)))
    [System.IO.File]::WriteAllLines((Join-Path $cmpAfterDir 'audit_logs.csv'), ($aBeforeLines | Select-Object -First 71), (New-Object System.Text.UTF8Encoding($true)))

    # 8.6 Brand-new table (new_batch_results: not in Before, 45 rows in After -> 追加（新規テーブル）)
    $nbLines = New-Object System.Collections.Generic.List[string]
    $nbLines.Add('id,job_name,elapsed_ms,executed_at')
    for ($i = 1; $i -le 45; $i++) {
        $nbLines.Add("$i,NIGHTLY_BATCH,$(120 + $i * 5),2026-11-24 02:00:00")
    }
    [System.IO.File]::WriteAllLines((Join-Path $cmpAfterDir 'new_batch_results.csv'), $nbLines, (New-Object System.Text.UTF8Encoding($true)))
    # Note: legacy_temp.csv is intentionally absent in After -> 削除（テーブル消失） + Alert!

    # Generate Before (both COM & OpenXML) and After (COM & OpenXML) analysis workbooks
    $beforeComXlsx = Join-Path $workDir 'Before_Analysis_COM.xlsx'
    $beforeXmlXlsx = Join-Path $workDir 'Before_Analysis_XML.xlsx'
    $afterComXlsx  = Join-Path $workDir 'After_Analysis_COM.xlsx'
    $afterXmlXlsx  = Join-Path $workDir 'After_Analysis_XML.xlsx'

    $null = Invoke-AnalyzerPipeline -RunMode 'AnalyzeOnly' -CsvFolder $cmpBeforeDir -ExcelOutputPath $beforeComXlsx -Engine 'Com'
    $null = Invoke-AnalyzerPipeline -RunMode 'AnalyzeOnly' -CsvFolder $cmpBeforeDir -ExcelOutputPath $beforeXmlXlsx -Engine 'OpenXml'
    $null = Invoke-AnalyzerPipeline -RunMode 'AnalyzeOnly' -CsvFolder $cmpAfterDir  -ExcelOutputPath $afterComXlsx  -Engine 'Com'
    $null = Invoke-AnalyzerPipeline -RunMode 'AnalyzeOnly' -CsvFolder $cmpAfterDir  -ExcelOutputPath $afterXmlXlsx  -Engine 'OpenXml'

    # 8.7 Cross-engine identity test (Before COM vs Before OpenXML -> 0 false positives!)
    $sameCmp = [AuroraAnalyzer.AnalysisComparisonEngine]::CompareAnalysisWorkbooks($beforeComXlsx, $beforeXmlXlsx)
    Assert-True ($sameCmp.TotalUniqueTableCount -eq 7 -and $sameCmp.UnchangedTableCount -eq 7 -and $sameCmp.UpdatedSheetCount -eq 0 -and $sameCmp.AddedSheetCount -eq 0 -and $sameCmp.DeletedTableCount -eq 0) "Cross-engine identical comparison (COM vs OpenXML) yields 0 false positives (7/7 unchanged)"

    # 8.8 Full Before vs After classification accuracy verification
    $cmpRes = [AuroraAnalyzer.AnalysisComparisonEngine]::CompareAnalysisWorkbooks($beforeComXlsx, $afterComXlsx)
    Assert-True ($cmpRes.TotalUniqueTableCount -eq 8) "Comparison detects exact 8 unique tables (7 before, 7 after)"
    $tProducts = $cmpRes.AllTables | Where-Object { $_.TableName -eq 'm_products' }
    $tEmpty    = $cmpRes.AllTables | Where-Object { $_.TableName -eq 'empty_staging' }
    $tUsers    = $cmpRes.AllTables | Where-Object { $_.TableName -eq 'users' }
    $tItems    = $cmpRes.AllTables | Where-Object { $_.TableName -eq 'order_items' }
    $tOrders   = $cmpRes.AllTables | Where-Object { $_.TableName -eq 'orders' }
    $tAudit    = $cmpRes.AllTables | Where-Object { $_.TableName -eq 'audit_logs' }
    $tNew      = $cmpRes.AllTables | Where-Object { $_.TableName -eq 'new_batch_results' }
    $tLegacy   = $cmpRes.AllTables | Where-Object { $_.TableName -eq 'legacy_temp' }

    Assert-True ($tProducts.PrimaryStatus -eq '変更なし' -and $tProducts.IsUnchanged) "m_products classified as 変更なし despite cross-table RowShare/SizeShare shifts"
    Assert-True ($tEmpty.PrimaryStatus -eq '変更なし' -and $tEmpty.IsUnchanged) "empty_staging (0->0 rows) classified as 変更なし"
    Assert-True ($tUsers.PrimaryStatus -eq '更新（データ更新）' -and $tUsers.IsUpdated -and $tUsers.RowDelta -eq 0 -and $tUsers.ChangedColumnCount -ge 2) "users (same 200 rows, plan & created_at updated) classified as 更新（データ更新）"
    Assert-True ($tItems.PrimaryStatus -eq '追加（レコード増加）' -and $tItems.IsAdded -and -not $tItems.IsUpdated -and $tItems.RowDelta -eq 100) "order_items (+100 pure inserts) classified as 追加（レコード増加）"
    Assert-True ($tOrders.PrimaryStatus -eq '追加・更新' -and $tOrders.IsAdded -and $tOrders.IsUpdated -and $tOrders.RowDelta -eq 50) "orders (+50 inserts + existing PENDING->PAID update) classified as 追加・更新"
    Assert-True ($tAudit.PrimaryStatus -eq '更新（レコード減少）' -and $tAudit.IsUpdated -and $tAudit.HasAlert -and $tAudit.RowDelta -eq -30) "audit_logs (-30 rows) classified as 更新（レコード減少） with Burgundy alert"
    Assert-True ($tNew.PrimaryStatus -eq '追加（新規テーブル）' -and $tNew.IsNewTable -and $tNew.IsAdded -and $tNew.AfterRows -eq 45) "new_batch_results classified as 追加（新規テーブル）"
    Assert-True ($tLegacy.PrimaryStatus -eq '削除（テーブル消失）' -and $tLegacy.IsDeleted -and $tLegacy.HasAlert) "legacy_temp classified as 削除（テーブル消失） with Burgundy alert"
    Assert-True ($cmpRes.UnchangedTables.Count -eq 2 -and $cmpRes.AddedTables.Count -eq 3 -and $cmpRes.UpdatedTables.Count -eq 3 -and $cmpRes.DeletedTables.Count -eq 1) "Sheet bucket counts exact (Unchanged=2, Added=3, Updated=3, Deleted=1)"

    # 8.9 Negative validation tests (invalid extension / missing workbook sheets)
    $badXlsx = Join-Path $workDir 'not_analysis.xlsx'
    $badZip = [System.IO.Compression.ZipFile]::Open($badXlsx, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        $e1 = $badZip.CreateEntry('xl/workbook.xml'); $sw = New-Object System.IO.StreamWriter($e1.Open()); $sw.Write('<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'); $sw.Dispose()
        $e2 = $badZip.CreateEntry('xl/_rels/workbook.xml.rels'); $sw2 = New-Object System.IO.StreamWriter($e2.Open()); $sw2.Write('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'); $sw2.Dispose()
    }
    finally { $badZip.Dispose() }
    $badErr = ''
    try { $null = [AuroraAnalyzer.AnalysisComparisonEngine]::CompareAnalysisWorkbooks($badXlsx, $afterComXlsx) }
    catch { $badErr = $_.Exception.Message }
    Assert-True ($badErr -match '02_テーブル一覧') "Invalid analysis workbook produces clear Japanese diagnostic error ($badErr)"

    # 8.10 Two-stage .bat -> .ps1 -Mode CompareAnalysis (both COM & OpenXML)
    $excelBeforeCmp = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $diffComOut = Join-Path $workDir 'Diff_Report_COM.xlsx'
    $diffXmlOut = Join-Path $workDir 'Diff_Report_OpenXML.xlsx'

    $batCmpCom = cmd.exe /c "`"$batPath`" -Mode CompareAnalysis -NoGui -BeforeAnalysisPath `"$beforeComXlsx`" -AfterAnalysisPath `"$afterComXlsx`" -OutputPath `"$diffComOut`" -ExcelEngine Com" 2>&1
    Assert-True ($LASTEXITCODE -eq 0 -and (Test-Path -LiteralPath $diffComOut)) "Two-stage .bat CompareAnalysis (Excel COM) exited 0 and created comparison report"
    if ($LASTEXITCODE -ne 0) { Write-Host ($batCmpCom -join "`n") -ForegroundColor Red }

    Start-Sleep -Milliseconds 500
    $excelAfterCmp = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $leakedCmpExcel = @($excelAfterCmp | Where-Object { $excelBeforeCmp -notcontains $_ })
    Assert-True ($leakedCmpExcel.Count -eq 0) "Zero lingering EXCEL.EXE COM processes after comparison report generation"

    $batCmpXml = cmd.exe /c "`"$batPath`" -Mode CompareAnalysis -NoGui -BeforeAnalysisPath `"$beforeXmlXlsx`" -AfterAnalysisPath `"$afterXmlXlsx`" -OutputPath `"$diffXmlOut`" -ExcelEngine OpenXml" 2>&1
    Assert-True ($LASTEXITCODE -eq 0 -and (Test-Path -LiteralPath $diffXmlOut)) "Two-stage .bat CompareAnalysis (OpenXML fallback) exited 0 and created comparison report"
    if ($LASTEXITCODE -ne 0) { Write-Host ($batCmpXml -join "`n") -ForegroundColor Red }

    # 8.11 Deep OOXML verification of Comparison Excel reports (Sheets, Meiryo UI, #333333 / #1A2B4C / #800020, Charts, Cell Contents)
    $expectedCmpSheets = @('01_比較サマリー', '02_全テーブル比較一覧', '03_更新のあったテーブル', '04_追加のあったテーブル', '05_変更のなかったテーブル', '06_カラム・統計差分詳細')
    foreach ($pair in @(@{ Name = 'Compare COM'; Path = $diffComOut; MinCharts = 4 }, @{ Name = 'Compare OpenXML'; Path = $diffXmlOut; MinCharts = 1 })) {
        $z = [System.IO.Compression.ZipFile]::OpenRead($pair.Path)
        try {
            $wbSr = New-Object System.IO.StreamReader($z.GetEntry('xl/workbook.xml').Open()); $wbText = $wbSr.ReadToEnd(); $wbSr.Dispose()
            $stSr = New-Object System.IO.StreamReader($z.GetEntry('xl/styles.xml').Open()); $stText = $stSr.ReadToEnd(); $stSr.Dispose()

            $allCmpSheets = $true
            foreach ($sn in $expectedCmpSheets) { if ($wbText -notmatch [regex]::Escape($sn)) { $allCmpSheets = $false } }
            Assert-True $allCmpSheets "[$($pair.Name)] All 6 Japanese comparison worksheets exist in xl/workbook.xml"
            Assert-True ($stText -match 'Meiryo UI') "[$($pair.Name)] Meiryo UI font verified in xl/styles.xml"
            Assert-True ($stText -match '333333') "[$($pair.Name)] Exact Dark Gray (#333333) verified in xl/styles.xml"
            Assert-True ($stText -match '1A2B4C') "[$($pair.Name)] Exact Midnight Navy (#1A2B4C) verified in xl/styles.xml"
            Assert-True ($stText -match '800020') "[$($pair.Name)] Exact Burgundy (#800020) verified in xl/styles.xml"

            $chartEntries = @($z.Entries | Where-Object { $_.FullName -match '^xl/charts/chart\d+\.xml$' })
            Assert-True ($chartEntries.Count -ge $pair.MinCharts) "[$($pair.Name)] Native comparison charts embedded (count=$($chartEntries.Count) >= $($pair.MinCharts))"
            if ($chartEntries.Count -gt 0) {
                $cSr = New-Object System.IO.StreamReader($chartEntries[0].Open()); $cText = $cSr.ReadToEnd(); $cSr.Dispose()
                Assert-True ($cText -match 'Meiryo UI') "[$($pair.Name)] Comparison chart XML uses Meiryo UI font"
            }

            # Verify sheet content routing (updated/added/unchanged tables appear in their expected sheets)
            $allXmlText = ''
            foreach ($ze in $z.Entries | Where-Object { $_.FullName -match '^xl/(sharedStrings|worksheets/sheet[3456])\.xml$' }) {
                $sr = New-Object System.IO.StreamReader($ze.Open()); $allXmlText += $sr.ReadToEnd(); $sr.Dispose()
            }
            Assert-True ($allXmlText -match 'users' -and $allXmlText -match 'new_batch_results' -and $allXmlText -match 'm_products' -and $allXmlText -match 'legacy_temp') "[$($pair.Name)] All categorized table names present in comparison workbook XML"
        }
        finally { $z.Dispose() }
    }

    # 8.12 Windows Forms GUI CompareAnalysis mode end-to-end test
    $guiDiffOut = Join-Path $workDir 'Diff_Report_GUI.xlsx'
    $cmpNotice = @{ Caption = ''; Text = '' }
    $recordCmpNotice = {
        param($owner, $text, $caption, $icon)
        $cmpNotice.Caption = $caption
        $cmpNotice.Text = $text
    }.GetNewClosure()
    $cmpForm = New-AnalyzerGuiForm -InitialRegistryRoot $testRegRoot -InitialBeforeAnalysisPath $beforeComXlsx -InitialAfterAnalysisPath $afterComXlsx -InitialOutputPath $guiDiffOut -ShowMessage $recordCmpNotice
    try {
        $cmpRadio  = $cmpForm.Controls.Find('CompareMode', $true)[0]
        $cmpGroup  = $cmpForm.Controls.Find('CompareGroup', $true)[0]
        $beforeBox = $cmpForm.Controls.Find('BeforeAnalysisPath', $true)[0]
        $afterBox  = $cmpForm.Controls.Find('AfterAnalysisPath', $true)[0]
        $swapBtn   = $cmpForm.Controls.Find('SwapCompareFiles', $true)[0]
        $dbGrp     = $cmpForm.Controls | Where-Object Text -like '3. A5:*'
        $openChk   = $cmpForm.Controls | Where-Object Text -like '完了後にExcel*'
        $runBtn    = $cmpForm.Controls | Where-Object Text -eq '実行開始'

        $cmpForm.Show()
        $openChk.Checked = $false
        Assert-True ($cmpRadio.Checked -and $cmpGroup.Visible -and -not $dbGrp.Visible) "GUI initializes in CompareAnalysis mode when Before/After analysis paths are given"

        $swapBtn.PerformClick()
        Assert-True ($beforeBox.Text -eq $afterComXlsx -and $afterBox.Text -eq $beforeComXlsx) "Swap button cleanly exchanges Before and After analysis file paths"
        $swapBtn.PerformClick()

        $beforeBox.Text = ''
        $runBtn.PerformClick()
        Assert-True ($cmpNotice.Caption -eq '入力チェック') "CompareAnalysis run with missing Before file shows validation message"

        $beforeBox.Text = $beforeComXlsx
        $runBtn.PerformClick()
        Assert-True ($cmpNotice.Caption -eq '完了' -and (Test-Path -LiteralPath $guiDiffOut) -and $runBtn.Enabled) "GUI CompareAnalysis run creates comparison Excel report cleanly"
    }
    finally { $cmpForm.Dispose() }

    Write-Host "`n=== 9. WCAG Contrast, Monotone Design, Chart Integrity & Hardening Regressions ===" -ForegroundColor Cyan

    function Get-HexContrastRatio {
        param([string]$Hex1, [string]$Hex2)
        function Get-RelLum([string]$h) {
            if ([string]::IsNullOrEmpty($h)) { return 1.0 }
            $s = $h.TrimStart('#').ToUpperInvariant()
            if ($s.Length -eq 8) { $s = $s.Substring(2) }
            if ($s.Length -ne 6) { return 0.0 }
            $r = [Convert]::ToInt32($s.Substring(0, 2), 16) / 255.0
            $g = [Convert]::ToInt32($s.Substring(2, 2), 16) / 255.0
            $b = [Convert]::ToInt32($s.Substring(4, 2), 16) / 255.0
            $lr = if ($r -le 0.03928) { $r / 12.92 } else { [Math]::Pow(($r + 0.055) / 1.055, 2.4) }
            $lg = if ($g -le 0.03928) { $g / 12.92 } else { [Math]::Pow(($g + 0.055) / 1.055, 2.4) }
            $lb = if ($b -le 0.03928) { $b / 12.92 } else { [Math]::Pow(($b + 0.055) / 1.055, 2.4) }
            return (0.2126 * $lr) + (0.7152 * $lg) + (0.0722 * $lb)
        }
        $l1 = Get-RelLum $Hex1
        $l2 = Get-RelLum $Hex2
        $hi = [Math]::Max($l1, $l2)
        $lo = [Math]::Min($l1, $l2)
        return ($hi + 0.05) / ($lo + 0.05)
    }

    function Test-WorkbookVisualAudit {
        param([string]$Label, [string]$XlsxPath, [bool]$IsAnalysisWorkbook = $true, [bool]$IsComWorkbook = $true)
        $z = [System.IO.Compression.ZipFile]::OpenRead($XlsxPath)
        try {
            $nsm = New-Object System.Xml.XmlNamespaceManager (New-Object System.Xml.NameTable)
            $nsm.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
            $nsm.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
            $nsm.AddNamespace('c', 'http://schemas.openxmlformats.org/drawingml/2006/chart')

            # Shared strings
            $sst = @()
            $ssEntry = $z.GetEntry('xl/sharedStrings.xml')
            if ($null -ne $ssEntry) {
                $ssDoc = New-Object System.Xml.XmlDocument
                $sr = New-Object System.IO.StreamReader($ssEntry.Open()); $ssDoc.LoadXml($sr.ReadToEnd()); $sr.Dispose()
                foreach ($si in $ssDoc.SelectNodes('//x:si', $nsm)) {
                    $sst += $si.InnerText
                }
            }

            # Styles
            $stDoc = New-Object System.Xml.XmlDocument
            $sr = New-Object System.IO.StreamReader($z.GetEntry('xl/styles.xml').Open()); $stText = $sr.ReadToEnd(); $sr.Dispose()
            $stDoc.LoadXml($stText)

            $fonts = @()
            foreach ($fn in $stDoc.SelectNodes('//x:fonts/x:font', $nsm)) {
                $nameNode = $fn.SelectSingleNode('x:name', $nsm)
                $colNode  = $fn.SelectSingleNode('x:color', $nsm)
                $rgb = if ($null -ne $colNode -and $colNode.HasAttribute('rgb')) { $colNode.GetAttribute('rgb').ToUpperInvariant() } else { 'FF000000' }
                if ($rgb.Length -eq 8) { $rgb = $rgb.Substring(2) }
                $fonts += [PSCustomObject]@{
                    Name = if ($null -ne $nameNode) { $nameNode.GetAttribute('val') } else { '' }
                    Rgb  = $rgb
                }
            }

            $fills = @()
            foreach ($fl in $stDoc.SelectNodes('//x:fills/x:fill', $nsm)) {
                $pf = $fl.SelectSingleNode('x:patternFill', $nsm)
                $pt = if ($null -ne $pf) { $pf.GetAttribute('patternType') } else { 'none' }
                $fg = if ($null -ne $pf) { $pf.SelectSingleNode('x:fgColor', $nsm) } else { $null }
                $rgb = 'FFFFFF'
                if ($pt -eq 'solid' -and $null -ne $fg -and $fg.HasAttribute('rgb')) {
                    $raw = $fg.GetAttribute('rgb').ToUpperInvariant()
                    if ($raw -ne '00000000') { $rgb = if ($raw.Length -eq 8) { $raw.Substring(2) } else { $raw } }
                }
                $fills += [PSCustomObject]@{ Pattern = $pt; Rgb = $rgb }
            }

            $xfs = @()
            foreach ($xf in $stDoc.SelectNodes('//x:cellXfs/x:xf', $nsm)) {
                $xfs += [PSCustomObject]@{
                    FontId = if ($xf.HasAttribute('fontId')) { [int]$xf.GetAttribute('fontId') } else { 0 }
                    FillId = if ($xf.HasAttribute('fillId')) { [int]$xf.GetAttribute('fillId') } else { 0 }
                }
            }

            Assert-True ($stText -notmatch 'F9ECEF|EAEDF2') "[$Label] Zero banned same-color-family tint fills (#F9ECEF / #EAEDF2) in xl/styles.xml"

            $lowContrastCount = 0
            $nonMeiryoCount   = 0
            $whiteOnNavyCount = 0
            $burgundyCount    = 0
            $falseOptimalWarn = 0
            $brokenFmtCount   = 0
            $has00TimeText    = $false
            $coercedTimeFloat = 0

            # Map sheet names to targets
            $wbDoc = New-Object System.Xml.XmlDocument
            $sr = New-Object System.IO.StreamReader($z.GetEntry('xl/workbook.xml').Open()); $wbDoc.LoadXml($sr.ReadToEnd()); $sr.Dispose()
            $relDoc = New-Object System.Xml.XmlDocument
            $sr = New-Object System.IO.StreamReader($z.GetEntry('xl/_rels/workbook.xml.rels').Open()); $relDoc.LoadXml($sr.ReadToEnd()); $sr.Dispose()
            $relMap = @{}
            foreach ($rn in $relDoc.DocumentElement.ChildNodes) {
                $relMap[$rn.GetAttribute('Id')] = $rn.GetAttribute('Target').TrimStart('/')
            }

            foreach ($shNode in $wbDoc.SelectNodes('//x:sheets/x:sheet', $nsm)) {
                $shName = $shNode.GetAttribute('name')
                $rid = $shNode.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
                $target = $relMap[$rid]
                if (-not $target.StartsWith('xl/')) { $target = 'xl/' + $target }
                $shEntry = $z.GetEntry($target)
                if ($null -eq $shEntry) { continue }
                $shDoc = New-Object System.Xml.XmlDocument
                $sr = New-Object System.IO.StreamReader($shEntry.Open()); $shDoc.LoadXml($sr.ReadToEnd()); $sr.Dispose()

                foreach ($cNode in $shDoc.SelectNodes('//x:sheetData/x:row/x:c', $nsm)) {
                    $cType = $cNode.GetAttribute('t')
                    $valStr = ''
                    if ($cType -eq 'inlineStr') {
                        $valStr = $cNode.InnerText
                    }
                    elseif ($cType -eq 's') {
                        $vNode = $cNode.SelectSingleNode('x:v', $nsm)
                        if ($null -ne $vNode) {
                            $siIdx = [int]$vNode.InnerText
                            if ($siIdx -ge 0 -and $siIdx -lt $sst.Count) { $valStr = $sst[$siIdx] }
                        }
                    }
                    else {
                        $vNode = $cNode.SelectSingleNode('x:v', $nsm)
                        if ($null -ne $vNode) { $valStr = $vNode.InnerText }
                    }
                    if ([string]::IsNullOrEmpty($valStr)) { continue }

                    $sIdx = if ($cNode.HasAttribute('s')) { [int]$cNode.GetAttribute('s') } else { 0 }
                    if ($sIdx -ge 0 -and $sIdx -lt $xfs.Count) {
                        $xf = $xfs[$sIdx]
                        $fn = $fonts[$xf.FontId]
                        $fl = $fills[$xf.FillId]
                        if (-not [string]::IsNullOrEmpty($fn.Name) -and $fn.Name -ne 'Meiryo UI') { $nonMeiryoCount++ }
                        $cr = Get-HexContrastRatio $fn.Rgb $fl.Rgb
                        if ($cr -lt 4.5) { $lowContrastCount++ }
                        if ($fn.Rgb -eq 'FFFFFF' -and $fl.Rgb -eq '1A2B4C') { $whiteOnNavyCount++ }
                        if ($fn.Rgb -eq '800020') {
                            $burgundyCount++
                            if ($valStr -match '最適') { $falseOptimalWarn++ }
                        }
                    }

                    if ($valStr -match '(^|\s)02 MB\b|111倍|\b61/51\b|^#(NAME\?|VALUE!|REF!|DIV/0!)$') {
                        $brokenFmtCount++
                    }
                    if ($shName -eq '06_時系列分析') {
                        if ($valStr -eq '00:00') { $has00TimeText = $true }
                        $cRef = $cNode.GetAttribute('r')
                        if ($cRef -match '^A\d+$' -and [string]::IsNullOrEmpty($cType)) {
                            $numTmp = 0.0
                            if ([double]::TryParse($valStr, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$numTmp) -and $numTmp -ge 0.0 -and $numTmp -lt 1.0) {
                                $coercedTimeFloat++
                            }
                        }
                    }
                }
            }

            Assert-True ($nonMeiryoCount -eq 0) "[$Label] 100% of populated cells use Meiryo UI font (non-Meiryo=$nonMeiryoCount)"
            Assert-True ($lowContrastCount -eq 0) "[$Label] 100% of populated cells meet WCAG AA contrast >= 4.5:1 (lowContrast=$lowContrastCount)"
            Assert-True ($whiteOnNavyCount -gt 0) "[$Label] Crisp White (#FFFFFF) on Midnight Navy (#1A2B4C) headers verified ($whiteOnNavyCount cells)"
            Assert-True ($burgundyCount -gt 0) "[$Label] Burgundy (#800020) warning cells actually applied ($burgundyCount cells)"
            Assert-True ($falseOptimalWarn -eq 0) "[$Label] Zero false Burgundy warnings on positive ratings like '最適 (高並列対応)'"
            Assert-True ($brokenFmtCount -eq 0) "[$Label] Zero broken C# format artifacts ('02 MB', '111倍', '61/51') or Excel errors"

            if ($IsAnalysisWorkbook -and $IsComWorkbook) {
                Assert-True ($has00TimeText -and $coercedTimeFloat -eq 0) "[$Label] 06_時系列分析 preserves '00:00' text verbatim without Excel float coercion"
            }

            if ($IsComWorkbook) {
                $brokenCharts = 0
                foreach ($ce in $z.Entries | Where-Object { $_.FullName -match '^xl/charts/chart\d+\.xml$' }) {
                    $cDoc = New-Object System.Xml.XmlDocument
                    $sr = New-Object System.IO.StreamReader($ce.Open()); $cDoc.LoadXml($sr.ReadToEnd()); $sr.Dispose()
                    $sers = $cDoc.SelectNodes('//c:ser', $nsm)
                    if ($sers.Count -ge 4) {
                        $singleCellCount = 0
                        foreach ($sn in $sers) {
                            $fNode = $sn.SelectSingleNode('c:val//c:f', $nsm)
                            if ($null -ne $fNode -and $fNode.InnerText -notmatch ':') { $singleCellCount++ }
                        }
                        if ($singleCellCount -eq $sers.Count) { $brokenCharts++ }
                    }
                }
                Assert-True ($brokenCharts -eq 0) "[$Label] Zero PowerShell COM Range-unrolled single-cell series charts"
            }
        }
        finally { $z.Dispose() }
    }

    Test-WorkbookVisualAudit -Label 'Analysis COM'    -XlsxPath $outCom     -IsAnalysisWorkbook $true  -IsComWorkbook $true
    Test-WorkbookVisualAudit -Label 'Analysis OpenXML' -XlsxPath $outXml    -IsAnalysisWorkbook $true  -IsComWorkbook $false
    Test-WorkbookVisualAudit -Label 'Compare COM'     -XlsxPath $diffComOut -IsAnalysisWorkbook $false -IsComWorkbook $true
    Test-WorkbookVisualAudit -Label 'Compare OpenXML' -XlsxPath $diffXmlOut -IsAnalysisWorkbook $false -IsComWorkbook $false

    # 9.1 Analyzer accuracy regressions (junction table role, dotted email, single-digit slash date, new table AlertNote)
    Assert-True ($oiTbl.InferredRole -eq '中間・関連テーブル') "order_items (2 FK columns, 5 cols) inferred as 中間・関連テーブル (got: $($oiTbl.InferredRole))"
    Assert-True ($tNew.AlertNote -eq '正常') "New table without warnings has AlertNote='正常' (got: '$($tNew.AlertNote)')"

    $edgeCsvDir = Join-Path $workDir 'edge_types_csv'
    New-Item -ItemType Directory -Path $edgeCsvDir -Force | Out-Null
    [System.IO.File]::WriteAllLines((Join-Path $edgeCsvDir 'contacts.csv'), @(
        'id,email,created_at',
        '1,taro.yamada@example.jp,2026/1/5 9:05:00',
        '2,hanako.sato@sub.example.co.jp,2026/12/31 23:59:59'
    ), (New-Object System.Text.UTF8Encoding($true)))
    $edgeAn = [AuroraAnalyzer.CsvAnalysisEngine]::AnalyzeFolder($edgeCsvDir)
    $edgeEmailCol = $edgeAn.Tables[0].Columns | Where-Object { $_.ColumnName -eq 'email' }
    $edgeDateCol  = $edgeAn.Tables[0].Columns | Where-Object { $_.ColumnName -eq 'created_at' }
    Assert-True ($edgeEmailCol.InferredType -eq 'メールアドレス') "Dotted email (taro.yamada@example.jp) inferred as メールアドレス (got: $($edgeEmailCol.InferredType))"
    Assert-True ($edgeDateCol.InferredType -eq '日時 (DateTime)') "Single-digit slash datetime (2026/1/5 9:05:00) inferred as 日時 (DateTime) (got: $($edgeDateCol.InferredType))"

    # 9.2 Atomic comparison export preservation on cancel + relative OutputPath resolution
    $atomicTarget = Join-Path $workDir 'existing_compare.xlsx'
    [System.IO.File]::WriteAllText($atomicTarget, 'PREVIOUS_GOOD_CONTENT')
    $preCancelled = New-Object AuroraAnalyzer.OperationCancelState
    $preCancelled.Cancel()
    $cancelThrew = $false
    try {
        $null = Export-ComparisonReport -Comparison $cmpRes -OutputPath $atomicTarget -Engine 'OpenXml' -CancelState $preCancelled
    }
    catch [System.OperationCanceledException] { $cancelThrew = $true }
    Assert-True ($cancelThrew -and ([System.IO.File]::ReadAllText($atomicTarget) -eq 'PREVIOUS_GOOD_CONTENT')) "Cancelled Export-ComparisonReport leaves pre-existing output file completely intact (atomic staging)"

    Push-Location $workDir
    try {
        $relName = 'relative_cmp_out.xlsx'
        $null = Export-ComparisonReport -Comparison $cmpRes -OutputPath $relName -Engine 'OpenXml'
        Assert-True (Test-Path -LiteralPath (Join-Path $workDir $relName)) "Export-ComparisonReport resolves relative -OutputPath cleanly"
    }
    finally { Pop-Location }

    # 9.3 SSH_ASKPASS batch metacharacter escaping & ProcessOutputDrain non-blocking drain
    $tmpAskCmd = Join-Path $workDir 'test_askpass_safe.cmd'
    [System.IO.File]::WriteAllText($tmpAskCmd, "@echo off`r`nsetlocal EnableDelayedExpansion`r`necho(!A5_SSH_SECRET!`r`n", [System.Text.Encoding]::ASCII)
    $nastySecret = 'p@ss&word|test<1>2^3%PATH%!var!(paren)"quote&&echo HACKED'
    $apPsi = New-Object System.Diagnostics.ProcessStartInfo
    $apPsi.FileName = 'cmd.exe'
    $apPsi.Arguments = '/c "' + $tmpAskCmd + '"'
    $apPsi.UseShellExecute = $false
    $apPsi.RedirectStandardOutput = $true
    $apPsi.RedirectStandardError = $true
    $apPsi.CreateNoWindow = $true
    $apPsi.EnvironmentVariables['A5_SSH_SECRET'] = $nastySecret
    $apProc = [System.Diagnostics.Process]::Start($apPsi)
    $apOut = $apProc.StandardOutput.ReadToEnd().TrimEnd("`r", "`n")
    $apErr = $apProc.StandardError.ReadToEnd()
    $apProc.WaitForExit()
    Assert-True ($apOut -ceq $nastySecret -and [string]::IsNullOrEmpty($apErr)) "SSH_ASKPASS delayed-expansion batch preserves metacharacters without shell injection"

    $drPsi = New-Object System.Diagnostics.ProcessStartInfo
    $drPsi.FileName = 'cmd.exe'
    $drPsi.Arguments = '/c "for /L %i in (1,1,400) do @echo STDOUT_LINE_%i && echo STDERR_MARKER 1>&2"'
    $drPsi.UseShellExecute = $false
    $drPsi.RedirectStandardOutput = $true
    $drPsi.RedirectStandardError = $true
    $drPsi.CreateNoWindow = $true
    $drProc = [System.Diagnostics.Process]::Start($drPsi)
    $drDrain = New-Object AuroraAnalyzer.ProcessOutputDrain
    $drDrain.Attach($drProc)
    $drExited = $drProc.WaitForExit(15000)
    if ($drExited) { $drProc.WaitForExit() }
    $drErr = $drDrain.GetStderrText()
    Assert-True ($drExited -and $drErr -match 'STDERR_MARKER') "ProcessOutputDrain drains high-volume stdout/stderr asynchronously without pipe deadlock"

    Write-Host "`n============================================================" -ForegroundColor Cyan
    Write-Host "VERIFICATION SUMMARY: PASS = $passCount, FAIL = $failCount" -ForegroundColor $(if ($failCount -eq 0) { 'Green' } else { 'Red' })
    Write-Host "============================================================" -ForegroundColor Cyan

    if ($failCount -gt 0) { exit 1 }
}
finally {
    if (-not $PortableOnly -and (Test-Path -LiteralPath $testRegRoot)) { Remove-Item -LiteralPath $testRegRoot -Recurse -Force -ErrorAction SilentlyContinue }
    if (Test-Path -LiteralPath $workDir) { Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue }
}
