# ============================================================================
# BOJClientVB 単体テスト用スタブサーバ
# ----------------------------------------------------------------------------
# 用途:
#   - AAVerifyID の異常系テスト（szRet = 2 / 3 / 4）を再現する
#   - 実サーバ（xinetd）を止めずに、クライアント側 PC 上で動作する
#
# 起動例:
#   PowerShell> .\stub_server.ps1 -Mode normal-ack
#   PowerShell> .\stub_server.ps1 -Mode ng
#   PowerShell> .\stub_server.ps1 -Mode malformed-commas
#   PowerShell> .\stub_server.ps1 -Mode malformed-header
#   PowerShell> .\stub_server.ps1 -Mode silent
#   PowerShell> .\stub_server.ps1 -Mode normal-ack -Port 50000
#
# 管理者権限: ポートバインド・ファイアウォール許可のため PowerShell は
#             「管理者として実行」で起動してください。
# ============================================================================

param(
    [ValidateSet(
        "normal-ack",       # 正常: "ACK user, pass, 1" + LF
        "normal-ack-closed",# 正常閉局: "ACK user, pass, 0" + LF
        "ng",               # 認証失敗: "NG" + LF  → szRet="3"
        "malformed-commas", # カンマ不足: "ACK onlyone" + LF → szRet="3"
        "malformed-header", # ヘッダ不正: "XYZ foo,bar,1" + LF → szRet="3"
        "silent"            # 無応答: accept のみで応答せず → szRet="4"
    )]
    [string]$Mode = "normal-ack",

    [int]$Port = 50000,

    [int]$SilentSleepSec = 120
)

$lf = [char]10

# モードに応じて返す電文を決定
function Get-Response([string]$mode) {
    switch ($mode) {
        "normal-ack"        { return "ACK testuser, AAAAA, 1" + $lf }
        "normal-ack-closed" { return "ACK testuser, AAAAA, 0" + $lf }
        "ng"                { return "NG" + $lf }
        "malformed-commas"  { return "ACK onlyone" + $lf }
        "malformed-header"  { return "XYZ foo,bar,1" + $lf }
        "silent"            { return $null }
        default             { return "NG" + $lf }
    }
}

$listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Any, $Port)
try {
    $listener.Start()
} catch {
    Write-Host "ERROR: ポート $Port を listen できません。管理者権限 / ファイアウォール / ポート競合をご確認ください。" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}

Write-Host "===============================================" -ForegroundColor Cyan
Write-Host " BOJClientVB Stub Server"                    -ForegroundColor Cyan
Write-Host "   Mode = $Mode"                              -ForegroundColor Cyan
Write-Host "   Port = $Port"                              -ForegroundColor Cyan
Write-Host "   Ctrl+C で終了"                             -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan

$connCount = 0

try {
    while ($true) {
        $client = $listener.AcceptTcpClient()
        $connCount++
        $remote = $client.Client.RemoteEndPoint
        Write-Host "[$connCount] Accepted from $remote (mode=$Mode)" -ForegroundColor Yellow

        try {
            $stream = $client.GetStream()

            if ($Mode -eq "silent") {
                # accept のみ。受信も応答もせずスリープ
                Write-Host "    ... silent mode: sleep $SilentSleepSec sec"
                Start-Sleep -Seconds $SilentSleepSec
            } else {
                # 要求受信（最大 1024 byte）
                $buf = New-Object byte[] 1024
                $n = 0
                $stream.ReadTimeout = 5000
                try {
                    $n = $stream.Read($buf, 0, $buf.Length)
                } catch {
                    Write-Host "    ... read failed (timeout or closed): $($_.Exception.Message)"
                }
                if ($n -gt 0) {
                    $req = [System.Text.Encoding]::ASCII.GetString($buf, 0, $n)
                    Write-Host "    REQ : $req".TrimEnd()
                }

                $response = Get-Response $Mode
                if ($response -ne $null) {
                    $bytes = [System.Text.Encoding]::ASCII.GetBytes($response)
                    $stream.Write($bytes, 0, $bytes.Length)
                    $stream.Flush()
                    Write-Host "    RES : $response".TrimEnd()
                }
            }
        } catch {
            Write-Host "    ERROR: $($_.Exception.Message)" -ForegroundColor Red
        } finally {
            try { $client.Close() } catch {}
        }
    }
} finally {
    $listener.Stop()
    Write-Host "Listener stopped."
}
