Imports System.Net.Sockets
Imports System.Text
Imports Microsoft.Win32

'***********************************************************************
'* BOJClientVBCLS --- BOJクライアント通信クラス（COM公開）
'*  [概要]
'*      従来 Access 32bit 環境で利用していた BOJClient.ocx の機能を
'*      Access 64bit から呼び出せるよう VB.NET で再実装した COM 公開クラス。
'*      レジストリから接続先ホスト・サービス名・タイムアウト値を読み込み、
'*      TCP/IP 経由で認証サーバへ VIDR コマンドを送信し、応答を解析する。
'*  [in]
'*      なし
'*  [out]
'*      なし
'*  [備考]
'*      ClassId / InterfaceId / EventsId は既存 OCX と互換性を保つため
'*      従来の GUID をそのまま使用する。
'***********************************************************************
<ComClass(BOJClientVBCLS.ClassId, BOJClientVBCLS.InterfaceId, BOJClientVBCLS.EventsId)>
Public Class BOJClientVBCLS

#Region "COM GUIDs"
    'クラスID（COM公開用）
    Public Const ClassId As String = "8C13C779-8032-11D2-9A81-00A024B9323E"
    'インターフェースID（COM公開用）
    Public Const InterfaceId As String = "8C13C777-8032-11D2-9A81-00A024B9323E"
    'イベントID（COM公開用）
    Public Const EventsId As String = "8C13C778-8032-11D2-9A81-00A024B9323E"
#End Region

    Private m_host1 As String = "hostname1"         '接続先ホスト名（1番目）
    Private m_host2 As String = "hostname2"         '接続先ホスト名（2番目：host1失敗時に使用）
    Private m_service As String = "servicename0"    'services ファイルから検索するサービス名
    Private m_nTimeout As Integer = 60000           'ソケットタイムアウト(秒) ※内部では ×1000 して使用
    Private m_bInited As Boolean = False            '初期化成否フラグ（True:初期化成功）

    'レジストリキーパス（通常用）
    Private Const REGISTRY_KEY_PATH As String = "SOFTWARE\sys148"
    'レジストリキーパス（プロファイル用 ※優先して参照）
    Private Const REGISTRY_KEY_PATH_PROFILES As String = "SOFTWARE\sys148\Profiles"

    '***********************************************************************
    '* New --- コンストラクタ
    '*  [概要]
    '*      インスタンス生成時にレジストリからホスト名等の設定値を読み込む。
    '*  [in]
    '*      なし
    '*  [out]
    '*      なし
    '*  [備考]
    '***********************************************************************
    Public Sub New()
        MyBase.New()
        '設定値をレジストリから取得
        InitializeFromRegistry()
    End Sub

    '***********************************************************************
    '* InitializeFromRegistry
    '*  [概要]
    '*      レジストリ(HKEY_LOCAL_MACHINE\SOFTWARE\sys148\Profiles または
    '*      SOFTWARE\sys148)から接続先ホスト・サービス名・タイムアウトを取得する。
    '*      Profiles キーが存在する場合はそちらを優先して読み込む。
    '*  [in]
    '*      なし
    '*  [out]
    '*      なし（メンバ変数 m_host1/m_host2/m_service/m_nTimeout/m_bInited を設定）
    '*  [備考]
    '*      例外発生時はメンバ変数の初期値を維持し、m_bInited を False に設定する。
    '***********************************************************************
    Private Sub InitializeFromRegistry()
        Try
            'まず Profiles サブキーを開く
            Using key As RegistryKey = Registry.LocalMachine.OpenSubKey(REGISTRY_KEY_PATH_PROFILES, False)
                If key Is Nothing Then
                    'Profiles が存在しない場合は、親キー(SOFTWARE\sys148)を参照
                    Using key2 As RegistryKey = Registry.LocalMachine.OpenSubKey(REGISTRY_KEY_PATH, False)
                        If key2 Is Nothing Then
                            '親キーも存在しない場合は初期化失敗
                            m_bInited = False
                            Return
                        End If
                        '親キーから値を読み込み
                        ReadRegistryValues(key2)
                    End Using
                Else
                    'Profiles キーから値を読み込み
                    ReadRegistryValues(key)
                End If
            End Using
            '正常終了時は初期化成功フラグON
            m_bInited = True
        Catch ex As Exception
            '例外時は初期化失敗フラグON
            m_bInited = False
        End Try
    End Sub

    '***********************************************************************
    '* ReadRegistryValues
    '*  [概要]
    '*      指定レジストリキーから Host01/Host02/ServiceName/SocketTimeOut の
    '*      値を読み込み、メンバ変数へ格納する。
    '*  [in]
    '*      key     RegistryKey     値を読み込むレジストリキー
    '*  [out]
    '*      なし（メンバ変数 m_host1/m_host2/m_service/m_nTimeout を設定）
    '*  [備考]
    '*      値が存在しない場合は既定値（コンストラクタ初期値）を維持する。
    '***********************************************************************
    Private Sub ReadRegistryValues(key As RegistryKey)
        '接続先ホスト名1
        Dim host01 As Object = key.GetValue("Host01")
        If host01 IsNot Nothing Then
            m_host1 = host01.ToString()
        End If

        '接続先ホスト名2
        Dim host02 As Object = key.GetValue("Host02")
        If host02 IsNot Nothing Then
            m_host2 = host02.ToString()
        End If

        'サービス名（services ファイル検索用）
        Dim serviceName As Object = key.GetValue("ServiceName")
        If serviceName IsNot Nothing Then
            m_service = serviceName.ToString()
        End If

        'ソケットタイムアウト(秒)
        Dim socketTimeout As Object = key.GetValue("SocketTimeOut")
        If socketTimeout IsNot Nothing Then
            m_nTimeout = Convert.ToInt32(socketTimeout)
        End If
    End Sub

    '***********************************************************************
    '* AAVerifyID --- ユーザID認証
    '*  [概要]
    '*      指定されたユーザID・パスワードを認証サーバへ送信し、
    '*      DBユーザID・DBパスワード・オンラインフラグを取得する。
    '*      host1 への接続に失敗した場合は host2 へフェイルオーバする。
    '*  [in]
    '*      szID        String      ユーザID
    '*      szPWD       String      パスワード
    '*  [out]
    '*      szID        String      （ゼロクリアして返却）
    '*      szPWD       String      （ゼロクリアして返却）
    '*      szDBID      String      DBユーザID
    '*      szDBPWD     String      DBパスワード
    '*      szON        String      オンラインフラグ
    '*      szRESULT    String      処理ステータス
    '*                                0:正常
    '*                                1:レジストリエラー
    '*                                2:コネクションエラー
    '*                                3:その他のエラー(szERRにエラー内容)
    '*                                4:受信タイムアウト
    '*      szERR       String      エラーメッセージ
    '*  [ret]
    '*      Integer     常に 0 を返す（処理結果は szRESULT で判定）
    '*  [備考]
    '*      既存 OCX インターフェースとの互換性を保つため引数・戻り値仕様は従来踏襲。
    '***********************************************************************
    Public Function AAVerifyID(
        ByRef szID As String,
        ByRef szPWD As String,
        ByRef szDBID As String,
        ByRef szDBPWD As String,
        ByRef szON As String,
        ByRef szRESULT As String,
        ByRef szERR As String
    ) As Integer

        Dim lStat As Integer = 0            '処理ステータス（エラー発生箇所の特定用）
        Dim sID As String = If(szID, "")    '入力ユーザID（Null 回避）
        Dim sPWD As String = If(szPWD, "")  '入力パスワード（Null 回避）
        Dim sDBUser As String = ""          'DBユーザID格納領域
        Dim sDBPass As String = ""          'DBパスワード格納領域
        Dim sOn As String = ""              'オンラインフラグ格納領域
        Dim sErr As String = ""             'エラーメッセージ格納領域

        '出力引数をゼロクリア
        szID = ""
        szPWD = ""
        szDBID = ""
        szDBPWD = ""
        szON = ""
        szRESULT = ""
        szERR = ""

        '--------------------------------------------------
        '初期化チェック
        '--------------------------------------------------
        lStat = 1
        If Not m_bInited Then
            'レジストリ読み込み失敗時は処理中断
            sErr = "OCX initialize failed."
            GoTo lblerr
        End If

        '--------------------------------------------------
        'サービスポート番号取得
        '--------------------------------------------------
        'services ファイルからサービス名に対応するポート番号を取得
        Dim port As Integer = GetServicePort(m_service)
        If port <= 0 Then
            '取得できなかった場合は既定値(50000)を使用
            port = 50000
        End If

        '--------------------------------------------------
        'サーバ接続（host1 → 失敗時 host2 フェイルオーバ）
        '--------------------------------------------------
        lStat = 2                       'コネクションエラー用ステータス
        Dim client As TcpClient = Nothing
        Dim nHost As Integer = 1    '接続成功したホスト番号（1:host1, 2:host2）

        Try
            'host1 へ接続試行
            client = New TcpClient()
            client.ReceiveTimeout = m_nTimeout * 1000
            client.SendTimeout = m_nTimeout * 1000
            client.Connect(m_host1, port)
        Catch ex As Exception
            'host1 接続失敗時は host2 へフェイルオーバ
            If client IsNot Nothing Then
                client.Close()
                client = Nothing
            End If
            nHost = 2
            Try
                'host2 へ接続試行
                client = New TcpClient()
                client.ReceiveTimeout = m_nTimeout * 1000
                client.SendTimeout = m_nTimeout * 1000
                client.Connect(m_host2, port)
            Catch ex2 As Exception
                'host2 も接続失敗時は処理中断
                If client IsNot Nothing Then
                    client.Close()
                    client = Nothing
                End If
                sErr = ex2.Message
                GoTo lblerr
            End Try
        End Try

        '--------------------------------------------------
        'VIDRコマンド送信
        '--------------------------------------------------
        lStat = 3
        Try
            Dim stream As NetworkStream = client.GetStream()
            'VIDRコマンド文字列生成（設計書仕様によりダミー文字列"DUMMY"固定）
            Dim command As String = BuildCmdVIDR("DUMMY", "DUMMY")
            '送信用バイト列へ変換（ASCIIエンコード）
            Dim sendData As Byte() = Encoding.ASCII.GetBytes(command)
            'サーバへ送信
            stream.Write(sendData, 0, sendData.Length)
        Catch ex As Exception
            '送信失敗時は処理中断
            sErr = "Send error: " & ex.Message
            client.Close()
            GoTo lblerr
        End Try

        '--------------------------------------------------
        '応答受信（タイムアウト付きポーリング）
        '--------------------------------------------------
        lStat = 4                       '受信タイムアウト用ステータス（例外発生時は 3 に切替）
        Dim response As String = ""     'サーバ応答文字列
        Try
            Dim stream As NetworkStream = client.GetStream()
            Dim buffer(1023) As Byte    '受信バッファ(1024バイト)
            Dim bytesRead As Integer = 0
            Dim totalWait As Integer = 0    '累計待機時間(ミリ秒)

            'タイムアウトまで 100ms 間隔でデータ到着を待機
            While totalWait < m_nTimeout * 1000
                If stream.DataAvailable Then
                    'データ到着時は読み込み
                    bytesRead = stream.Read(buffer, 0, buffer.Length)
                    If bytesRead > 0 Then
                        'ASCII 文字列へ変換しループ脱出
                        response = Encoding.ASCII.GetString(buffer, 0, bytesRead)
                        Exit While
                    End If
                Else
                    'データ未到着時は 100ms スリープして再試行
                    System.Threading.Thread.Sleep(100)
                    totalWait += 100
                End If
            End While

            If String.IsNullOrEmpty(response) Then
                'タイムアウト時は処理中断
                sErr = "Receive timeout"
                client.Close()
                GoTo lblerr
            End If
        Catch ex As Exception
            '受信中の例外はタイムアウトではなく「その他のエラー」として扱う
            lStat = 3
            sErr = "Receive error: " & ex.Message
            client.Close()
            GoTo lblerr
        End Try

        '--------------------------------------------------
        '応答解析
        '--------------------------------------------------
        lStat = 3                       '解析エラーは「その他のエラー」として扱う
        'ソケットクローズ
        client.Close()

        '応答長チェック
        If response.Length < 1 Then
            sErr = "Too short response"
            GoTo lblerr
        End If

        Dim pUser As String = ""        '解析結果：DBユーザID
        Dim pPass As String = ""        '解析結果：DBパスワード
        Dim nOnline As Integer = 0      '解析結果：オンラインフラグ

        'VIDR応答を解析
        If Not AnalyseResVIDR(response, pUser, pPass, nOnline) Then
            '解析失敗(NAK応答 等)時は処理中断
            sErr = "NG"
            GoTo lblerr
        End If

        '解析結果を出力引数へ格納
        sDBUser = pUser
        szDBID = sDBUser
        sOn = nOnline.ToString()
        szON = sOn
        sDBPass = pPass
        szDBPWD = sDBPass
        '正常終了
        lStat = 0

lblerr:
        '処理ステータスとエラーメッセージを出力引数へ格納
        szRESULT = lStat.ToString()
        szERR = sErr
        Return 0
    End Function

    '***********************************************************************
    '* BuildCmdVIDR
    '*  [概要]
    '*      VIDR コマンド文字列を組み立てる。
    '*      書式: "VIDR <user>, <pass><LF>"
    '*  [in]
    '*      user    String      ユーザID
    '*      pass    String      パスワード
    '*  [out]
    '*      なし
    '*  [ret]
    '*      String  組み立てた VIDR コマンド文字列
    '*  [備考]
    '***********************************************************************
    Private Function BuildCmdVIDR(user As String, pass As String) As String
        Return String.Format("VIDR {0}, {1}" & vbLf, user, pass)
    End Function

    '***********************************************************************
    '* AnalyseResVIDR
    '*  [概要]
    '*      VIDR コマンドに対するサーバ応答文字列を解析する。
    '*      応答書式: "ACK <user>, <pass>, <online><LF>"
    '*  [in]
    '*      buf     String      サーバ応答文字列
    '*  [out]
    '*      user    String      DBユーザID
    '*      pass    String      DBパスワード
    '*      online  Integer     オンラインフラグ
    '*  [ret]
    '*      Boolean     True:解析成功  False:解析失敗(ACKでない/項目不足)
    '*  [備考]
    '***********************************************************************
    Private Function AnalyseResVIDR(buf As String, ByRef user As String, ByRef pass As String, ByRef online As Integer) As Boolean
        '出力引数を初期化
        user = ""
        pass = ""
        online = 0

        '応答先頭が "ACK" でない場合は解析失敗
        If Not buf.StartsWith("ACK") Then
            Return False
        End If

        '"ACK" を除去し、カンマ区切りで分解
        Dim content As String = buf.Substring(3).TrimStart()
        Dim parts As String() = content.Split(","c)

        '項目数チェック(user, pass, online の 3 要素必須)
        If parts.Length < 3 Then
            Return False
        End If

        'user / pass を取り出し
        user = parts(0).Trim()
        pass = parts(1).Trim()

        'オンラインフラグは LF までの部分を抽出し整数化
        Dim onlineStr As String = parts(2).Trim()
        Dim nlPos As Integer = onlineStr.IndexOf(vbLf)
        If nlPos >= 0 Then
            onlineStr = onlineStr.Substring(0, nlPos)
        End If
        Integer.TryParse(onlineStr, online)

        Return True
    End Function

    '***********************************************************************
    '* GetServicePort
    '*  [概要]
    '*      Windows の services ファイル(%SystemRoot%\System32\drivers\etc\services)
    '*      から、指定されたサービス名に対応するポート番号を取得する。
    '*  [in]
    '*      serviceName     String      検索するサービス名
    '*  [out]
    '*      なし
    '*  [ret]
    '*      Integer     ポート番号。取得できない場合は -1。
    '*  [備考]
    '*      行頭が "#" の行(コメント行)・空行はスキップする。
    '*      サービス名の比較は大文字小文字を区別しない。
    '***********************************************************************
    Private Function GetServicePort(serviceName As String) As Integer
        Try
            'services ファイルのフルパスを組み立て
            Dim servicesPath As String = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.System),
                "drivers\etc\services")

            If System.IO.File.Exists(servicesPath) Then
                '全行読み込み
                Dim lines As String() = System.IO.File.ReadAllLines(servicesPath)
                For Each line As String In lines
                    '空行・コメント行はスキップ
                    If String.IsNullOrWhiteSpace(line) OrElse line.StartsWith("#") Then
                        Continue For
                    End If

                    '空白・タブ区切りで分解
                    Dim parts As String() = line.Split(New Char() {" "c, vbTab(0)}, StringSplitOptions.RemoveEmptyEntries)
                    'サービス名一致チェック（大文字小文字無視）
                    If parts.Length >= 2 AndAlso parts(0).Equals(serviceName, StringComparison.OrdinalIgnoreCase) Then
                        '"ポート/プロトコル" 形式を "/" で分解
                        Dim portProto As String() = parts(1).Split("/"c)
                        If portProto.Length >= 1 Then
                            Dim port As Integer
                            'ポート番号の数値変換に成功したら返却
                            If Integer.TryParse(portProto(0), port) Then
                                Return port
                            End If
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            '例外は無視して -1 を返す
        End Try

        'サービス名未登録
        Return -1
    End Function

End Class
